"""Daily monitor for the MDOT Engineering Standards/Guides/Manuals page."""

from __future__ import annotations

import argparse
import base64
import copy
import difflib
import hashlib
import html as html_module
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from typing import BinaryIO, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urljoin, urlsplit, urlunsplit
from urllib.request import Request, urlopen
from xml.etree import ElementTree


APP_NAME = "MDOTStandardsMonitor"
DEFAULT_URL = "https://mdot.ms.gov/portal/engineering_standards_guides_manuals"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/151 Safari/537.36 MDOTStandardsMonitor/1.0"
)
DOCUMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm",
    ".txt", ".csv", ".zip", ".dgn", ".dwg",
}
BASE_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / APP_NAME
CONFIG_FILE = BASE_DIR / "config.json"
STATE_FILE = BASE_DIR / "state.json"
RUNTIME_FILE = BASE_DIR / "runtime.json"
HISTORY_FILE = BASE_DIR / "history.json"
HISTORY_DASHBOARD = BASE_DIR / "history.html"
HISTORY_REPORTS_DIR = BASE_DIR / "history_reports"
PREVIEW_DIR = BASE_DIR / "page_previews"
LOCK_FILE = BASE_DIR / "monitor.lock"
LOG_FILE = BASE_DIR / "monitor.log"
SCRIPT_DIR = Path(__file__).resolve().parent
EMAIL_HELPER = SCRIPT_DIR / "send_outlook.ps1"
TASK_NAME = "MDOT Standards Daily Monitor"
DEFAULT_RETRY_ATTEMPTS = 3
DEFAULT_RETRY_DELAY_SECONDS = 5
DEFAULT_CONFIRMATION_DELAY_SECONDS = 120
DEFAULT_HEARTBEAT_DAYS = 7
DEFAULT_HISTORY_LIMIT = 100
MAX_PDF_PREVIEW_PAIRS = 3
SNAPSHOT_SCHEMA_VERSION = 3
EXCLUDED_TOP_CATEGORIES = {"Construction", "Materials"}


class MonitorError(RuntimeError):
    """An expected monitoring failure."""


class MainContentParser(HTMLParser):
    """Extract normalized main-page text, headings, and links."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_main = False
        self.suppressed_depth = 0
        self.generated_inventory_depth = 0
        self.text_parts: list[str] = []
        self.links: list[dict[str, str]] = []
        self.folders: list[str] = []
        self.current_link: dict[str, object] | None = None
        self.current_heading: list[str] | None = None
        self.section = ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        attr_map = dict(attrs)
        folder_path = normalize_text(attr_map.get("data-mdot-folder-path") or "")
        if tag == "tr" and folder_path and folder_path not in self.folders:
            self.folders.append(folder_path)
        if tag == "main" and not self.in_main:
            self.in_main = True
            return
        if not self.in_main:
            return
        if self.generated_inventory_depth:
            self.generated_inventory_depth += 1
        elif attr_map.get("data-mdot-monitor-inventory") == "true":
            # render_page adds this table solely to capture links and folders from
            # MDOT's virtualized grids. Its text is not visible page copy.
            self.generated_inventory_depth = 1
        if tag in {"script", "style", "svg", "noscript"}:
            self.suppressed_depth += 1
        if self.suppressed_depth:
            return
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            self.current_heading = []
        if tag == "a" and attr_map.get("href"):
            self.current_link = {
                "href": attr_map["href"] or "",
                "text": [],
                "section": folder_path or self.section,
            }

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if not self.in_main:
            return
        if tag == "main":
            self.in_main = False
            return
        if self.generated_inventory_depth:
            self.generated_inventory_depth -= 1
        if tag in {"script", "style", "svg", "noscript"} and self.suppressed_depth:
            self.suppressed_depth -= 1
        if not self.suppressed_depth:
            if tag == "a" and self.current_link is not None:
                link_text = normalize_text(" ".join(self.current_link["text"]))
                self.links.append(
                    {
                        "url": str(self.current_link["href"]),
                        "title": link_text,
                        "section": str(self.current_link.get("section", self.section)),
                    }
                )
                self.current_link = None
            if tag in {"h1", "h2", "h3", "h4", "h5", "h6"} and self.current_heading is not None:
                heading = normalize_text(" ".join(self.current_heading))
                if heading:
                    self.section = heading
                self.current_heading = None

    def handle_data(self, data: str) -> None:
        if not self.in_main or self.suppressed_depth:
            return
        text = normalize_text(data)
        if not text:
            return
        if not self.generated_inventory_depth:
            self.text_parts.append(text)
        if self.current_heading is not None and not self.generated_inventory_depth:
            self.current_heading.append(text)
        if self.current_link is not None:
            self.current_link["text"].append(text)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\u200c", " ").replace("\xa0", " ")).strip()


def retry_operation(operation, attempts: int, delay_seconds: float, logger: logging.Logger, label: str):
    """Retry one network/render operation without hiding the final error."""
    attempts = max(1, int(attempts))
    delay_seconds = max(0.0, float(delay_seconds))
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as exc:
            if attempt >= attempts:
                raise
            wait = delay_seconds * (2 ** (attempt - 1))
            logger.warning(
                "%s failed on attempt %d/%d (%s); retrying in %.1f seconds",
                label,
                attempt,
                attempts,
                normalize_text(str(exc)) or type(exc).__name__,
                wait,
            )
            time.sleep(wait)


def canonicalize_url(base_url: str, value: str) -> str:
    absolute = urljoin(base_url, value.strip())
    parts = urlsplit(absolute)
    scheme = parts.scheme.lower()
    hostname = (parts.hostname or "").lower()
    port = f":{parts.port}" if parts.port else ""
    encoded_path = quote(parts.path, safe="/%:@!$&'()*+,;=-._~")
    encoded_query = quote(parts.query, safe="=&?/:;+,%@!$'()*-._~")
    return urlunsplit((scheme, hostname + port, encoded_path, encoded_query, ""))


def is_mdot_document(url: str) -> bool:
    parts = urlsplit(url)
    suffix = Path(unquote(parts.path)).suffix.lower()
    return parts.hostname in {"mdot.ms.gov", "www.mdot.ms.gov"} and (
        parts.path.lower().startswith("/documents/") or suffix in DOCUMENT_EXTENSIONS
    )


def is_excluded_document(item: dict[str, object]) -> bool:
    """Exclude the two top-level website sections the team does not monitor."""
    section = str(item.get("section", "")).split(" / ", 1)[0].casefold()
    if section in {value.casefold() for value in EXCLUDED_TOP_CATEGORIES}:
        return True
    path = unquote(urlsplit(str(item.get("url", ""))).path).replace("\\", "/").casefold()
    return path.startswith("/documents/construction/") or path.startswith("/documents/materials/")


def parse_rendered_page(rendered_html: str, page_url: str) -> dict[str, object]:
    parser = MainContentParser()
    parser.feed(rendered_html)
    parser.close()
    if not parser.text_parts:
        raise MonitorError("Chrome returned a page without readable <main> content")

    links: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for item in parser.links:
        url = canonicalize_url(page_url, item["url"])
        if urlsplit(url).scheme not in {"http", "https"}:
            continue
        title = item["title"] or Path(unquote(urlsplit(url).path)).name
        normalized = (url, normalize_text(title), normalize_text(item["section"]))
        if normalized in seen:
            continue
        seen.add(normalized)
        links.append({"url": normalized[0], "title": normalized[1], "section": normalized[2]})

    links.sort(key=lambda item: (item["url"].casefold(), item["title"].casefold()))
    folders = [
        {
            "path": path,
            "title": path.rsplit(" / ", 1)[-1],
            "section": path,
            "url": page_url,
        }
        for path in sorted(parser.folders, key=str.casefold)
    ]
    return {"text": normalize_text(" ".join(parser.text_parts)), "links": links, "folders": folders}


def find_chrome() -> Path:
    candidates = [
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    discovered = shutil.which("chrome") or shutil.which("chrome.exe")
    if discovered:
        return Path(discovered)
    raise MonitorError("Google Chrome was not found")


def render_page(url: str, timeout_seconds: int = 90) -> str:
    chrome = find_chrome()
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise MonitorError(
            "Playwright is required to expand the MDOT folder tree. Re-run install.ps1 to install dependencies."
        ) from exc

    timeout_ms = max(1, int(timeout_seconds * 1000))
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(chrome),
                headless=True,
                args=[
                    "--disable-gpu",
                    "--disable-extensions",
                    "--disable-background-networking",
                    "--no-first-run",
                    "--no-default-browser-check",
                ],
            )
            try:
                context = browser.new_context(user_agent=USER_AGENT)
                page = context.new_page()
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                if urlsplit(url).hostname not in {"mdot.ms.gov", "www.mdot.ms.gov"}:
                    page.wait_for_selector("main", timeout=timeout_ms)
                    page.wait_for_timeout(250)
                    return page.content()
                page.wait_for_selector(
                    "main .dx-datagrid tr.dx-group-row",
                    state="attached",
                    timeout=timeout_ms,
                )
                page.wait_for_timeout(750)
                folders: set[str] = set()
                folder_links: dict[tuple[str, str, str], dict[str, str]] = {}
                included_top_categories: set[str] = set()
                grid_count = page.locator("main .dx-datagrid").count()
                for grid_index in range(grid_count):
                    top_category = page.evaluate(
                        """(gridIndex) => {
                            const rowLevel = (row) => {
                                const cell = row.querySelector('td[aria-label="Expand"], td[aria-label="Collapse"]');
                                return Math.max(0, Number(cell?.getAttribute('aria-colindex') || 1) - 1);
                            };
                            const grid = document.querySelectorAll('main .dx-datagrid')[gridIndex];
                            if (!grid) return '';
                            for (const row of grid.querySelectorAll('tr.dx-group-row')) {
                                if (rowLevel(row) === 0 && (row.innerText || '').trim()) {
                                    return (row.innerText || '').trim();
                                }
                            }
                            return '';
                        }""",
                        grid_index,
                    )
                    if top_category in EXCLUDED_TOP_CATEGORIES:
                        continue
                    if top_category:
                        included_top_categories.add(top_category)

                    visited_pages: set[str] = set()
                    current_path: list[str] = []
                    while True:
                        for _ in range(500):
                            expanded_path = page.evaluate(
                                """(gridIndex) => {
                                    const rowLevel = (row) => {
                                        const cell = row.querySelector('td[aria-label="Expand"], td[aria-label="Collapse"]');
                                        return Math.max(0, Number(cell?.getAttribute('aria-colindex') || 1) - 1);
                                    };
                                    const grid = document.querySelectorAll('main .dx-datagrid')[gridIndex];
                                    if (!grid) return null;
                                    const path = [];
                                    for (const row of grid.querySelectorAll('tr.dx-group-row')) {
                                        const name = (row.innerText || '').trim();
                                        if (!name) continue;
                                        const level = rowLevel(row);
                                        path.length = level;
                                        path[level] = name;
                                        if (row.getAttribute('aria-expanded') === 'false') {
                                            (row.querySelector('p') || row).click();
                                            return path.join(' / ');
                                        }
                                    }
                                    return null;
                                }""",
                                grid_index,
                            )
                            if expanded_path is None:
                                break
                            page.wait_for_timeout(60)
                        else:
                            raise MonitorError("A single MDOT grid exceeded the safe 500-folder expansion limit")

                        page_inventory = page.evaluate(
                            """({gridIndex, initialPath}) => {
                                const rowLevel = (row) => {
                                    const cell = row.querySelector('td[aria-label="Expand"], td[aria-label="Collapse"]');
                                    return Math.max(0, Number(cell?.getAttribute('aria-colindex') || 1) - 1);
                                };
                                const grid = document.querySelectorAll('main .dx-datagrid')[gridIndex];
                                const path = [...initialPath];
                                const folders = [];
                                const links = [];
                                if (!grid) return {folders, links, endPath:path, currentPage:'1', pages:['1']};
                                for (const row of grid.querySelectorAll('tr')) {
                                    if (row.classList.contains('dx-group-row')) {
                                        const name = (row.innerText || '').trim();
                                        if (!name) continue;
                                        const level = rowLevel(row);
                                        path.length = level;
                                        path[level] = name;
                                        folders.push(path.join(' / '));
                                    } else if (row.classList.contains('dx-data-row')) {
                                        const folderPath = path.join(' / ');
                                        for (const anchor of row.querySelectorAll('a[href]')) {
                                            links.push({
                                                url: anchor.href,
                                                title: (anchor.innerText || '').trim(),
                                                section: folderPath,
                                            });
                                        }
                                    }
                                }
                                const pageButtons = [...grid.querySelectorAll('.dx-page[aria-label^="Page "]')];
                                const selected = pageButtons.find(button => button.getAttribute('aria-current') === 'page');
                                return {
                                    folders,
                                    links,
                                    endPath:path,
                                    currentPage:selected ? (selected.innerText || '').trim() : '1',
                                    pages:pageButtons.map(button => (button.innerText || '').trim()).filter(Boolean),
                                };
                            }""",
                            {"gridIndex": grid_index, "initialPath": current_path},
                        )
                        current_path = list(page_inventory["endPath"])
                        folders.update(path for path in page_inventory["folders"] if path)
                        for item in page_inventory["links"]:
                            key = (item["url"], item["title"], item["section"])
                            folder_links[key] = item
                        visited_pages.add(str(page_inventory["currentPage"]))
                        next_page = next(
                            (str(value) for value in page_inventory["pages"] if str(value) not in visited_pages),
                            None,
                        )
                        if next_page is None:
                            break
                        grid = page.locator("main .dx-datagrid").nth(grid_index)
                        grid.locator(f'.dx-page[aria-label="Page {next_page}"]').click()
                        page.wait_for_timeout(150)

                valid_folders = {
                    path
                    for path in folders
                    if path.split(" / ", 1)[0] in included_top_categories
                    and all(part.strip() for part in path.split(" / "))
                }
                normalized_folder_links: dict[tuple[str, str, str], dict[str, str]] = {}
                for item in folder_links.values():
                    parts = urlsplit(item["url"])
                    path_parts = [part for part in unquote(parts.path).split("/") if part]
                    if (
                        parts.hostname in {"mdot.ms.gov", "www.mdot.ms.gov"}
                        and len(path_parts) >= 3
                        and path_parts[0].casefold() == "documents"
                    ):
                        folder_parts = path_parts[1:-1]
                        item["section"] = " / ".join(folder_parts)
                        for length in range(1, len(folder_parts) + 1):
                            valid_folders.add(" / ".join(folder_parts[:length]))
                    key = (item["url"], item["title"], item["section"])
                    normalized_folder_links[key] = item
                inventory = {
                    "folders": sorted(valid_folders, key=str.casefold),
                    "links": list(normalized_folder_links.values()),
                }
                page.evaluate(
                    """(inventory) => {
                        const main = document.querySelector('main');
                        if (!main) return;
                        const table = document.createElement('table');
                        table.setAttribute('data-mdot-monitor-inventory', 'true');
                        const body = document.createElement('tbody');
                        table.appendChild(body);
                        for (const path of inventory.folders) {
                            const row = document.createElement('tr');
                            row.setAttribute('data-mdot-folder-path', path);
                            const cell = document.createElement('td');
                            cell.textContent = path;
                            row.appendChild(cell);
                            body.appendChild(row);
                        }
                        for (const item of inventory.links) {
                            const row = document.createElement('tr');
                            row.className = 'dx-data-row';
                            const cell = document.createElement('td');
                            const anchor = document.createElement('a');
                            anchor.href = item.url;
                            anchor.textContent = item.title;
                            anchor.setAttribute('data-mdot-folder-path', item.section);
                            cell.appendChild(anchor);
                            row.appendChild(cell);
                            body.appendChild(row);
                        }
                        for (const grid of [...main.querySelectorAll('.dx-datagrid')]) grid.remove();
                        main.appendChild(table);
                    }""",
                    inventory,
                )
                rendered_html = page.content()
                if not inventory["folders"] or not inventory["links"]:
                    raise MonitorError("Chrome rendered the MDOT page but did not expose its folder documents")
                return rendered_html
            finally:
                browser.close()
    except PlaywrightTimeoutError as exc:
        raise MonitorError(f"Chrome did not finish rendering within {timeout_seconds} seconds") from exc
    except MonitorError:
        raise
    except Exception as exc:
        raise MonitorError(f"Chrome could not render and expand the MDOT folder tree: {exc}") from exc


def hash_stream(stream: BinaryIO, chunk_size: int = 1024 * 1024) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    while True:
        chunk = stream.read(chunk_size)
        if not chunk:
            break
        digest.update(chunk)
        size += len(chunk)
    return digest.hexdigest(), size


def download_document(item: dict[str, str], timeout_seconds: int = 180) -> dict[str, object]:
    request = Request(item["url"], headers={"User-Agent": USER_AGENT, "Accept": "*/*"})
    suffix = Path(unquote(urlsplit(item["url"]).path)).suffix.lower() or ".download"
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("wb", suffix=suffix, delete=False) as temp_file:
            temp_path = Path(temp_file.name)
            with urlopen(request, timeout=timeout_seconds) as response:
                content_type = response.headers.get_content_type()
                expected_extension = Path(unquote(urlsplit(item["url"]).path)).suffix.lower()
                if expected_extension in DOCUMENT_EXTENSIONS and content_type.lower() in {
                    "text/html", "application/xhtml+xml",
                }:
                    raise MonitorError(
                        f"Expected a {expected_extension} file but the server returned {content_type} instead"
                    )
                digest = hashlib.sha256()
                size = 0
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    digest.update(chunk)
                    temp_file.write(chunk)
                    size += len(chunk)
            return {
                "url": item["url"],
                "title": item["title"],
                "section": item["section"],
                "identity": item.get("identity", item["url"]),
                "sha256": digest.hexdigest(),
                "size": size,
                "content_type": content_type,
                "content_disposition": response.headers.get("Content-Disposition", ""),
                "etag": response.headers.get("ETag", ""),
                "last_modified": response.headers.get("Last-Modified", ""),
                "_temp_path": str(temp_path),
            }
    except (HTTPError, URLError, TimeoutError, OSError, MonitorError) as exc:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        raise MonitorError(f"Could not download {item['title']} ({item['url']}): {exc}") from exc


def normalized_lines(value: str) -> list[str]:
    return [normalize_text(line) for line in value.splitlines() if normalize_text(line)]


def analyze_pdf(path: Path, preview_directory: Path | None = None) -> dict[str, object]:
    try:
        import pymupdf
    except ImportError as exc:
        raise MonitorError("PyMuPDF is required for page-level PDF comparisons") from exc

    pages: list[dict[str, object]] = []
    try:
        if preview_directory is not None:
            preview_directory.mkdir(parents=True, exist_ok=True)
        with pymupdf.open(path, filetype="pdf") as document:
            for index, page in enumerate(document):
                text = "\n".join(normalized_lines(page.get_text("text", sort=True)))
                # A low-resolution grayscale rendering catches drawings and scanned pages
                # that contain little or no extractable text.
                pixmap = page.get_pixmap(dpi=48, colorspace=pymupdf.csGRAY, alpha=False, annots=True)
                page_snapshot = {
                    "page": index + 1,
                    "text": text,
                    "text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
                    "visual_sha256": hashlib.sha256(pixmap.samples).hexdigest(),
                }
                if preview_directory is not None:
                    preview_path = preview_directory / f"page-{index + 1:04d}.png"
                    pixmap.save(preview_path)
                    page_snapshot["preview_path"] = str(preview_path)
                pages.append(page_snapshot)
    except Exception as exc:
        raise MonitorError(f"Could not analyze PDF {path.name}: {exc}") from exc
    return {"kind": "pdf_pages", "pages": pages}


def analyze_word(path: Path) -> dict[str, object]:
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        with zipfile.ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("word/document.xml"))
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise MonitorError(f"Could not analyze Word document {path.name}: {exc}") from exc
    paragraphs: list[str] = []
    for paragraph in root.iter(namespace + "p"):
        text = normalize_text("".join(node.text or "" for node in paragraph.iter(namespace + "t")))
        if text:
            paragraphs.append(text)
    return {"kind": "paragraphs", "paragraphs": paragraphs}


def analyze_excel(path: Path) -> dict[str, object]:
    cells: list[dict[str, str]] = []
    spreadsheet_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    relationship_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    package_relationship_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    try:
        with zipfile.ZipFile(path) as archive:
            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in archive.namelist():
                shared_root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
                for item in shared_root.iter(spreadsheet_ns + "si"):
                    shared_strings.append(
                        normalize_text("".join(node.text or "" for node in item.iter(spreadsheet_ns + "t")))
                    )

            relationship_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relationships = {
                relation.attrib["Id"]: relation.attrib["Target"]
                for relation in relationship_root.iter(package_relationship_ns + "Relationship")
            }
            workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            for sheet in workbook_root.iter(spreadsheet_ns + "sheet"):
                sheet_name = sheet.attrib.get("name", "Worksheet")
                relationship_id = sheet.attrib.get(relationship_ns + "id", "")
                target = relationships.get(relationship_id, "")
                if not target:
                    continue
                worksheet_path = target.lstrip("/")
                if not worksheet_path.startswith("xl/"):
                    worksheet_path = "xl/" + worksheet_path
                worksheet_root = ElementTree.fromstring(archive.read(worksheet_path))
                for cell in worksheet_root.iter(spreadsheet_ns + "c"):
                    coordinate = cell.attrib.get("r", "")
                    cell_type = cell.attrib.get("t", "")
                    formula = cell.find(spreadsheet_ns + "f")
                    value_node = cell.find(spreadsheet_ns + "v")
                    if formula is not None and formula.text is not None:
                        value = "=" + formula.text
                    elif cell_type == "inlineStr":
                        value = "".join(node.text or "" for node in cell.iter(spreadsheet_ns + "t"))
                    elif value_node is None or value_node.text is None:
                        continue
                    elif cell_type == "s":
                        try:
                            value = shared_strings[int(value_node.text)]
                        except (ValueError, IndexError):
                            value = value_node.text
                    else:
                        value = value_node.text
                    cells.append(
                        {
                            "sheet": sheet_name,
                            "cell": coordinate,
                            "value": normalize_text(str(value)),
                        }
                    )
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise MonitorError(f"Could not analyze Excel workbook {path.name}: {exc}") from exc
    return {"kind": "excel_cells", "cells": cells}


def analyze_text_file(path: Path) -> dict[str, object]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    return {"kind": "lines", "lines": normalized_lines(text)}


def infer_document_extension(
    url: str,
    content_type: str = "",
    content_disposition: str = "",
) -> str:
    extension = Path(unquote(urlsplit(url).path)).suffix.lower()
    if extension in DOCUMENT_EXTENSIONS:
        return extension
    filename_match = re.search(
        r"filename\*?=(?:UTF-8''|\")?([^\";]+)",
        content_disposition,
        flags=re.IGNORECASE,
    )
    if filename_match:
        disposition_extension = Path(unquote(filename_match.group(1).strip())).suffix.lower()
        if disposition_extension:
            return disposition_extension
    return {
        "application/pdf": ".pdf",
        "text/plain": ".txt",
        "text/csv": ".csv",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsm",
    }.get(content_type.lower(), extension)


def analyze_document(
    path: Path,
    url: str,
    content_type: str = "",
    content_disposition: str = "",
    preview_directory: Path | None = None,
) -> dict[str, object]:
    extension = infer_document_extension(url, content_type, content_disposition)
    if extension == ".pdf":
        return analyze_pdf(path, preview_directory)
    if extension in {".docx", ".docm"}:
        return analyze_word(path)
    if extension in {".xlsx", ".xlsm"}:
        return analyze_excel(path)
    if extension in {".txt", ".csv"}:
        return analyze_text_file(path)
    return {
        "kind": "unsupported",
        "reason": f"Detailed comparison is not available for {extension or 'this file type'}",
    }


def resolve_dynamic_document(rule: dict[str, str], timeout_seconds: int = 60) -> dict[str, str]:
    source_url = canonicalize_url(DEFAULT_URL, str(rule["source_url"]))
    request = Request(source_url, headers={"User-Agent": USER_AGENT, "Accept": "application/pdf,*/*"})
    try:
        with urlopen(request, timeout=timeout_seconds) as response:
            content = response.read(25 * 1024 * 1024 + 1)
    except (HTTPError, URLError, TimeoutError, OSError) as exc:
        raise MonitorError(f"Could not download dynamic-link source {source_url}: {exc}") from exc
    if len(content) > 25 * 1024 * 1024:
        raise MonitorError(f"Dynamic-link source is unexpectedly larger than 25 MB: {source_url}")
    try:
        import pymupdf
        with pymupdf.open(stream=content, filetype="pdf") as document:
            links = {
                canonicalize_url(source_url, str(link["uri"]))
                for page in document
                for link in page.get_links()
                if link.get("uri")
            }
    except Exception as exc:
        raise MonitorError(f"Could not extract links from dynamic-link source {source_url}: {exc}") from exc

    match_host = str(rule["match_host"]).lower()
    match_path = str(rule.get("match_path_contains", ""))
    matches = sorted(
        link
        for link in links
        if urlsplit(link).hostname == match_host
        and (not match_path or match_path.lower() in urlsplit(link).path.lower())
    )
    if len(matches) != 1:
        raise MonitorError(
            f"Expected exactly one matching link in {source_url}, found {len(matches)}"
        )
    return {
        "url": matches[0],
        "title": normalize_text(str(rule["title"])),
        "section": normalize_text(str(rule.get("section", "Dynamically monitored documents"))),
        "identity": "dynamic:" + normalize_text(str(rule["title"])),
        "source_url": source_url,
    }


def normalize_filters(filters: dict[str, object] | None) -> dict[str, list[str]]:
    source = filters or {}
    return {
        "sections": [normalize_text(str(value)) for value in source.get("sections", []) if normalize_text(str(value))],
        "titles": [normalize_text(str(value)) for value in source.get("titles", []) if normalize_text(str(value))],
        "extensions": [
            ("." + str(value).lstrip(".")).lower()
            for value in source.get("extensions", [])
            if str(value).strip()
        ],
    }


def item_matches_filters(item: dict[str, object], filters: dict[str, object] | None) -> bool:
    normalized = normalize_filters(filters)
    section = str(item.get("section", "")).casefold()
    title = str(item.get("title", item.get("new_title", ""))).casefold()
    path = unquote(urlsplit(str(item.get("url", ""))).path)
    extension = Path(path).suffix.lower()
    if normalized["sections"] and not any(value.casefold() in section for value in normalized["sections"]):
        return False
    if normalized["titles"] and not any(value.casefold() in title for value in normalized["titles"]):
        return False
    if normalized["extensions"] and extension not in normalized["extensions"]:
        return False
    return True


def analysis_can_be_reused(analysis: dict[str, object], store_previews: bool) -> bool:
    if not analysis:
        return False
    if not store_previews or analysis.get("kind") != "pdf_pages":
        return True
    pages = list(analysis.get("pages", []))
    return bool(pages) and all(
        page.get("preview_path") and Path(str(page["preview_path"])).is_file()
        for page in pages
    )


def build_snapshot(
    page_url: str,
    logger: logging.Logger,
    previous_snapshot: dict[str, object] | None = None,
    force_analysis: bool = False,
    extra_documents: list[dict[str, str]] | None = None,
    dynamic_documents: list[dict[str, str]] | None = None,
    filters: dict[str, object] | None = None,
    store_previews: bool = True,
    preview_root: Path | None = None,
    retry_attempts: int = 1,
    retry_delay_seconds: float = 0.0,
) -> dict[str, object]:
    rendered = retry_operation(
        lambda: render_page(page_url),
        retry_attempts,
        retry_delay_seconds,
        logger,
        "MDOT page render",
    )
    page = parse_rendered_page(rendered, page_url)
    document_by_url: dict[str, dict[str, str]] = {}
    for item in page["links"]:
        if not is_mdot_document(item["url"]) or is_excluded_document(item):
            continue
        current = document_by_url.get(item["url"])
        if current is None or (not current.get("section") and item.get("section")):
            document_by_url[item["url"]] = item
    document_links = list(document_by_url.values())
    known_urls = {item["url"] for item in document_links}
    previous_documents = {
        item.get("identity", item["url"]): item
        for item in (previous_snapshot or {}).get("documents", [])
    }
    previous_by_url = {
        item["url"]: item for item in (previous_snapshot or {}).get("documents", [])
    }
    document_errors: list[dict[str, str]] = []
    carried_documents: list[dict[str, object]] = []
    for extra in extra_documents or []:
        extra_url = canonicalize_url(page_url, str(extra["url"]))
        if extra_url not in known_urls:
            document_links.append(
                {
                    "url": extra_url,
                    "title": normalize_text(str(extra["title"])),
                    "section": normalize_text(str(extra.get("section", "Explicitly monitored documents"))),
                }
            )
            known_urls.add(extra_url)
    for rule in dynamic_documents or []:
        dynamic_title = normalize_text(str(rule.get("title", "dynamic document")))
        try:
            resolved = retry_operation(
                lambda rule=rule: resolve_dynamic_document(rule),
                retry_attempts,
                retry_delay_seconds,
                logger,
                f"Dynamic document resolution: {dynamic_title}",
            )
        except Exception as exc:
            error = normalize_text(str(exc)) or type(exc).__name__
            logger.error("Could not resolve %s after all retries; continuing: %s", dynamic_title, error)
            document_errors.append({"title": dynamic_title, "url": str(rule.get("source_url", "")), "error": error})
            previous = previous_documents.get("dynamic:" + dynamic_title)
            if previous and item_matches_filters(previous, filters):
                carried_documents.append(dict(previous))
            continue
        if resolved["url"] not in known_urls:
            document_links.append(resolved)
            known_urls.add(resolved["url"])
            logger.info("Resolved dynamic document %s to %s", resolved["title"], resolved["url"])
    document_links = [item for item in document_links if item_matches_filters(item, filters)]
    filtered_page_links = [
        item for item in page["links"]
        if item_matches_filters(item, {"sections": normalize_filters(filters)["sections"], "titles": normalize_filters(filters)["titles"]})
    ]
    logger.info("Rendered page with %d links and %d monitored documents", len(page["links"]), len(document_links))
    documents = carried_documents
    for index, item in enumerate(document_links, start=1):
        logger.info("Hashing document %d/%d: %s", index, len(document_links), item["title"])
        previous = previous_documents.get(item.get("identity", item["url"])) or previous_by_url.get(item["url"])
        try:
            downloaded = retry_operation(
                lambda item=item: download_document(item),
                retry_attempts,
                retry_delay_seconds,
                logger,
                f"Document {index}/{len(document_links)}: {item['title']}",
            )
        except Exception as exc:
            error = normalize_text(str(exc)) or type(exc).__name__
            logger.error("Could not hash %s after all retries; continuing: %s", item["title"], error)
            document_errors.append({"title": item["title"], "url": item["url"], "error": error})
            if previous:
                documents.append(dict(previous))
            continue
        temp_path = Path(str(downloaded.pop("_temp_path")))
        try:
            if (
                not force_analysis
                and previous
                and previous.get("sha256") == downloaded["sha256"]
                and analysis_can_be_reused(previous.get("analysis") or {}, store_previews)
            ):
                downloaded["analysis"] = previous["analysis"]
            else:
                logger.info("Creating detailed content snapshot: %s", item["title"])
                downloaded["analysis"] = analyze_document(
                    temp_path,
                    item["url"],
                    str(downloaded.get("content_type", "")),
                    str(downloaded.get("content_disposition", "")),
                    (preview_root or PREVIEW_DIR) / str(downloaded["sha256"]) if store_previews else None,
                )
            documents.append(downloaded)
        finally:
            temp_path.unlink(missing_ok=True)
    documents.sort(key=lambda item: item["url"].casefold())
    return {
        "schema_version": SNAPSHOT_SCHEMA_VERSION,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "page_url": page_url,
        "page_text": page["text"],
        "folders": page.get("folders", []),
        "links": filtered_page_links,
        "documents": documents,
        "document_errors": document_errors,
    }


def compact_excerpt(old_lines: list[str], new_lines: list[str], limit: int = 700) -> str:
    removed: list[str] = []
    added: list[str] = []
    for line in difflib.ndiff(old_lines, new_lines):
        if line.startswith("- ") and len(removed) < 3:
            removed.append(line[2:])
        elif line.startswith("+ ") and len(added) < 3:
            added.append(line[2:])
    parts = []
    if removed:
        parts.append("Removed: " + " | ".join(removed))
    if added:
        parts.append("Added: " + " | ".join(added))
    excerpt = " ".join(parts) or "Content changed"
    return excerpt[:limit] + ("…" if len(excerpt) > limit else "")


def range_label(numbers: list[int]) -> str:
    if not numbers:
        return ""
    return str(numbers[0]) if len(numbers) == 1 else f"{numbers[0]}–{numbers[-1]}"


def describe_pdf_changes(old: dict[str, object], new: dict[str, object]) -> list[str]:
    old_pages = list(old.get("pages", []))
    new_pages = list(new.get("pages", []))
    old_signatures = [page.get("visual_sha256") or page.get("text_sha256") for page in old_pages]
    new_signatures = [page.get("visual_sha256") or page.get("text_sha256") for page in new_pages]
    matcher = difflib.SequenceMatcher(a=old_signatures, b=new_signatures, autojunk=False)
    details: list[str] = []
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        if operation == "delete":
            pages = [int(page["page"]) for page in old_pages[old_start:old_end]]
            details.append(f"Old page(s) {range_label(pages)} removed")
            continue
        if operation == "insert":
            pages = [int(page["page"]) for page in new_pages[new_start:new_end]]
            details.append(f"New page(s) {range_label(pages)} added")
            continue

        old_block = old_pages[old_start:old_end]
        new_block = new_pages[new_start:new_end]
        paired = min(len(old_block), len(new_block))
        for index in range(paired):
            old_page, new_page = old_block[index], new_block[index]
            old_number, new_number = int(old_page["page"]), int(new_page["page"])
            label = f"Page {new_number}" if old_number == new_number else f"Old page {old_number} / new page {new_number}"
            excerpt = compact_excerpt(
                normalized_lines(str(old_page.get("text", ""))),
                normalized_lines(str(new_page.get("text", ""))),
            )
            details.append(f"{label} changed — {excerpt}")
        if len(old_block) > paired:
            pages = [int(page["page"]) for page in old_block[paired:]]
            details.append(f"Old page(s) {range_label(pages)} removed")
        if len(new_block) > paired:
            pages = [int(page["page"]) for page in new_block[paired:]]
            details.append(f"New page(s) {range_label(pages)} added")
    if not details:
        details.append("File packaging or metadata changed; no rendered page difference was found")
    return details[:20] + ([f"{len(details) - 20} additional page change(s) omitted"] if len(details) > 20 else [])


def pdf_preview_pairs(old: dict[str, object], new: dict[str, object]) -> list[dict[str, object]]:
    if old.get("kind") != "pdf_pages" or new.get("kind") != "pdf_pages":
        return []
    old_pages = list(old.get("pages", []))
    new_pages = list(new.get("pages", []))
    old_signatures = [page.get("visual_sha256") or page.get("text_sha256") for page in old_pages]
    new_signatures = [page.get("visual_sha256") or page.get("text_sha256") for page in new_pages]
    matcher = difflib.SequenceMatcher(a=old_signatures, b=new_signatures, autojunk=False)
    previews: list[dict[str, object]] = []
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        paired = min(old_end - old_start, new_end - new_start)
        for offset in range(paired):
            old_page = old_pages[old_start + offset]
            new_page = new_pages[new_start + offset]
            old_path = str(old_page.get("preview_path", ""))
            new_path = str(new_page.get("preview_path", ""))
            if old_path and new_path and Path(old_path).is_file() and Path(new_path).is_file():
                previews.append({
                    "old_page": int(old_page["page"]),
                    "new_page": int(new_page["page"]),
                    "old_path": old_path,
                    "new_path": new_path,
                })
            if len(previews) >= MAX_PDF_PREVIEW_PAIRS:
                return previews
    return previews


def describe_sequence_changes(
    old_items: list[str],
    new_items: list[str],
    item_name: str,
) -> list[str]:
    matcher = difflib.SequenceMatcher(a=old_items, b=new_items, autojunk=False)
    details: list[str] = []
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        old_numbers = list(range(old_start + 1, old_end + 1))
        new_numbers = list(range(new_start + 1, new_end + 1))
        if operation == "insert":
            details.append(f"{item_name.title()}(s) {range_label(new_numbers)} added — {compact_excerpt([], new_items[new_start:new_end])}")
        elif operation == "delete":
            details.append(f"Old {item_name}(s) {range_label(old_numbers)} removed — {compact_excerpt(old_items[old_start:old_end], [])}")
        else:
            details.append(
                f"{item_name.title()}(s) {range_label(new_numbers)} changed — "
                f"{compact_excerpt(old_items[old_start:old_end], new_items[new_start:new_end])}"
            )
    return details[:20] + ([f"{len(details) - 20} additional change group(s) omitted"] if len(details) > 20 else [])


def parse_sqs_pay_item(line: str) -> dict[str, str] | None:
    """Parse one pipe-delimited SQS library row, ignoring its shifting row number."""
    parts = [part.strip() for part in line.split("|")]
    if len(parts) < 5 or not parts[0].isdigit() or not parts[1]:
        return None
    return {
        "code": parts[1],
        "description": parts[2],
        "unit": parts[3],
        "category": " | ".join(part for part in parts[4:] if part),
    }


def sqs_pay_items(analysis: dict[str, object]) -> dict[str, dict[str, str]]:
    items: dict[str, dict[str, str]] = {}
    for line in analysis.get("lines", []):
        item = parse_sqs_pay_item(str(line))
        if item is None:
            continue
        current = items.get(item["code"])
        # The combined English library can contain repeated or truncated copies.
        # Retain the most complete representation of each pay-item code.
        if current is None or sum(map(len, item.values())) > sum(map(len, current.values())):
            items[item["code"]] = item
    return items


def compare_sqs_pay_items(
    old_analysis: dict[str, object],
    new_analysis: dict[str, object],
) -> dict[str, list[dict[str, str]]]:
    old_items = sqs_pay_items(old_analysis)
    new_items = sqs_pay_items(new_analysis)
    return {
        "added": [new_items[code] for code in sorted(set(new_items) - set(old_items))],
        "removed": [old_items[code] for code in sorted(set(old_items) - set(new_items))],
        "updated": [
            {**new_items[code], "old_description": old_items[code]["description"],
             "old_unit": old_items[code]["unit"], "old_category": old_items[code]["category"]}
            for code in sorted(set(old_items) & set(new_items))
            if old_items[code] != new_items[code]
        ],
    }


def is_sqs_pay_item_document(document: dict[str, object]) -> bool:
    title = str(document.get("title", "")).casefold()
    url = unquote(urlsplit(str(document.get("url", ""))).path).casefold()
    return title.startswith("sqs-daily-") or "/payitems/sqs-daily-" in url


def describe_excel_changes(old: dict[str, object], new: dict[str, object]) -> list[str]:
    def cells(analysis: dict[str, object]) -> dict[tuple[str, str], str]:
        return {
            (str(item["sheet"]), str(item["cell"])): str(item.get("value", ""))
            for item in analysis.get("cells", [])
        }

    old_cells, new_cells = cells(old), cells(new)
    details: list[str] = []
    for key in sorted(set(old_cells) | set(new_cells)):
        sheet, cell = key
        if key not in old_cells:
            details.append(f"{sheet}!{cell} added: {new_cells[key][:300]}")
        elif key not in new_cells:
            details.append(f"{sheet}!{cell} removed: {old_cells[key][:300]}")
        elif old_cells[key] != new_cells[key]:
            details.append(f"{sheet}!{cell} changed: {old_cells[key][:150]} → {new_cells[key][:150]}")
    return details[:20] + ([f"{len(details) - 20} additional cell change(s) omitted"] if len(details) > 20 else [])


def describe_analysis_changes(old: dict[str, object], new: dict[str, object]) -> list[str]:
    old_analysis = old.get("analysis") or {}
    new_analysis = new.get("analysis") or {}
    if old_analysis.get("kind") != new_analysis.get("kind"):
        return ["The document format or detailed-analysis method changed"]
    kind = new_analysis.get("kind")
    if kind == "pdf_pages":
        return describe_pdf_changes(old_analysis, new_analysis)
    if kind == "paragraphs":
        return describe_sequence_changes(
            list(old_analysis.get("paragraphs", [])),
            list(new_analysis.get("paragraphs", [])),
            "paragraph",
        )
    if kind == "lines":
        return describe_sequence_changes(
            list(old_analysis.get("lines", [])),
            list(new_analysis.get("lines", [])),
            "line",
        )
    if kind == "excel_cells":
        return describe_excel_changes(old_analysis, new_analysis)
    return [str(new_analysis.get("reason", "Detailed internal comparison is unavailable for this file type"))]


def is_render_only_pdf_repackaging(old: dict[str, object], new: dict[str, object]) -> bool:
    """Return whether two PDF files differ internally but render identically.

    PDF producers commonly rewrite metadata, object ordering, or compression without
    changing a page the reader can see.  Those updates are recorded in the baseline
    through the file hash, but they are not useful change notifications.
    """
    old_analysis = old.get("analysis") or {}
    new_analysis = new.get("analysis") or {}
    if old_analysis.get("kind") != "pdf_pages" or new_analysis.get("kind") != "pdf_pages":
        return False
    old_pages = list(old_analysis.get("pages", []))
    new_pages = list(new_analysis.get("pages", []))
    if len(old_pages) != len(new_pages):
        return False
    return all(
        (old_page.get("visual_sha256") or old_page.get("text_sha256"))
        == (new_page.get("visual_sha256") or new_page.get("text_sha256"))
        for old_page, new_page in zip(old_pages, new_pages)
    )


def describe_page_text_changes(old_text: str, new_text: str) -> list[str]:
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(a=old_words, b=new_words, autojunk=False)
    details: list[str] = []
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        old_value = " ".join(old_words[old_start:old_end])
        new_value = " ".join(new_words[new_start:new_end])
        if operation == "insert":
            details.append(f"Added: {new_value[:500]}")
        elif operation == "delete":
            details.append(f"Removed: {old_value[:500]}")
        else:
            details.append(f"Changed: {old_value[:250]} → {new_value[:250]}")
    return details[:10] + ([f"{len(details) - 10} additional text change(s) omitted"] if len(details) > 10 else [])


def compare_snapshots(old: dict[str, object], new: dict[str, object]) -> dict[str, object]:
    old_schema = int(old.get("schema_version", 1))
    new_schema = int(new.get("schema_version", 1))
    folder_crawl_migration = old_schema < 2 <= new_schema
    page_text_inventory_migration = old_schema < 3 <= new_schema
    old_docs = {
        item.get("identity", item["url"]): item
        for item in old.get("documents", [])
        if not is_excluded_document(item)
    }
    new_docs = {
        item.get("identity", item["url"]): item
        for item in new.get("documents", [])
        if not is_excluded_document(item)
    }
    added_urls = set(new_docs) - set(old_docs)
    removed_urls = set(old_docs) - set(new_docs)
    moved: list[dict[str, object]] = []
    for removed_url in list(removed_urls):
        match = next(
            (url for url in added_urls if new_docs[url].get("sha256") == old_docs[removed_url].get("sha256")),
            None,
        )
        if match:
            moved.append({"old": old_docs[removed_url], "new": new_docs[match]})
            removed_urls.remove(removed_url)
            added_urls.remove(match)

    modified = [
        {
            "old": old_docs[url],
            "new": new_docs[url],
            "details": describe_analysis_changes(old_docs[url], new_docs[url]),
            "previews": pdf_preview_pairs(
                old_docs[url].get("analysis") or {},
                new_docs[url].get("analysis") or {},
            ),
        }
        for url in sorted(set(old_docs) & set(new_docs))
        if (
            old_docs[url].get("sha256") != new_docs[url].get("sha256")
            and not is_render_only_pdf_repackaging(old_docs[url], new_docs[url])
        )
    ]
    renamed = [
        {"old": old_docs[url], "new": new_docs[url]}
        for url in sorted(set(old_docs) & set(new_docs))
        if old_docs[url].get("title") != new_docs[url].get("title")
    ]
    relinked = [
        {"old": old_docs[key], "new": new_docs[key]}
        for key in sorted(set(old_docs) & set(new_docs))
        if old_docs[key].get("url") != new_docs[key].get("url")
    ]

    def link_map(snapshot: dict[str, object]) -> dict[str, dict[str, object]]:
        return {
            item["url"]: item
            for item in snapshot.get("links", [])
            if not is_excluded_document(item)
        }

    old_links, new_links = link_map(old), link_map(new)
    link_added = [new_links[url] for url in sorted(set(new_links) - set(old_links))]
    link_removed = [old_links[url] for url in sorted(set(old_links) - set(new_links))]
    link_renamed = [
        {
            "url": url,
            "old_title": old_links[url].get("title", ""),
            "new_title": new_links[url].get("title", ""),
            "section": new_links[url].get("section", old_links[url].get("section", "")),
        }
        for url in sorted(set(old_links) & set(new_links))
        if old_links[url].get("title", "") != new_links[url].get("title", "")
    ]
    old_folders = {item["path"]: item for item in old.get("folders", [])}
    new_folders = {item["path"]: item for item in new.get("folders", [])}
    compare_folders = "folders" in old and "folders" in new
    return {
        "page_text_changed": False if (folder_crawl_migration or page_text_inventory_migration) else old.get("page_text") != new.get("page_text"),
        "page_text_details": [] if (folder_crawl_migration or page_text_inventory_migration) else describe_page_text_changes(
            str(old.get("page_text", "")),
            str(new.get("page_text", "")),
        ),
        "documents_added": [] if folder_crawl_migration else [new_docs[url] for url in sorted(added_urls)],
        "documents_removed": [old_docs[url] for url in sorted(removed_urls)],
        "documents_modified": modified,
        "documents_renamed": renamed,
        "documents_relinked": relinked,
        "documents_moved": moved,
        "folders_added": [new_folders[path] for path in sorted(set(new_folders) - set(old_folders))]
        if compare_folders and not folder_crawl_migration else [],
        "folders_removed": [old_folders[path] for path in sorted(set(old_folders) - set(new_folders))]
        if compare_folders and not folder_crawl_migration else [],
        "links_added": [] if folder_crawl_migration else link_added,
        "links_removed": [] if folder_crawl_migration else link_removed,
        "links_renamed": [] if folder_crawl_migration else link_renamed,
    }


def has_changes(changes: dict[str, object]) -> bool:
    return bool(
        changes["page_text_changed"]
        or any(
            changes[key]
            for key in changes
            if key not in {"page_text_changed", "page_text_details"}
        )
    )


def change_count(changes: dict[str, object]) -> int:
    return int(bool(changes["page_text_changed"])) + sum(
        len(value)
        for key, value in changes.items()
        if key not in {"page_text_changed", "page_text_details"}
    )


def change_subject(item: dict[str, object]) -> dict[str, object]:
    for key in ("new", "old"):
        value = item.get(key)
        if isinstance(value, dict):
            return value
    return item


def filter_changes(changes: dict[str, object], filters: dict[str, object], include_page_text: bool = False) -> dict[str, object]:
    filtered: dict[str, object] = {
        "page_text_changed": bool(changes.get("page_text_changed")) and include_page_text,
        "page_text_details": list(changes.get("page_text_details", [])) if include_page_text else [],
    }
    for key, value in changes.items():
        if key in {"page_text_changed", "page_text_details"}:
            continue
        filtered[key] = [item for item in value if item_matches_filters(change_subject(item), filters)]
    return filtered


def notification_batches(
    changes: dict[str, object],
    config: dict[str, object],
) -> list[tuple[list[str], dict[str, object], str]]:
    batches: list[tuple[list[str], dict[str, object], str]] = []
    global_recipients = list(dict.fromkeys(str(value) for value in config.get("recipients", []) if str(value)))
    if global_recipients:
        batches.append((global_recipients, changes, "All monitored changes"))
    for rule in config.get("recipient_rules", []):
        recipients = [
            str(value) for value in rule.get("recipients", [])
            if str(value) and str(value) not in global_recipients
        ]
        routed = filter_changes(
            changes,
            rule.get("filters", {}),
            include_page_text=bool(rule.get("include_page_text", False)),
        )
        if recipients and has_changes(routed):
            batches.append((list(dict.fromkeys(recipients)), routed, str(rule["name"])))
    return batches


def preview_content_id(path: str) -> str:
    return "mdot-preview-" + hashlib.sha256(path.encode("utf-8")).hexdigest()[:20]


def preview_image_source(path: str, image_mode: str) -> str:
    if path.startswith("data:image/"):
        return path
    if image_mode == "data":
        try:
            encoded = base64.b64encode(Path(path).read_bytes()).decode("ascii")
            return "data:image/png;base64," + encoded
        except OSError:
            return ""
    return "cid:" + preview_content_id(path)


def inline_images_for_changes(changes: dict[str, object]) -> list[dict[str, str]]:
    images: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in changes.get("documents_modified", []):
        for preview in item.get("previews", []):
            for key in ("old_path", "new_path"):
                path = str(preview.get(key, ""))
                if path and path not in seen and Path(path).is_file():
                    seen.add(path)
                    images.append({"path": path, "content_id": preview_content_id(path)})
    return images


def format_change_email(
    changes: dict[str, object],
    snapshot: dict[str, object],
    image_mode: str = "cid",
) -> str:
    e = html_module.escape
    sections: list[str] = []
    link_style = (
        "color:#176b87;text-decoration:underline;font-weight:600;"
        "word-break:break-word;overflow-wrap:anywhere;"
    )

    def link(url: object, title: object) -> str:
        return f'<a href="{e(str(url))}" style="{link_style}">{e(str(title))}</a>'

    def item_rows(items: Iterable[dict[str, object]], formatter) -> str:
        rendered = list(items)
        return (
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'
            + "".join(
                '<tr><td width="20" valign="top" style="padding:0 8px 11px 0;'
                'font-family:Segoe UI,Arial,sans-serif;font-size:16px;line-height:22px;color:#d97706;">&#8226;</td>'
                '<td valign="top" style="padding:0 0 11px 0;font-family:Segoe UI,Arial,sans-serif;'
                'font-size:15px;line-height:22px;color:#253746;word-break:break-word;overflow-wrap:anywhere;">'
                f"{formatter(item)}</td></tr>"
                for item in rendered
            )
            + "</table>"
        )

    def add_section(title: str, items: Iterable[dict[str, object]], formatter) -> None:
        rendered = list(items)
        if not rendered:
            return
        content = item_rows(rendered, formatter)
        sections.append(
            '<tr><td style="padding:0 28px 16px 28px;">'
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
            'style="border:1px solid #d7e0e5;border-collapse:separate;">'
            '<tr><td style="padding:12px 16px;background-color:#f2f6f8;border-bottom:1px solid #d7e0e5;'
            'font-family:Segoe UI,Arial,sans-serif;font-size:16px;line-height:22px;font-weight:700;color:#12324a;">'
            f'{e(title)} <span style="font-size:12px;line-height:18px;color:#526675;font-weight:600;">({len(rendered)})</span>'
            '</td></tr><tr><td style="padding:15px 16px 4px 16px;">'
            f"{content}</td></tr></table></td></tr>"
        )

    def modified_document(item: dict[str, object]) -> str:
        new_document = item["new"]
        details = list(item.get("details", []))
        detail_html = ""
        if details:
            detail_html = (
                '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
                'style="margin-top:9px;background-color:#f7fafb;border-left:4px solid #5b8fa3;">'
                + "".join(
                    '<tr><td style="padding:8px 10px;font-family:Segoe UI,Arial,sans-serif;font-size:13px;'
                    'line-height:19px;color:#405463;word-break:break-word;overflow-wrap:anywhere;">'
                    f"{e(str(detail))}</td></tr>"
                    for detail in details
                )
                + "</table>"
            )
        preview_html = ""
        preview_rows = []
        for preview in item.get("previews", []):
            old_source = preview_image_source(str(preview.get("old_path", "")), image_mode)
            new_source = preview_image_source(str(preview.get("new_path", "")), image_mode)
            if not old_source or not new_source:
                continue
            old_page = int(preview["old_page"])
            new_page = int(preview["new_page"])
            preview_rows.append(
                '<tr><td width="50%" valign="top" style="padding:10px 5px 4px 0;">'
                f'<div style="padding-bottom:5px;font-size:11px;line-height:16px;font-weight:700;color:#687985;">BEFORE — PAGE {old_page}</div>'
                f'<img src="{e(old_source)}" alt="Before page {old_page}" width="280" style="display:block;width:100%;max-width:280px;height:auto;border:1px solid #cbd6dc;">'
                '</td><td width="50%" valign="top" style="padding:10px 0 4px 5px;">'
                f'<div style="padding-bottom:5px;font-size:11px;line-height:16px;font-weight:700;color:#687985;">AFTER — PAGE {new_page}</div>'
                f'<img src="{e(new_source)}" alt="After page {new_page}" width="280" style="display:block;width:100%;max-width:280px;height:auto;border:1px solid #cbd6dc;">'
                '</td></tr>'
            )
        if preview_rows:
            preview_html = (
                '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
                'style="margin-top:8px;">' + "".join(preview_rows) + "</table>"
            )
        return f'{link(new_document["url"], new_document["title"])}{detail_html}{preview_html}'

    sqs_documents = [
        item for item in changes["documents_modified"]
        if is_sqs_pay_item_document(item["new"])
    ]
    if sqs_documents:
        consolidated: dict[tuple[str, str], dict[str, str]] = {}
        for document_change in sqs_documents:
            pay_item_changes = compare_sqs_pay_items(
                document_change["old"].get("analysis") or {},
                document_change["new"].get("analysis") or {},
            )
            for change_kind, pay_items in pay_item_changes.items():
                for pay_item in pay_items:
                    consolidated[(change_kind, pay_item["code"])] = pay_item

        badge_styles = {
            "added": ("ADDED", "#e7f6ec", "#176b3a"),
            "removed": ("REMOVED", "#fdecec", "#9b2c2c"),
            "updated": ("UPDATED", "#e8f2f7", "#176b87"),
        }

        def pay_item_row(entry: tuple[tuple[str, str], dict[str, str]]) -> str:
            (change_kind, _), pay_item = entry
            badge, badge_background, badge_color = badge_styles[change_kind]
            previous = ""
            if change_kind == "updated":
                old_summary = " · ".join(
                    value for value in (
                        pay_item.get("old_description", ""),
                        pay_item.get("old_unit", ""),
                        pay_item.get("old_category", ""),
                    ) if value
                )
                previous = (
                    '<div style="padding-top:4px;font-size:12px;line-height:17px;color:#71818c;">'
                    f'Previously: {e(old_summary)}</div>'
                )
            metadata = " &nbsp;&bull;&nbsp; ".join(
                e(value) for value in (pay_item["unit"], pay_item["category"]) if value
            )
            return (
                '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'
                '<tr><td valign="top" width="82" style="padding-right:10px;">'
                f'<span style="display:inline-block;padding:3px 7px;background-color:{badge_background};'
                f'color:{badge_color};font-size:10px;line-height:14px;font-weight:700;letter-spacing:.4px;">'
                f'{badge}</span></td><td valign="top">'
                f'<div style="font-size:14px;line-height:20px;color:#253746;"><strong>{e(pay_item["code"])}</strong>'
                f' &mdash; {e(pay_item["description"])}</div>'
                f'<div style="padding-top:2px;font-size:12px;line-height:17px;color:#526675;">{metadata}</div>'
                f'{previous}</td></tr></table>'
            )

        if consolidated:
            sorted_pay_items = sorted(consolidated.items(), key=lambda entry: (entry[0][0], entry[0][1]))
            affected_links = " &nbsp;&bull;&nbsp; ".join(
                link(item["new"]["url"], item["new"]["title"])
                for item in sqs_documents
            )
            sections.append(
                '<tr><td style="padding:0 28px 16px 28px;">'
                '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
                'style="border:1px solid #d7e0e5;border-collapse:separate;">'
                '<tr><td style="padding:12px 16px;background-color:#f2f6f8;border-bottom:1px solid #d7e0e5;'
                'font-family:Segoe UI,Arial,sans-serif;font-size:16px;line-height:22px;font-weight:700;color:#12324a;">'
                f'SQS pay item library <span style="font-size:12px;line-height:18px;color:#526675;font-weight:600;">'
                f'({len(consolidated)} item change(s))</span></td></tr>'
                '<tr><td style="padding:13px 16px 5px 16px;font-family:Segoe UI,Arial,sans-serif;">'
                '<div style="padding-bottom:12px;font-size:12px;line-height:18px;color:#687985;">'
                'Duplicate entries across library variants are consolidated below.</div>'
                f'{item_rows(sorted_pay_items, pay_item_row)}'
                '<div style="padding:2px 0 9px 0;font-size:12px;line-height:18px;color:#687985;">'
                f'Affected files: {affected_links}</div></td></tr></table></td></tr>'
            )

    add_section("Documents added", changes["documents_added"], lambda x: link(x["url"], x["title"]))
    add_section("Documents removed", changes["documents_removed"], lambda x: e(str(x["title"])))
    add_section(
        "Document contents modified",
        [item for item in changes["documents_modified"] if item not in sqs_documents],
        modified_document,
    )
    add_section("Folders added", changes.get("folders_added", []), lambda x: e(str(x["path"])))
    add_section("Folders removed", changes.get("folders_removed", []), lambda x: e(str(x["path"])))
    add_section(
        "Documents renamed",
        changes["documents_renamed"],
        lambda x: f'{e(str(x["old"]["title"]))} <span style="color:#7a8994;">&rarr;</span> '
        f'{link(x["new"]["url"], x["new"]["title"])}',
    )
    add_section(
        "Document links changed",
        changes["documents_relinked"],
        lambda x: f'<span style="color:#526675;word-break:break-all;">{e(str(x["old"]["url"]))}</span> '
        f'<span style="color:#7a8994;">&rarr;</span> {link(x["new"]["url"], x["new"]["title"])}',
    )
    add_section(
        "Documents moved or replaced at a new URL",
        changes["documents_moved"],
        lambda x: f'<span style="color:#526675;word-break:break-all;">{e(str(x["old"]["url"]))}</span> '
        f'<span style="color:#7a8994;">&rarr;</span> {link(x["new"]["url"], x["new"]["title"])}',
    )
    add_section("Links added", changes["links_added"], lambda x: link(x["url"], x["title"]))
    add_section(
        "Links removed",
        changes["links_removed"],
        lambda x: f'{e(str(x["title"]))}<br><span style="font-size:12px;color:#687985;word-break:break-all;">'
        f'{e(str(x["url"]))}</span>',
    )
    add_section(
        "Link titles changed",
        changes["links_renamed"],
        lambda x: f'{e(str(x["old_title"]))} <span style="color:#7a8994;">&rarr;</span> '
        f'{link(x["url"], x["new_title"])}',
    )
    if changes["page_text_changed"]:
        page_details = list(changes.get("page_text_details", []))
        detail_items = [{"detail": detail} for detail in page_details]
        detail_html = item_rows(detail_items, lambda item: e(str(item["detail"]))) if detail_items else ""
        sections.insert(
            0,
            '<tr><td style="padding:0 28px 16px 28px;">'
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
            'style="border:1px solid #d7e0e5;border-collapse:separate;">'
            '<tr><td style="padding:12px 16px;background-color:#f2f6f8;border-bottom:1px solid #d7e0e5;'
            'font-family:Segoe UI,Arial,sans-serif;font-size:16px;line-height:22px;font-weight:700;color:#12324a;">'
            'Page content changed</td></tr><tr><td style="padding:14px 16px 4px 16px;">'
            '<p style="margin:0 0 12px 0;font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:21px;color:#526675;">'
            'Visible headings or explanatory text changed.</p>'
            f"{detail_html}</td></tr></table></td></tr>",
        )
    checked = e(str(snapshot["generated_at"]))
    page_url = e(str(snapshot["page_url"]))
    total = change_count(changes)
    return (
        '<!doctype html><html><head><meta charset="utf-8"></head>'
        '<body style="margin:0;padding:0;background-color:#eaf0f3;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#eaf0f3;">'
        '<tr><td align="center" style="padding:24px 10px;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="width:100%;max-width:760px;background-color:#ffffff;border:1px solid #ced9df;">'
        '<!-- LOCAL-TEST-BANNER -->'
        '<tr><td style="padding:0;background-color:#d97706;height:6px;font-size:0;line-height:0;">&nbsp;</td></tr>'
        '<tr><td style="padding:28px;background-color:#12324a;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        '<td valign="middle" style="font-family:Segoe UI,Arial,sans-serif;color:#ffffff;">'
        '<div style="font-size:12px;line-height:18px;font-weight:700;letter-spacing:1px;color:#9fd3e3;">MDOT STANDARDS MONITOR</div>'
        '<div style="padding-top:5px;font-size:25px;line-height:32px;font-weight:700;">Engineering standards update</div>'
        '</td><td width="92" align="center" valign="middle" style="padding-left:16px;">'
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;">'
        f'<tr><td align="center" style="padding:9px 14px;font-family:Segoe UI,Arial,sans-serif;font-size:24px;line-height:26px;font-weight:700;color:#12324a;">{total}</td></tr>'
        '<tr><td align="center" style="padding:0 10px 8px 10px;font-family:Segoe UI,Arial,sans-serif;font-size:10px;line-height:12px;font-weight:700;color:#526675;">CHANGES</td></tr>'
        '</table></td></tr></table></td></tr>'
        '<tr><td style="padding:20px 28px 18px 28px;">'
        '<p style="margin:0;font-family:Segoe UI,Arial,sans-serif;font-size:15px;line-height:23px;color:#364b5a;">'
        f'The daily check detected <strong>{total} change(s)</strong>.</p>'
        '<p style="margin:5px 0 0 0;font-family:Segoe UI,Arial,sans-serif;font-size:12px;line-height:18px;color:#71818c;">'
        f'Checked {checked}</p></td></tr>'
        + "".join(sections)
        + '<tr><td style="padding:4px 28px 28px 28px;">'
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'
        '<td style="background-color:#176b87;padding:12px 18px;">'
        f'<a href="{page_url}" style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:18px;font-weight:700;color:#ffffff;text-decoration:none;">Open MDOT standards page&nbsp;&rarr;</a>'
        '</td></tr></table></td></tr>'
        '<tr><td style="padding:16px 28px;background-color:#f2f6f8;border-top:1px solid #d7e0e5;'
        'font-family:Segoe UI,Arial,sans-serif;font-size:11px;line-height:17px;color:#687985;">'
        'Automated by MDOT Standards Monitor. This message reports detected changes; it does not modify MDOT content.'
        '</td></tr></table></td></tr></table></body></html>'
    )


def format_status_email(title: str, message: str, page_url: str) -> str:
    return (
        '<!doctype html><html><head><meta charset="utf-8"></head>'
        '<body style="margin:0;padding:0;background-color:#eaf0f3;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        '<td align="center" style="padding:24px 10px;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="width:100%;max-width:680px;background-color:#ffffff;border:1px solid #ced9df;">'
        '<tr><td style="padding:0;background-color:#d97706;height:6px;font-size:0;line-height:0;">&nbsp;</td></tr>'
        '<tr><td style="padding:24px 28px;background-color:#12324a;font-family:Segoe UI,Arial,sans-serif;'
        f'font-size:22px;line-height:29px;font-weight:700;color:#ffffff;">{html_module.escape(title)}</td></tr>'
        '<tr><td style="padding:24px 28px;font-family:Segoe UI,Arial,sans-serif;font-size:15px;line-height:23px;color:#364b5a;">'
        f'{html_module.escape(message)}</td></tr><tr><td style="padding:0 28px 28px 28px;">'
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'
        '<td style="background-color:#176b87;padding:12px 18px;">'
        f'<a href="{html_module.escape(page_url)}" style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;'
        'line-height:18px;font-weight:700;color:#ffffff;text-decoration:none;">Open monitored MDOT page&nbsp;&rarr;</a>'
        '</td></tr></table></td></tr></table></td></tr></table></body></html>'
    )


def setup_logging() -> logging.Logger:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(APP_NAME)
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    file_handler.setFormatter(formatter)
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console)
    return logger


def read_json(path: Path, default=None):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise MonitorError(f"Could not read {path}: {exc}") from exc


def atomic_write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, delete=False, suffix=".tmp") as temp:
        json.dump(value, temp, indent=2, ensure_ascii=False)
        temp.write("\n")
        temp_path = Path(temp.name)
    os.replace(temp_path, path)


def summarize_changes(changes: dict[str, object]) -> list[str]:
    lines: list[str] = []
    if changes.get("page_text_changed"):
        lines.append("Visible page content changed")
    labels = (
        ("folders_added", "folder(s) added"),
        ("folders_removed", "folder(s) removed"),
        ("documents_added", "document(s) added"),
        ("documents_removed", "document(s) removed"),
        ("documents_modified", "document(s) modified"),
        ("documents_renamed", "document(s) renamed"),
        ("documents_relinked", "document link(s) changed"),
        ("documents_moved", "document(s) moved"),
        ("links_added", "page link(s) added"),
        ("links_removed", "page link(s) removed"),
        ("links_renamed", "page link title(s) changed"),
    )
    for key, label in labels:
        count = len(changes.get(key, []))
        if count:
            lines.append(f"{count} {label}")
    return lines


def write_history_dashboard(events: list[dict[str, object]]) -> None:
    e = html_module.escape
    rows = []
    colors = {"change": "#d97706", "heartbeat": "#2f855a", "failure": "#b42318", "suppressed": "#687985"}
    for event in events:
        kind = str(event.get("kind", "event"))
        report = str(event.get("report", ""))
        report_link = ""
        if report and Path(report).is_file():
            report_link = f'<a href="{e(Path(report).resolve().as_uri())}" style="color:#176b87;font-weight:700;">View report</a>'
        details = "<br>".join(e(str(value)) for value in event.get("details", []))
        rows.append(
            '<tr><td style="padding:14px;border-bottom:1px solid #d7e0e5;vertical-align:top;white-space:nowrap;">'
            f'<span style="display:inline-block;padding:4px 8px;background:{colors.get(kind, "#526675")};color:white;'
            f'font-size:11px;font-weight:700;text-transform:uppercase;">{e(kind)}</span></td>'
            '<td style="padding:14px;border-bottom:1px solid #d7e0e5;vertical-align:top;">'
            f'<strong>{e(str(event.get("title", "Monitor event")))}</strong><br>'
            f'<span style="color:#687985;font-size:12px;">{e(str(event.get("timestamp", "")))}</span>'
            f'<div style="padding-top:6px;line-height:20px;">{details}</div></td>'
            f'<td style="padding:14px;border-bottom:1px solid #d7e0e5;vertical-align:top;white-space:nowrap;">{report_link}</td></tr>'
        )
    body = "".join(rows) or '<tr><td style="padding:24px;">No monitor history has been recorded yet.</td></tr>'
    html = (
        '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>MDOT Standards Monitor History</title></head><body style="margin:0;background:#eaf0f3;font-family:Segoe UI,Arial,sans-serif;color:#253746;">'
        '<div style="max-width:980px;margin:24px auto;background:white;border:1px solid #ced9df;">'
        '<div style="height:6px;background:#d97706;"></div><div style="padding:24px 28px;background:#12324a;color:white;">'
        '<div style="font-size:12px;font-weight:700;letter-spacing:1px;color:#9fd3e3;">MDOT STANDARDS MONITOR</div>'
        '<h1 style="margin:6px 0 0;font-size:26px;">Change history</h1></div>'
        '<div style="padding:18px 28px;color:#526675;">Newest events appear first. Reports and previews remain on this computer.</div>'
        '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">'
        f'{body}</table></div></body></html>'
    )
    HISTORY_DASHBOARD.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_DASHBOARD.write_text(html, encoding="utf-8")


def record_history(
    kind: str,
    title: str,
    details: list[str],
    limit: int = DEFAULT_HISTORY_LIMIT,
    changes: dict[str, object] | None = None,
    snapshot: dict[str, object] | None = None,
) -> dict[str, object]:
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    event: dict[str, object] = {"timestamp": timestamp, "kind": kind, "title": title, "details": details}
    if changes is not None and snapshot is not None:
        HISTORY_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        report_name = re.sub(r"[^0-9A-Za-z_-]", "-", timestamp) + ".html"
        report_path = HISTORY_REPORTS_DIR / report_name
        report_path.write_text(format_change_email(changes, snapshot, image_mode="data"), encoding="utf-8")
        event["report"] = str(report_path)
        event["change_count"] = change_count(changes)
    events = read_json(HISTORY_FILE, []) or []
    if not isinstance(events, list):
        events = []
    events.insert(0, event)
    keep = max(1, int(limit))
    removed = events[keep:]
    events = events[:keep]
    for old_event in removed:
        old_report = str(old_event.get("report", "")) if isinstance(old_event, dict) else ""
        if old_report:
            Path(old_report).unlink(missing_ok=True)
    atomic_write_json(HISTORY_FILE, events)
    write_history_dashboard(events)
    return event


def cleanup_preview_cache(snapshot: dict[str, object]) -> None:
    if not PREVIEW_DIR.is_dir():
        return
    keep = {str(item.get("sha256", "")) for item in snapshot.get("documents", [])}
    for directory in PREVIEW_DIR.iterdir():
        if directory.is_dir() and re.fullmatch(r"[0-9a-f]{64}", directory.name) and directory.name not in keep:
            shutil.rmtree(directory, ignore_errors=True)


class FileLock:
    def __enter__(self):
        BASE_DIR.mkdir(parents=True, exist_ok=True)
        if LOCK_FILE.exists() and time.time() - LOCK_FILE.stat().st_mtime > 6 * 60 * 60:
            LOCK_FILE.unlink(missing_ok=True)
        try:
            descriptor = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError as exc:
            raise MonitorError("Another monitor run is already in progress") from exc
        with os.fdopen(descriptor, "w", encoding="ascii") as handle:
            handle.write(str(os.getpid()))
        return self

    def __exit__(self, exc_type, exc, traceback):
        LOCK_FILE.unlink(missing_ok=True)


def load_config(require_recipients: bool = False) -> dict[str, object]:
    config = read_json(CONFIG_FILE, {}) or {}
    config.setdefault("page_url", DEFAULT_URL)
    config.setdefault("recipients", [])
    config.setdefault("failure_recipient", "")
    config.setdefault("extra_documents", [])
    config.setdefault("dynamic_documents", [])
    config.setdefault("filters", {"sections": [], "titles": [], "extensions": []})
    config.setdefault("recipient_rules", [])
    config.setdefault("retry_attempts", DEFAULT_RETRY_ATTEMPTS)
    config.setdefault("retry_delay_seconds", DEFAULT_RETRY_DELAY_SECONDS)
    config.setdefault("confirmation_delay_seconds", DEFAULT_CONFIRMATION_DELAY_SECONDS)
    config.setdefault("heartbeat_days", DEFAULT_HEARTBEAT_DAYS)
    config.setdefault("history_limit", DEFAULT_HISTORY_LIMIT)
    recipients = config["recipients"]
    if not isinstance(recipients, list):
        raise MonitorError(f"recipients must be a list in {CONFIG_FILE}")
    if require_recipients and not recipients:
        raise MonitorError(f"No recipients are configured. Run install.ps1 first. Config: {CONFIG_FILE}")
    if not isinstance(config["failure_recipient"], str):
        raise MonitorError(f"failure_recipient must be a string in {CONFIG_FILE}")
    if not isinstance(config["extra_documents"], list) or any(
        not isinstance(item, dict) or not item.get("url") or not item.get("title")
        for item in config["extra_documents"]
    ):
        raise MonitorError(f"extra_documents must be a list of objects with url and title in {CONFIG_FILE}")
    if not isinstance(config["dynamic_documents"], list) or any(
        not isinstance(item, dict)
        or not item.get("source_url")
        or not item.get("title")
        or not item.get("match_host")
        for item in config["dynamic_documents"]
    ):
        raise MonitorError(
            f"dynamic_documents must include source_url, title, and match_host in {CONFIG_FILE}"
        )
    if not isinstance(config["filters"], dict) or any(
        not isinstance(config["filters"].get(key, []), list)
        for key in ("sections", "titles", "extensions")
    ):
        raise MonitorError(f"filters must contain list values for sections, titles, and extensions in {CONFIG_FILE}")
    if not isinstance(config["recipient_rules"], list) or any(
        not isinstance(rule, dict)
        or not rule.get("name")
        or not isinstance(rule.get("recipients", []), list)
        or not isinstance(rule.get("filters", {}), dict)
        for rule in config["recipient_rules"]
    ):
        raise MonitorError(f"recipient_rules must contain name, recipients, and filters in {CONFIG_FILE}")
    for key in ("retry_attempts", "retry_delay_seconds", "confirmation_delay_seconds", "heartbeat_days", "history_limit"):
        if not isinstance(config[key], (int, float)) or config[key] < 0:
            raise MonitorError(f"{key} must be a non-negative number in {CONFIG_FILE}")
    return config


def send_outlook(
    subject: str,
    body_html: str,
    recipients: list[str],
    inline_images: list[dict[str, str]] | None = None,
) -> None:
    if not EMAIL_HELPER.exists():
        raise MonitorError(f"Outlook helper is missing: {EMAIL_HELPER}")
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".html", delete=False) as body_file:
        body_file.write(body_html)
        body_path = Path(body_file.name)
    image_manifest_path: Path | None = None
    try:
        command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(EMAIL_HELPER), "-Recipients", ";".join(recipients),
            "-Subject", subject, "-HtmlBodyFile", str(body_path),
        ]
        if inline_images:
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as manifest:
                json.dump(inline_images, manifest, ensure_ascii=False)
                image_manifest_path = Path(manifest.name)
            command.extend(["-InlineImagesJsonFile", str(image_manifest_path)])
        result = subprocess.run(command, capture_output=True, text=True, timeout=90, check=False)
        if result.returncode != 0:
            detail = normalize_text(result.stderr or result.stdout)
            raise MonitorError(f"Outlook could not send the email: {detail}")
    finally:
        body_path.unlink(missing_ok=True)
        if image_manifest_path is not None:
            image_manifest_path.unlink(missing_ok=True)


def default_runtime() -> dict[str, object]:
    return {
        "consecutive_failures": 0,
        "last_error": "",
        "last_attempt_at": "",
        "last_success_at": "",
        "last_change_at": "",
        "last_notification_at": "",
        "last_heartbeat_at": "",
        "last_document_count": 0,
        "last_change_count": 0,
        "last_duration_seconds": 0.0,
    }


def load_runtime() -> dict[str, object]:
    runtime = read_json(RUNTIME_FILE, {}) or {}
    defaults = default_runtime()
    defaults.update(runtime)
    return defaults


def record_failure(error: Exception, config: dict[str, object], logger: logging.Logger) -> None:
    runtime = load_runtime()
    runtime["consecutive_failures"] = int(runtime.get("consecutive_failures", 0)) + 1
    runtime["last_error"] = str(error)
    failure_recipient = str(config.get("failure_recipient", "")).strip()
    if failure_recipient:
        body = format_status_email(
            "MDOT standards monitor failure",
            f"The scheduled check or its change-notification email failed. Error: {error}",
            str(config.get("page_url", DEFAULT_URL)),
        )
        try:
            send_outlook(
                "[MDOT Standards] Monitor failure",
                body,
                [failure_recipient],
            )
            logger.info("Sent failure notice to the configured failure recipient")
        except Exception as notification_error:
            logger.error("Could not send the failure notice: %s", notification_error)
    else:
        logger.warning("No failure recipient is configured; failure was logged only")
    atomic_write_json(RUNTIME_FILE, runtime)


def notify_document_errors(
    snapshot: dict[str, object],
    config: dict[str, object],
    logger: logging.Logger,
) -> None:
    """Privately report unavailable documents without failing an otherwise valid check."""
    errors = list(snapshot.get("document_errors", []))
    if not errors:
        return
    recipient = str(config.get("failure_recipient", "")).strip()
    if not recipient:
        logger.warning("%d document warning(s) were logged, but no failure recipient is configured", len(errors))
        return
    details = " | ".join(
        f"{normalize_text(str(item.get('title', 'Document')))}: {normalize_text(str(item.get('error', 'Unavailable')))}"
        for item in errors
    )
    if len(details) > 3500:
        details = details[:3499] + "…"
    message = (
        f"The check continued after {len(errors)} document(s) could not be loaded. "
        "Their last known snapshots were retained when available, and other MDOT changes were processed normally. "
        f"Details: {details}"
    )
    try:
        send_outlook(
            "[MDOT Standards] Document download warning",
            format_status_email("MDOT document download warning", message, str(config.get("page_url", DEFAULT_URL))),
            [recipient],
        )
        logger.info("Sent document warning to the configured failure recipient")
    except Exception as notification_error:
        logger.error("Could not send the document warning: %s", notification_error)


def clear_failure_state(config: dict[str, object], logger: logging.Logger) -> None:
    runtime = load_runtime()
    runtime["consecutive_failures"] = 0
    runtime["last_error"] = ""
    atomic_write_json(RUNTIME_FILE, runtime)


def heartbeat_is_due(runtime: dict[str, object], heartbeat_days: float) -> bool:
    if heartbeat_days <= 0:
        return False
    value = str(runtime.get("last_heartbeat_at", ""))
    if not value:
        return True
    try:
        last = datetime.fromisoformat(value)
    except ValueError:
        return True
    return (datetime.now().astimezone() - last).total_seconds() >= heartbeat_days * 86400


def save_success_runtime(
    runtime: dict[str, object],
    snapshot: dict[str, object],
    started: float,
    changes: int = 0,
) -> None:
    runtime.update({
        "consecutive_failures": 0,
        "last_error": "",
        "last_success_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "last_document_count": len(snapshot.get("documents", [])),
        "last_change_count": changes,
        "last_duration_seconds": round(time.monotonic() - started, 2),
    })
    atomic_write_json(RUNTIME_FILE, runtime)


def run_check(args: argparse.Namespace) -> int:
    logger = setup_logging()
    config = load_config(require_recipients=False)
    started = time.monotonic()
    runtime = load_runtime()
    if not args.dry_run:
        runtime["last_attempt_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
        atomic_write_json(RUNTIME_FILE, runtime)
    with FileLock():
        try:
            baseline = read_json(STATE_FILE)
            def create_snapshot():
                return build_snapshot(
                    str(config["page_url"]),
                    logger,
                    previous_snapshot=baseline,
                    force_analysis=False,
                    extra_documents=list(config.get("extra_documents", [])),
                    dynamic_documents=list(config.get("dynamic_documents", [])),
                    filters=dict(config.get("filters", {})),
                    store_previews=not args.dry_run,
                    retry_attempts=int(config["retry_attempts"]),
                    retry_delay_seconds=float(config["retry_delay_seconds"]),
                )

            snapshot = create_snapshot()
            if args.dry_run:
                if baseline:
                    changes = compare_snapshots(baseline, snapshot)
                    print(json.dumps(changes, indent=2))
                    print(f"Dry run complete: {change_count(changes)} change(s); no state or email changed.")
                else:
                    print(f"Dry run complete: found {len(snapshot['documents'])} documents; no baseline exists.")
                return 0

            if args.initialize or baseline is None:
                atomic_write_json(STATE_FILE, snapshot)
                notify_document_errors(snapshot, config, logger)
                save_success_runtime(runtime, snapshot, started)
                cleanup_preview_cache(snapshot)
                logger.info("Baseline initialized with %d documents; no update email sent", len(snapshot["documents"]))
                return 0

            changes = compare_snapshots(baseline, snapshot)
            if not has_changes(changes):
                atomic_write_json(STATE_FILE, snapshot)
                if (
                    float(config["heartbeat_days"]) > 0
                    and not runtime.get("last_heartbeat_at")
                    and not runtime.get("last_success_at")
                ):
                    # Start the weekly interval without surprising an upgraded installation
                    # with an immediate health email on its first successful run.
                    runtime["last_heartbeat_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
                elif not snapshot.get("document_errors") and heartbeat_is_due(runtime, float(config["heartbeat_days"])):
                    heartbeat_recipient = str(config.get("failure_recipient", "")).strip()
                    if heartbeat_recipient:
                        send_outlook(
                            "[MDOT Standards] Weekly monitor health summary",
                            format_status_email(
                                "MDOT standards monitor is operating normally",
                                f"The latest check completed successfully. The baseline currently contains {len(snapshot['documents'])} monitored documents, and no new MDOT changes were detected.",
                                str(config["page_url"]),
                            ),
                            [heartbeat_recipient],
                        )
                        runtime["last_heartbeat_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
                        record_history(
                            "heartbeat",
                            "Weekly health summary sent",
                            [f"Monitoring {len(snapshot['documents'])} documents", "No MDOT changes detected"],
                            int(config["history_limit"]),
                        )
                    else:
                        logger.warning("Weekly health summary was due, but no failure recipient is configured")
                notify_document_errors(snapshot, config, logger)
                save_success_runtime(runtime, snapshot, started)
                cleanup_preview_cache(snapshot)
                logger.info("No MDOT updates detected")
                return 0

            confirmation_delay = float(config["confirmation_delay_seconds"])
            if confirmation_delay > 0 and not getattr(args, "no_confirm", False):
                logger.info("Waiting %.1f seconds before confirming detected changes", confirmation_delay)
                time.sleep(confirmation_delay)
                confirmed_snapshot = create_snapshot()
                confirmed_changes = compare_snapshots(baseline, confirmed_snapshot)
                if not has_changes(confirmed_changes):
                    atomic_write_json(STATE_FILE, confirmed_snapshot)
                    record_history(
                        "suppressed",
                        "Transient change suppressed",
                        ["The first check detected a change, but the confirmation check returned to the baseline."],
                        int(config["history_limit"]),
                    )
                    notify_document_errors(confirmed_snapshot, config, logger)
                    save_success_runtime(runtime, confirmed_snapshot, started)
                    cleanup_preview_cache(confirmed_snapshot)
                    logger.info("Suppressed a transient change that did not survive confirmation")
                    return 0
                snapshot = confirmed_snapshot
                changes = confirmed_changes

            batches = notification_batches(changes, config)
            if not batches:
                raise MonitorError("Changes were detected, but no matching notification recipients are configured")
            for recipients, routed_changes, label in batches:
                subject = f"[MDOT Standards] {change_count(routed_changes)} update(s) detected"
                send_outlook(
                    subject,
                    format_change_email(routed_changes, snapshot),
                    recipients,
                    inline_images_for_changes(routed_changes),
                )
                logger.info("Sent %s update email to %d recipient(s)", label, len(recipients))
            atomic_write_json(STATE_FILE, snapshot)
            now = datetime.now().astimezone().isoformat(timespec="seconds")
            runtime["last_change_at"] = now
            runtime["last_notification_at"] = now
            record_history(
                "change",
                f"{change_count(changes)} MDOT update(s) detected",
                summarize_changes(changes),
                int(config["history_limit"]),
                changes,
                snapshot,
            )
            notify_document_errors(snapshot, config, logger)
            save_success_runtime(runtime, snapshot, started, change_count(changes))
            cleanup_preview_cache(snapshot)
            return 0
        except Exception as exc:
            logger.exception("Monitor check failed: %s", exc)
            if not args.dry_run:
                record_failure(exc, config, logger)
                try:
                    record_history(
                        "failure",
                        "Monitor check failed",
                        [str(exc)],
                        int(config.get("history_limit", DEFAULT_HISTORY_LIMIT)),
                    )
                except Exception as history_error:
                    logger.error("Could not update history after failure: %s", history_error)
            return 1


def outlook_is_registered() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID"):
            return True
    except OSError:
        return False


def scheduled_task_status() -> dict[str, object]:
    if os.name != "nt":
        return {"installed": False, "state": "Windows only"}
    script = (
        f"$task = Get-ScheduledTask -TaskName '{TASK_NAME}' -ErrorAction SilentlyContinue; "
        "if ($task) { $info = Get-ScheduledTaskInfo -TaskName $task.TaskName; "
        "[pscustomobject]@{installed=$true;state=[string]$task.State;last_run=[string]$info.LastRunTime;"
        "last_result=$info.LastTaskResult;next_run=[string]$info.NextRunTime;"
        "execute=[string]$task.Actions[0].Execute;arguments=[string]$task.Actions[0].Arguments} | ConvertTo-Json -Compress } "
        "else { '{\"installed\":false,\"state\":\"Not installed\"}' }"
    )
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", script],
        capture_output=True,
        text=True,
        timeout=20,
        check=False,
    )
    try:
        return json.loads(result.stdout.strip())
    except json.JSONDecodeError:
        return {"installed": False, "state": "Could not query", "error": normalize_text(result.stderr)}


def collect_status() -> dict[str, object]:
    config = load_config(require_recipients=False)
    runtime = load_runtime()
    baseline = read_json(STATE_FILE, {}) or {}
    history = read_json(HISTORY_FILE, []) or []
    try:
        chrome = str(find_chrome())
    except MonitorError:
        chrome = "Not found"
    git_revision = "Unknown"
    result = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=SCRIPT_DIR,
        capture_output=True,
        text=True,
        timeout=10,
        check=False,
    )
    if result.returncode == 0:
        git_revision = result.stdout.strip()
    return {
        "version": git_revision,
        "data_directory": str(BASE_DIR),
        "config_file": str(CONFIG_FILE),
        "configured_recipients": len(config.get("recipients", [])),
        "recipient_rules": len(config.get("recipient_rules", [])),
        "filters": normalize_filters(dict(config.get("filters", {}))),
        "baseline_generated_at": baseline.get("generated_at", "Not initialized"),
        "baseline_documents": len(baseline.get("documents", [])),
        "history_events": len(history) if isinstance(history, list) else 0,
        "runtime": runtime,
        "chrome": chrome,
        "classic_outlook_registered": outlook_is_registered(),
        "scheduled_task": scheduled_task_status(),
    }


def run_status(as_json: bool = False) -> int:
    status = collect_status()
    if as_json:
        print(json.dumps(status, indent=2, ensure_ascii=False))
        return 0
    runtime = status["runtime"]
    task = status["scheduled_task"]
    lines = [
        "MDOT Standards Monitor status",
        f"  Version: {status['version']}",
        f"  Baseline: {status['baseline_documents']} documents ({status['baseline_generated_at']})",
        f"  Last success: {runtime.get('last_success_at') or 'Never recorded'}",
        f"  Last change: {runtime.get('last_change_at') or 'None recorded'}",
        f"  Last error: {runtime.get('last_error') or 'None'}",
        f"  Consecutive failures: {runtime.get('consecutive_failures', 0)}",
        f"  Last run duration: {runtime.get('last_duration_seconds', 0)} seconds",
        f"  History events: {status['history_events']}",
        f"  Recipients: {status['configured_recipients']} global, {status['recipient_rules']} routing rule(s)",
        f"  Chrome: {status['chrome']}",
        f"  Classic Outlook registered: {'Yes' if status['classic_outlook_registered'] else 'No'}",
        f"  Scheduled task: {task.get('state', 'Unknown')}",
    ]
    if task.get("installed"):
        lines.extend([
            f"  Task last run: {task.get('last_run') or 'Unknown'} (result {task.get('last_result')})",
            f"  Task next run: {task.get('next_run') or 'Unknown'}",
        ])
    lines.append(f"  Data: {status['data_directory']}")
    print("\n".join(lines))
    return 0


def run_history(open_dashboard: bool = False) -> int:
    events = read_json(HISTORY_FILE, []) or []
    if not isinstance(events, list):
        raise MonitorError(f"History is malformed: {HISTORY_FILE}")
    write_history_dashboard(events)
    print(f"History contains {len(events)} event(s).")
    print(f"Dashboard: {HISTORY_DASHBOARD}")
    if open_dashboard:
        import webbrowser
        webbrowser.open(HISTORY_DASHBOARD.resolve().as_uri())
    return 0


def valid_email(value: str) -> bool:
    return bool(re.fullmatch(r"[^\s;@]+@[^\s;@]+\.[^\s;@]+", value.strip()))


def save_config(config: dict[str, object]) -> None:
    atomic_write_json(CONFIG_FILE, config)
    load_config(require_recipients=False)


def run_config_command(args: argparse.Namespace) -> int:
    config = load_config(require_recipients=False)
    action = args.config_action
    if action == "show":
        print(json.dumps(config, indent=2, ensure_ascii=False))
        return 0
    if action == "validate":
        print(f"Configuration is valid: {CONFIG_FILE}")
        return 0
    if action in {"add-recipient", "remove-recipient", "set-failure-recipient"}:
        address = args.email.strip()
        if action != "remove-recipient" and address and not valid_email(address):
            raise MonitorError(f"Recipient does not look like an email address: {address}")
        if action == "add-recipient":
            config["recipients"] = list(dict.fromkeys(list(config["recipients"]) + [address]))
        elif action == "remove-recipient":
            config["recipients"] = [value for value in config["recipients"] if str(value).casefold() != address.casefold()]
        else:
            config["failure_recipient"] = address
    elif action == "set-filter":
        config["filters"] = {
            "sections": list(args.section or []),
            "titles": list(args.title or []),
            "extensions": [("." + value.lstrip(".")).lower() for value in (args.extension or [])],
        }
    elif action == "clear-filters":
        config["filters"] = {"sections": [], "titles": [], "extensions": []}
    elif action == "add-route":
        recipients = [value.strip() for value in args.recipients.split(";") if value.strip()]
        invalid = [value for value in recipients if not valid_email(value)]
        if invalid:
            raise MonitorError("Invalid route recipient(s): " + ", ".join(invalid))
        route = {
            "name": args.name,
            "recipients": recipients,
            "include_page_text": bool(args.include_page_text),
            "filters": {
                "sections": list(args.section or []),
                "titles": list(args.title or []),
                "extensions": [("." + value.lstrip(".")).lower() for value in (args.extension or [])],
            },
        }
        config["recipient_rules"] = [rule for rule in config["recipient_rules"] if str(rule["name"]).casefold() != args.name.casefold()]
        config["recipient_rules"].append(route)
    elif action == "remove-route":
        config["recipient_rules"] = [rule for rule in config["recipient_rules"] if str(rule["name"]).casefold() != args.name.casefold()]
    elif action == "set-heartbeat-days":
        if args.days < 0:
            raise MonitorError("Heartbeat days cannot be negative")
        config["heartbeat_days"] = args.days
    elif action == "set-confirmation-delay":
        if args.seconds < 0:
            raise MonitorError("Confirmation delay cannot be negative")
        config["confirmation_delay_seconds"] = args.seconds
    else:
        raise MonitorError(f"Unknown configuration action: {action}")
    save_config(config)
    print(f"Configuration updated: {CONFIG_FILE}")
    return 0


def git_command(arguments: list[str], cwd: Path = SCRIPT_DIR, check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        ["git", *arguments], cwd=cwd, capture_output=True, text=True, timeout=300, check=False
    )
    if check and result.returncode != 0:
        raise MonitorError(normalize_text(result.stderr or result.stdout) or f"git {' '.join(arguments)} failed")
    return result


def run_update(check_only: bool = False) -> int:
    if git_command(["status", "--porcelain"]).stdout.strip():
        raise MonitorError("The repository has local changes. Commit or preserve them before updating.")
    branch = git_command(["branch", "--show-current"]).stdout.strip()
    if not branch:
        raise MonitorError("The repository is in detached-HEAD state")
    git_command(["fetch", "origin", branch])
    local = git_command(["rev-parse", "HEAD"]).stdout.strip()
    remote_ref = f"origin/{branch}"
    remote = git_command(["rev-parse", remote_ref]).stdout.strip()
    if local == remote:
        print(f"Already up to date on {branch} ({local[:8]}).")
        return 0
    ancestor = git_command(["merge-base", "--is-ancestor", local, remote_ref], check=False)
    if ancestor.returncode != 0:
        raise MonitorError(f"Local {branch} has diverged from {remote_ref}; automatic update was stopped")
    if check_only:
        print(f"Update available: {local[:8]} → {remote[:8]}")
        return 0
    with tempfile.TemporaryDirectory(prefix="mdot-monitor-update-") as parent:
        candidate = Path(parent) / "candidate"
        git_command(["worktree", "add", "--detach", str(candidate), remote_ref])
        try:
            tests = subprocess.run(
                [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
                cwd=candidate,
                capture_output=True,
                text=True,
                timeout=300,
                check=False,
            )
            if tests.returncode != 0:
                raise MonitorError("Candidate update failed its tests:\n" + (tests.stdout + tests.stderr)[-4000:])
        finally:
            git_command(["worktree", "remove", "--force", str(candidate)], check=False)
    git_command(["merge", "--ff-only", remote_ref])
    print(f"Updated {branch}: {local[:8]} → {remote[:8]}; candidate tests passed.")
    print("The existing scheduled task will use the updated files on its next run.")
    return 0


def send_test() -> int:
    config = load_config(require_recipients=True)
    recipients = list(config["recipients"])
    send_outlook(
        "[MDOT Standards] Test email",
        format_status_email(
            "MDOT standards monitor test",
            "Email delivery is configured successfully. This is only a test; no website update was detected.",
            str(config["page_url"]),
        ),
        recipients,
    )
    print(f"Test email sent to {len(recipients)} recipient(s).")
    return 0


def send_change_preview() -> int:
    """Send an explicitly labeled example of a real change notification."""
    config = load_config(require_recipients=True)
    recipients = list(config["recipients"])
    baseline = read_json(STATE_FILE, {}) or {}
    real_documents = list(baseline.get("documents", []))

    def real_example(index: int) -> dict[str, object]:
        if index < len(real_documents):
            return real_documents[index]
        return {
            "url": str(config["page_url"]),
            "title": "MDOT Engineering Standards page",
            "sha256": f"fallback-{index}",
        }

    added_example = real_example(0)
    modified_example = real_example(1)
    renamed_example = real_example(2)
    changes = {
        "page_text_changed": False,
        "documents_added": [
            {
                "url": added_example["url"],
                "title": f'Example newly added document — {added_example["title"]}',
                "section": "Design Specifications and Manuals",
                "sha256": "example-new",
            }
        ],
        "documents_removed": [],
        "documents_modified": [
            {
                "old": {"url": modified_example["url"], "title": modified_example["title"], "sha256": "example-old"},
                "new": {"url": modified_example["url"], "title": f'Example modified content — {modified_example["title"]}', "sha256": "example-updated"},
                "details": [
                    "Page 7 changed — Removed: Minimum thickness 6 inches | Added: Minimum thickness 8 inches"
                ],
            }
        ],
        "documents_renamed": [
            {
                "old": {"url": renamed_example["url"], "title": renamed_example["title"], "sha256": "example-updated"},
                "new": {"url": renamed_example["url"], "title": f'Example renamed document — {renamed_example["title"]}', "sha256": "example-updated"},
            }
        ],
        "documents_relinked": [],
        "documents_moved": [],
        "links_added": [],
        "links_removed": [],
        "links_renamed": [],
    }
    snapshot = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "page_url": str(config["page_url"]),
    }
    body = format_change_email(changes, snapshot)
    banner = (
        '<tr><td style="padding:14px 20px;background-color:#fff3cd;border-bottom:1px solid #d39e00;'
        'font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:20px;color:#604b00;">'
        '<strong>TEST PREVIEW ONLY — No real MDOT change was detected.</strong><br>'
        'The items below are fictional examples.</td></tr>'
    )
    body = body.replace("<!-- LOCAL-TEST-BANNER -->", banner, 1)
    send_outlook(
        "[MDOT Standards] TEST PREVIEW — Example change notification",
        body,
        recipients,
    )
    print(f"Change-preview email sent to {len(recipients)} recipient(s).")
    return 0


def run_local_test(output: Path, open_report: bool = False) -> int:
    """Exercise the monitoring pipeline against a temporary local MDOT-like site."""
    from functools import partial
    from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
    import threading
    import webbrowser

    try:
        import pymupdf
    except ImportError as exc:
        raise MonitorError("PyMuPDF is required for the local end-to-end test") from exc

    class QuietHandler(SimpleHTTPRequestHandler):
        def log_message(self, format: str, *args: object) -> None:
            pass

    def write_pdf(path: Path, requirement: str) -> None:
        with pymupdf.open() as document:
            cover = document.new_page()
            cover.insert_text((72, 72), "Local test design manual")
            requirement_page = document.new_page()
            requirement_page.insert_text((72, 72), requirement)
            document.save(path)

    def write_page(path: Path, requirement: str, manual_title: str, include_bulletin: bool) -> None:
        bulletin = '<a href="/bulletin.txt">New construction bulletin</a>' if include_bulletin else ""
        path.write_text(
            "<html><body><main><h1>Engineering Standards</h1>"
            f"<p>{html_module.escape(requirement)}</p>"
            f'<a href="/manual.pdf">{html_module.escape(manual_title)}</a>'
            f"{bulletin}</main></body></html>",
            encoding="utf-8",
        )

    logger = logging.getLogger(APP_NAME + ".local-test")
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False

    with tempfile.TemporaryDirectory(prefix="mdot-monitor-local-test-") as directory:
        site = Path(directory)
        page_path = site / "index.html"
        manual_path = site / "manual.pdf"
        bulletin_path = site / "bulletin.txt"
        handler = partial(QuietHandler, directory=str(site))
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        server_thread = threading.Thread(target=server.serve_forever, daemon=True)
        server_thread.start()
        base_url = f"http://127.0.0.1:{server.server_port}"
        page_url = base_url + "/index.html"
        try:
            write_page(page_path, "Minimum pavement thickness is 6 inches.", "2025 Design Manual", False)
            write_pdf(manual_path, "Minimum pavement thickness: 6 inches")
            old_snapshot = build_snapshot(
                page_url,
                logger,
                force_analysis=True,
                extra_documents=[{
                    "url": base_url + "/manual.pdf",
                    "title": "2025 Design Manual",
                    "section": "Local test documents",
                }],
                preview_root=site / "previews",
            )

            write_page(page_path, "Minimum pavement thickness is 8 inches.", "2026 Design Manual", True)
            write_pdf(manual_path, "Minimum pavement thickness: 8 inches")
            bulletin_path.write_text("New construction bulletin\nEffective immediately\n", encoding="utf-8")
            new_snapshot = build_snapshot(
                page_url,
                logger,
                previous_snapshot=old_snapshot,
                force_analysis=True,
                extra_documents=[
                    {
                        "url": base_url + "/manual.pdf",
                        "title": "2026 Design Manual",
                        "section": "Local test documents",
                    },
                    {
                        "url": base_url + "/bulletin.txt",
                        "title": "New construction bulletin",
                        "section": "Local test documents",
                    },
                ],
                preview_root=site / "previews",
            )
            changes = compare_snapshots(old_snapshot, new_snapshot)
            # Materialize temporary thumbnails before the temporary site is removed.
            for item in changes.get("documents_modified", []):
                for preview in item.get("previews", []):
                    preview["old_path"] = preview_image_source(str(preview["old_path"]), "data")
                    preview["new_path"] = preview_image_source(str(preview["new_path"]), "data")
        finally:
            server.shutdown()
            server.server_close()
            server_thread.join(timeout=5)

    modified_details = [
        str(detail)
        for item in changes["documents_modified"]
        for detail in item.get("details", [])
    ]
    expected_results = {
        "visible page text change": bool(changes["page_text_changed"]),
        "added document": len(changes["documents_added"]) == 1,
        "modified PDF": len(changes["documents_modified"]) == 1,
        "renamed document": len(changes["documents_renamed"]) == 1,
        "PDF page-level detail": any("Page 2 changed" in detail for detail in modified_details),
    }
    failed = [name for name, passed in expected_results.items() if not passed]
    if failed:
        raise MonitorError("Local test did not detect: " + ", ".join(failed))

    report = format_change_email(changes, new_snapshot, image_mode="data")
    banner = (
        '<tr><td style="padding:14px 20px;background-color:#d9f0f5;border-bottom:1px solid #8ab8c5;'
        'font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:20px;color:#123f4d;">'
        '<strong>LOCAL END-TO-END TEST — PASS</strong><br>'
        'Generated from a temporary fictional website. No MDOT state was changed and no email was sent.'
        '</td></tr>'
    )
    report = report.replace("<!-- LOCAL-TEST-BANNER -->", banner, 1)
    output = output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(report, encoding="utf-8")
    print(f"Local end-to-end test passed: {change_count(changes)} change(s) detected.")
    print(f"HTML report: {output}")
    print("No Outlook message was sent, no scheduled task was created, and saved monitor state was not changed.")
    if open_report:
        webbrowser.open(output.as_uri())
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    check = subparsers.add_parser("check", help="Run the normal scheduled check")
    check.add_argument("--dry-run", action="store_true", help="Check without changing state or sending email")
    check.add_argument("--no-confirm", action="store_true", help="Skip the configured second confirmation check")
    initialize = subparsers.add_parser("initialize", help="Replace the baseline without sending an update email")
    initialize.add_argument("--dry-run", action="store_true", help=argparse.SUPPRESS)
    initialize.add_argument("--no-confirm", action="store_true", help=argparse.SUPPRESS)
    subparsers.add_parser("send-test", help="Send a test email through Outlook")
    subparsers.add_parser("send-preview", help="Send a labeled example change notification")
    status = subparsers.add_parser("status", help="Show monitor, dependency, baseline, and scheduled-task health")
    status.add_argument("--json", action="store_true", help="Print machine-readable JSON")
    history = subparsers.add_parser("history", help="Build or open the local change-history dashboard")
    history.add_argument("--open", action="store_true", help="Open the dashboard in the default browser")
    update = subparsers.add_parser("update", help="Safely fast-forward to a tested update from GitHub")
    update.add_argument("--check-only", action="store_true", help="Fetch and report whether an update is available")

    config_parser = subparsers.add_parser("config", help="View or update monitor configuration")
    config_actions = config_parser.add_subparsers(dest="config_action", required=True)
    config_actions.add_parser("show", help="Print the effective configuration")
    config_actions.add_parser("validate", help="Validate the configuration file")
    for action_name in ("add-recipient", "remove-recipient", "set-failure-recipient"):
        action_parser = config_actions.add_parser(action_name)
        action_parser.add_argument("email")
    set_filter = config_actions.add_parser("set-filter", help="Replace global document filters")
    set_filter.add_argument("--section", action="append", help="Section-name substring; repeatable")
    set_filter.add_argument("--title", action="append", help="Document-title substring; repeatable")
    set_filter.add_argument("--extension", action="append", help="File extension such as pdf; repeatable")
    config_actions.add_parser("clear-filters", help="Monitor all sections, titles, and file types")
    add_route = config_actions.add_parser("add-route", help="Add or replace a filtered recipient route")
    add_route.add_argument("name")
    add_route.add_argument("--recipients", required=True, help="Semicolon-separated email addresses")
    add_route.add_argument("--section", action="append")
    add_route.add_argument("--title", action="append")
    add_route.add_argument("--extension", action="append")
    add_route.add_argument("--include-page-text", action="store_true")
    remove_route = config_actions.add_parser("remove-route", help="Remove a recipient route by name")
    remove_route.add_argument("name")
    heartbeat = config_actions.add_parser("set-heartbeat-days", help="Set weekly-summary interval; zero disables")
    heartbeat.add_argument("days", type=float)
    confirmation = config_actions.add_parser("set-confirmation-delay", help="Set seconds before confirming a change")
    confirmation.add_argument("seconds", type=float)
    local_test = subparsers.add_parser(
        "local-test",
        help="Run an end-to-end test locally without MDOT, Outlook, or Task Scheduler",
    )
    local_test.add_argument(
        "--output",
        type=Path,
        default=BASE_DIR / "local-test-report.html",
        help="HTML report path (default: %(default)s)",
    )
    local_test.add_argument("--open", action="store_true", help="Open the generated report in the default browser")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command in {"status", "history", "config", "update"}:
        try:
            if args.command == "status":
                return run_status(args.json)
            if args.command == "history":
                return run_history(args.open)
            if args.command == "config":
                return run_config_command(args)
            return run_update(args.check_only)
        except Exception as exc:
            print(f"{args.command.title()} failed: {exc}", file=sys.stderr)
            return 1
    if args.command == "local-test":
        try:
            return run_local_test(args.output, args.open)
        except Exception as exc:
            print(f"Local end-to-end test failed: {exc}", file=sys.stderr)
            return 1
    if args.command in {"send-test", "send-preview"}:
        try:
            return send_test() if args.command == "send-test" else send_change_preview()
        except Exception as exc:
            setup_logging().exception("Test email failed: %s", exc)
            return 1
    args.initialize = args.command == "initialize"
    return run_check(args)


if __name__ == "__main__":
    raise SystemExit(main())
