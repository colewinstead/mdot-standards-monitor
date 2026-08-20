"""Daily monitor for the MDOT Engineering Standards/Guides/Manuals page."""

from __future__ import annotations

import argparse
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
LOCK_FILE = BASE_DIR / "monitor.lock"
LOG_FILE = BASE_DIR / "monitor.log"
SCRIPT_DIR = Path(__file__).resolve().parent
EMAIL_HELPER = SCRIPT_DIR / "send_outlook.ps1"


class MonitorError(RuntimeError):
    """An expected monitoring failure."""


class MainContentParser(HTMLParser):
    """Extract normalized main-page text, headings, and links."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_main = False
        self.suppressed_depth = 0
        self.text_parts: list[str] = []
        self.links: list[dict[str, str]] = []
        self.current_link: dict[str, object] | None = None
        self.current_heading: list[str] | None = None
        self.section = ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        attr_map = dict(attrs)
        if tag == "main" and not self.in_main:
            self.in_main = True
            return
        if not self.in_main:
            return
        if tag in {"script", "style", "svg", "noscript"}:
            self.suppressed_depth += 1
        if self.suppressed_depth:
            return
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            self.current_heading = []
        if tag == "a" and attr_map.get("href"):
            self.current_link = {"href": attr_map["href"] or "", "text": []}

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if not self.in_main:
            return
        if tag == "main":
            self.in_main = False
            return
        if tag in {"script", "style", "svg", "noscript"} and self.suppressed_depth:
            self.suppressed_depth -= 1
        if not self.suppressed_depth:
            if tag == "a" and self.current_link is not None:
                link_text = normalize_text(" ".join(self.current_link["text"]))
                self.links.append(
                    {
                        "url": str(self.current_link["href"]),
                        "title": link_text,
                        "section": self.section,
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
        self.text_parts.append(text)
        if self.current_heading is not None:
            self.current_heading.append(text)
        if self.current_link is not None:
            self.current_link["text"].append(text)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\u200c", " ").replace("\xa0", " ")).strip()


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
    return {"text": normalize_text(" ".join(parser.text_parts)), "links": links}


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
    with tempfile.TemporaryDirectory(prefix="mdot-monitor-chrome-") as profile:
        command = [
            str(chrome),
            "--headless=new",
            "--disable-gpu",
            "--disable-extensions",
            "--disable-background-networking",
            "--no-first-run",
            "--no-default-browser-check",
            f"--user-data-dir={profile}",
            "--virtual-time-budget=20000",
            "--dump-dom",
            url,
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_seconds,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            raise MonitorError(f"Chrome did not finish rendering within {timeout_seconds} seconds") from exc
    if result.returncode != 0 or "<main" not in result.stdout.lower():
        detail = normalize_text(result.stderr)[-500:]
        raise MonitorError(f"Chrome could not render the MDOT page (exit {result.returncode}): {detail}")
    return result.stdout


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
                "content_type": response.headers.get_content_type(),
                "content_disposition": response.headers.get("Content-Disposition", ""),
                "etag": response.headers.get("ETag", ""),
                "last_modified": response.headers.get("Last-Modified", ""),
                "_temp_path": str(temp_path),
            }
    except (HTTPError, URLError, TimeoutError, OSError) as exc:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        raise MonitorError(f"Could not download {item['title']} ({item['url']}): {exc}") from exc


def normalized_lines(value: str) -> list[str]:
    return [normalize_text(line) for line in value.splitlines() if normalize_text(line)]


def analyze_pdf(path: Path) -> dict[str, object]:
    try:
        import pymupdf
    except ImportError as exc:
        raise MonitorError("PyMuPDF is required for page-level PDF comparisons") from exc

    pages: list[dict[str, object]] = []
    try:
        with pymupdf.open(path, filetype="pdf") as document:
            for index, page in enumerate(document):
                text = "\n".join(normalized_lines(page.get_text("text", sort=True)))
                # A low-resolution grayscale rendering catches drawings and scanned pages
                # that contain little or no extractable text.
                pixmap = page.get_pixmap(dpi=48, colorspace=pymupdf.csGRAY, alpha=False, annots=True)
                pages.append(
                    {
                        "page": index + 1,
                        "text": text,
                        "text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
                        "visual_sha256": hashlib.sha256(pixmap.samples).hexdigest(),
                    }
                )
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
) -> dict[str, object]:
    extension = infer_document_extension(url, content_type, content_disposition)
    if extension == ".pdf":
        return analyze_pdf(path)
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


def build_snapshot(
    page_url: str,
    logger: logging.Logger,
    previous_snapshot: dict[str, object] | None = None,
    force_analysis: bool = False,
    extra_documents: list[dict[str, str]] | None = None,
    dynamic_documents: list[dict[str, str]] | None = None,
) -> dict[str, object]:
    rendered = render_page(page_url)
    page = parse_rendered_page(rendered, page_url)
    document_links = [item for item in page["links"] if is_mdot_document(item["url"])]
    known_urls = {item["url"] for item in document_links}
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
        resolved = resolve_dynamic_document(rule)
        if resolved["url"] not in known_urls:
            document_links.append(resolved)
            known_urls.add(resolved["url"])
            logger.info("Resolved dynamic document %s to %s", resolved["title"], resolved["url"])
    logger.info("Rendered page with %d links and %d monitored documents", len(page["links"]), len(document_links))
    previous_documents = {
        item.get("identity", item["url"]): item
        for item in (previous_snapshot or {}).get("documents", [])
    }
    previous_by_url = {
        item["url"]: item for item in (previous_snapshot or {}).get("documents", [])
    }
    documents = []
    for index, item in enumerate(document_links, start=1):
        logger.info("Hashing document %d/%d: %s", index, len(document_links), item["title"])
        downloaded = download_document(item)
        temp_path = Path(str(downloaded.pop("_temp_path")))
        try:
            previous = previous_documents.get(item.get("identity", item["url"])) or previous_by_url.get(item["url"])
            if (
                not force_analysis
                and previous
                and previous.get("sha256") == downloaded["sha256"]
                and previous.get("analysis")
            ):
                downloaded["analysis"] = previous["analysis"]
            else:
                logger.info("Creating detailed content snapshot: %s", item["title"])
                downloaded["analysis"] = analyze_document(
                    temp_path,
                    item["url"],
                    str(downloaded.get("content_type", "")),
                    str(downloaded.get("content_disposition", "")),
                )
            documents.append(downloaded)
        finally:
            temp_path.unlink(missing_ok=True)
    documents.sort(key=lambda item: item["url"].casefold())
    return {
        "schema_version": 1,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "page_url": page_url,
        "page_text": page["text"],
        "links": page["links"],
        "documents": documents,
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
    old_docs = {item.get("identity", item["url"]): item for item in old.get("documents", [])}
    new_docs = {item.get("identity", item["url"]): item for item in new.get("documents", [])}
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
        }
        for url in sorted(set(old_docs) & set(new_docs))
        if old_docs[url].get("sha256") != new_docs[url].get("sha256")
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

    def link_map(snapshot: dict[str, object]) -> dict[str, str]:
        return {item["url"]: item.get("title", "") for item in snapshot.get("links", [])}

    old_links, new_links = link_map(old), link_map(new)
    link_added = [{"url": url, "title": new_links[url]} for url in sorted(set(new_links) - set(old_links))]
    link_removed = [{"url": url, "title": old_links[url]} for url in sorted(set(old_links) - set(new_links))]
    link_renamed = [
        {"url": url, "old_title": old_links[url], "new_title": new_links[url]}
        for url in sorted(set(old_links) & set(new_links))
        if old_links[url] != new_links[url]
    ]
    return {
        "page_text_changed": old.get("page_text") != new.get("page_text"),
        "page_text_details": describe_page_text_changes(
            str(old.get("page_text", "")),
            str(new.get("page_text", "")),
        ),
        "documents_added": [new_docs[url] for url in sorted(added_urls)],
        "documents_removed": [old_docs[url] for url in sorted(removed_urls)],
        "documents_modified": modified,
        "documents_renamed": renamed,
        "documents_relinked": relinked,
        "documents_moved": moved,
        "links_added": link_added,
        "links_removed": link_removed,
        "links_renamed": link_renamed,
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


def format_change_email(changes: dict[str, object], snapshot: dict[str, object]) -> str:
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
        return f'{link(new_document["url"], new_document["title"])}{detail_html}'

    add_section("Documents added", changes["documents_added"], lambda x: link(x["url"], x["title"]))
    add_section("Documents removed", changes["documents_removed"], lambda x: e(str(x["title"])))
    add_section("Document contents modified", changes["documents_modified"], modified_document)
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
    return config


def send_outlook(subject: str, body_html: str, recipients: list[str]) -> None:
    if not EMAIL_HELPER.exists():
        raise MonitorError(f"Outlook helper is missing: {EMAIL_HELPER}")
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".html", delete=False) as body_file:
        body_file.write(body_html)
        body_path = Path(body_file.name)
    try:
        command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(EMAIL_HELPER), "-Recipients", ";".join(recipients),
            "-Subject", subject, "-HtmlBodyFile", str(body_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=90, check=False)
        if result.returncode != 0:
            detail = normalize_text(result.stderr or result.stdout)
            raise MonitorError(f"Outlook could not send the email: {detail}")
    finally:
        body_path.unlink(missing_ok=True)


def default_runtime() -> dict[str, object]:
    return {"consecutive_failures": 0, "last_error": ""}


def record_failure(error: Exception, config: dict[str, object], logger: logging.Logger) -> None:
    runtime = read_json(RUNTIME_FILE, default_runtime()) or default_runtime()
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


def clear_failure_state(config: dict[str, object], logger: logging.Logger) -> None:
    atomic_write_json(RUNTIME_FILE, default_runtime())


def run_check(args: argparse.Namespace) -> int:
    logger = setup_logging()
    config = load_config(require_recipients=False)
    with FileLock():
        try:
            baseline = read_json(STATE_FILE)
            snapshot = build_snapshot(
                str(config["page_url"]),
                logger,
                previous_snapshot=baseline,
                force_analysis=False,
                extra_documents=list(config.get("extra_documents", [])),
                dynamic_documents=list(config.get("dynamic_documents", [])),
            )
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
                clear_failure_state(config, logger)
                logger.info("Baseline initialized with %d documents; no update email sent", len(snapshot["documents"]))
                return 0

            changes = compare_snapshots(baseline, snapshot)
            if not has_changes(changes):
                atomic_write_json(STATE_FILE, snapshot)
                clear_failure_state(config, logger)
                logger.info("No MDOT updates detected")
                return 0

            recipients = list(load_config(require_recipients=True)["recipients"])
            subject = f"[MDOT Standards] {change_count(changes)} update(s) detected"
            send_outlook(subject, format_change_email(changes, snapshot), recipients)
            atomic_write_json(STATE_FILE, snapshot)
            clear_failure_state(config, logger)
            logger.info("Sent MDOT update email to %d recipient(s)", len(recipients))
            return 0
        except Exception as exc:
            logger.exception("Monitor check failed: %s", exc)
            if not args.dry_run:
                record_failure(exc, config, logger)
            return 1


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
            )
        finally:
            server.shutdown()
            server.server_close()
            server_thread.join(timeout=5)

    changes = compare_snapshots(old_snapshot, new_snapshot)
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

    report = format_change_email(changes, new_snapshot)
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
    initialize = subparsers.add_parser("initialize", help="Replace the baseline without sending an update email")
    initialize.add_argument("--dry-run", action="store_true", help=argparse.SUPPRESS)
    subparsers.add_parser("send-test", help="Send a test email through Outlook")
    subparsers.add_parser("send-preview", help="Send a labeled example change notification")
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
