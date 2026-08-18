"""Daily monitor for the MDOT Engineering Standards/Guides/Manuals page."""

from __future__ import annotations

import argparse
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
from datetime import datetime
from html.parser import HTMLParser
from typing import BinaryIO, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urljoin, urlsplit, urlunsplit
from urllib.request import Request, urlopen


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
    try:
        with urlopen(request, timeout=timeout_seconds) as response:
            digest, size = hash_stream(response)
            return {
                "url": item["url"],
                "title": item["title"],
                "section": item["section"],
                "sha256": digest,
                "size": size,
                "content_type": response.headers.get_content_type(),
                "etag": response.headers.get("ETag", ""),
                "last_modified": response.headers.get("Last-Modified", ""),
            }
    except (HTTPError, URLError, TimeoutError, OSError) as exc:
        raise MonitorError(f"Could not download {item['title']} ({item['url']}): {exc}") from exc


def build_snapshot(page_url: str, logger: logging.Logger) -> dict[str, object]:
    rendered = render_page(page_url)
    page = parse_rendered_page(rendered, page_url)
    document_links = [item for item in page["links"] if is_mdot_document(item["url"])]
    logger.info("Rendered page with %d links and %d MDOT documents", len(page["links"]), len(document_links))
    documents = []
    for index, item in enumerate(document_links, start=1):
        logger.info("Hashing document %d/%d: %s", index, len(document_links), item["title"])
        documents.append(download_document(item))
    documents.sort(key=lambda item: item["url"].casefold())
    return {
        "schema_version": 1,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "page_url": page_url,
        "page_text": page["text"],
        "links": page["links"],
        "documents": documents,
    }


def compare_snapshots(old: dict[str, object], new: dict[str, object]) -> dict[str, object]:
    old_docs = {item["url"]: item for item in old.get("documents", [])}
    new_docs = {item["url"]: item for item in new.get("documents", [])}
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
        {"old": old_docs[url], "new": new_docs[url]}
        for url in sorted(set(old_docs) & set(new_docs))
        if old_docs[url].get("sha256") != new_docs[url].get("sha256")
    ]
    renamed = [
        {"old": old_docs[url], "new": new_docs[url]}
        for url in sorted(set(old_docs) & set(new_docs))
        if old_docs[url].get("title") != new_docs[url].get("title")
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
        "documents_added": [new_docs[url] for url in sorted(added_urls)],
        "documents_removed": [old_docs[url] for url in sorted(removed_urls)],
        "documents_modified": modified,
        "documents_renamed": renamed,
        "documents_moved": moved,
        "links_added": link_added,
        "links_removed": link_removed,
        "links_renamed": link_renamed,
    }


def has_changes(changes: dict[str, object]) -> bool:
    return bool(changes["page_text_changed"] or any(changes[key] for key in changes if key != "page_text_changed"))


def change_count(changes: dict[str, object]) -> int:
    return int(bool(changes["page_text_changed"])) + sum(
        len(value) for key, value in changes.items() if key != "page_text_changed"
    )


def format_change_email(changes: dict[str, object], snapshot: dict[str, object]) -> str:
    e = html_module.escape
    sections: list[str] = []

    def doc_list(title: str, items: Iterable[dict[str, object]], formatter) -> None:
        rendered = list(items)
        if rendered:
            sections.append(f"<h3>{e(title)}</h3><ul>" + "".join(f"<li>{formatter(item)}</li>" for item in rendered) + "</ul>")

    doc_list("Documents added", changes["documents_added"], lambda x: f'<a href="{e(str(x["url"]))}">{e(str(x["title"]))}</a>')
    doc_list("Documents removed", changes["documents_removed"], lambda x: e(str(x["title"])))
    doc_list("Document contents modified", changes["documents_modified"], lambda x: f'<a href="{e(str(x["new"]["url"]))}">{e(str(x["new"]["title"]))}</a>')
    doc_list("Documents renamed", changes["documents_renamed"], lambda x: f'{e(str(x["old"]["title"]))} &rarr; <a href="{e(str(x["new"]["url"]))}">{e(str(x["new"]["title"]))}</a>')
    doc_list("Documents moved or replaced at a new URL", changes["documents_moved"], lambda x: f'{e(str(x["old"]["url"]))} &rarr; <a href="{e(str(x["new"]["url"]))}">{e(str(x["new"]["title"]))}</a>')
    doc_list("Links added", changes["links_added"], lambda x: f'<a href="{e(str(x["url"]))}">{e(str(x["title"]))}</a>')
    doc_list("Links removed", changes["links_removed"], lambda x: e(f'{x["title"]} ({x["url"]})'))
    doc_list("Link titles changed", changes["links_renamed"], lambda x: f'{e(str(x["old_title"]))} &rarr; <a href="{e(str(x["url"]))}">{e(str(x["new_title"]))}</a>')
    if changes["page_text_changed"]:
        sections.insert(0, "<h3>Page content changed</h3><p>Visible headings or explanatory text changed.</p>")
    checked = e(str(snapshot["generated_at"]))
    page_url = e(str(snapshot["page_url"]))
    return (
        "<html><body style='font-family:Segoe UI,Arial,sans-serif'>"
        "<h2>MDOT Engineering Standards update detected</h2>"
        f"<p>The daily check found {change_count(changes)} change(s) at {checked}.</p>"
        + "".join(sections)
        + f'<p><a href="{page_url}">Open the MDOT Engineering Standards/Guides/Manuals page</a></p>'
        "<p style='color:#666;font-size:9pt'>Automated by MDOT Standards Monitor.</p></body></html>"
    )


def format_status_email(title: str, message: str, page_url: str) -> str:
    return (
        "<html><body style='font-family:Segoe UI,Arial,sans-serif'>"
        f"<h2>{html_module.escape(title)}</h2><p>{html_module.escape(message)}</p>"
        f'<p><a href="{html_module.escape(page_url)}">Open the monitored MDOT page</a></p>'
        "</body></html>"
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
    recipients = config["recipients"]
    if not isinstance(recipients, list):
        raise MonitorError(f"recipients must be a list in {CONFIG_FILE}")
    if require_recipients and not recipients:
        raise MonitorError(f"No recipients are configured. Run install.ps1 first. Config: {CONFIG_FILE}")
    if not isinstance(config["failure_recipient"], str):
        raise MonitorError(f"failure_recipient must be a string in {CONFIG_FILE}")
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
            snapshot = build_snapshot(str(config["page_url"]), logger)
            baseline = read_json(STATE_FILE)
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
            }
        ],
        "documents_renamed": [
            {
                "old": {"url": renamed_example["url"], "title": renamed_example["title"], "sha256": "example-updated"},
                "new": {"url": renamed_example["url"], "title": f'Example renamed document — {renamed_example["title"]}', "sha256": "example-updated"},
            }
        ],
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
        "<div style='padding:12px;background:#fff3cd;border:1px solid #d39e00;"
        "font-weight:700'>TEST PREVIEW ONLY — No real MDOT change was detected. "
        "The items below are fictional examples.</div>"
    )
    body = body.replace("<h2>", banner + "<h2>", 1)
    send_outlook(
        "[MDOT Standards] TEST PREVIEW — Example change notification",
        body,
        recipients,
    )
    print(f"Change-preview email sent to {len(recipients)} recipient(s).")
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
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
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
