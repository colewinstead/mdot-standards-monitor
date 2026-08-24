import argparse
import io
import tempfile
import unittest
from unittest import mock
from urllib.error import URLError
import zipfile
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import monitor


class PageParsingTests(unittest.TestCase):
    def test_extracts_main_text_links_and_documents(self):
        source = """
        <html><body><header><a href='/noise.pdf'>Noise</a></header>
        <main><h2>Manuals</h2><p>  Useful   text </p>
        <a href='/documents/Test%20Manual.pdf'>Test Manual</a>
        <a href='https://example.com/guide'>External Guide</a></main></body></html>
        """
        result = monitor.parse_rendered_page(source, monitor.DEFAULT_URL)
        self.assertIn("Manuals Useful text", result["text"])
        self.assertEqual(2, len(result["links"]))
        self.assertTrue(all(item["section"] == "Manuals" for item in result["links"]))
        by_host = {monitor.is_mdot_document(item["url"]) for item in result["links"]}
        self.assertEqual({True, False}, by_host)

    def test_ignores_fragments_when_canonicalizing(self):
        result = monitor.canonicalize_url("https://MDOT.ms.gov/portal/page", "/documents/a.pdf#page=2")
        self.assertEqual("https://mdot.ms.gov/documents/a.pdf", result)

    def test_encodes_spaces_in_mdot_urls(self):
        result = monitor.canonicalize_url(
            monitor.DEFAULT_URL,
            "/documents/Bridge Design/MDOT Manual.pdf",
        )
        self.assertEqual(
            "https://mdot.ms.gov/documents/Bridge%20Design/MDOT%20Manual.pdf",
            result,
        )

    def test_rejects_page_without_main_content(self):
        with self.assertRaises(monitor.MonitorError):
            monitor.parse_rendered_page("<html><body>Loading</body></html>", monitor.DEFAULT_URL)

    def test_extracts_expanded_folder_paths_and_assigns_them_to_documents(self):
        source = """
        <main><table>
        <tr class='dx-group-row' data-mdot-folder-path='Roadway Design / Standards'>
          <td>Standards</td></tr>
        <tr class='dx-group-row' data-mdot-folder-path='Roadway Design / Standards / Manuals'>
          <td>Manuals</td></tr>
        <tr class='dx-data-row'><td>
          <a data-mdot-folder-path='Roadway Design / Standards / Manuals'
             href='/documents/Roadway Design/Standards/Manuals/manual.pdf'>Road Manual</a>
        </td></tr></table></main>
        """
        result = monitor.parse_rendered_page(source, monitor.DEFAULT_URL)
        self.assertEqual(
            ["Roadway Design / Standards", "Roadway Design / Standards / Manuals"],
            [item["path"] for item in result["folders"]],
        )
        self.assertEqual("Roadway Design / Standards / Manuals", result["links"][0]["section"])


class HashTests(unittest.TestCase):
    def test_hash_stream(self):
        digest, size = monitor.hash_stream(io.BytesIO(b"abc"))
        self.assertEqual("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", digest)
        self.assertEqual(3, size)

    def test_inaccessible_document_is_a_monitor_failure(self):
        item = {"url": "https://mdot.ms.gov/documents/missing.pdf", "title": "Missing", "section": "Manuals"}
        with mock.patch("monitor.urlopen", side_effect=URLError("offline")):
            with self.assertRaises(monitor.MonitorError):
                monitor.download_document(item)

    def test_projectwise_download_is_identified_as_pdf_from_headers(self):
        extension = monitor.infer_document_extension(
            "https://pwdocs.mdot.state.ms.us/Resources/Services/ProjectWise/Download.ashx/View?key=123",
            "application/pdf",
            'inline; filename="RWD_Workflow_Training_ORD.pdf"',
        )
        self.assertEqual(".pdf", extension)


def snapshot(text="Page", documents=None, links=None, folders=None):
    return {
        "page_text": text,
        "documents": documents or [],
        "links": links or [],
        "folders": folders or [],
    }


def document(url, title, digest):
    return {"url": url, "title": title, "sha256": digest, "section": "Manuals", "size": 3}


class ComparisonTests(unittest.TestCase):
    def test_unchanged(self):
        value = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "1")])
        changes = monitor.compare_snapshots(value, value)
        self.assertFalse(monitor.has_changes(changes))

    def test_added_removed_modified_and_renamed(self):
        old = snapshot(documents=[
            document("https://mdot.ms.gov/documents/remove.pdf", "Remove", "r"),
            document("https://mdot.ms.gov/documents/change.pdf", "Old title", "old"),
        ])
        new = snapshot(documents=[
            document("https://mdot.ms.gov/documents/add.pdf", "Add", "a"),
            document("https://mdot.ms.gov/documents/change.pdf", "New title", "new"),
        ])
        changes = monitor.compare_snapshots(old, new)
        self.assertEqual(1, len(changes["documents_added"]))
        self.assertEqual(1, len(changes["documents_removed"]))
        self.assertEqual(1, len(changes["documents_modified"]))
        self.assertEqual(1, len(changes["documents_renamed"]))

    def test_same_content_at_new_url_is_move(self):
        old = snapshot(documents=[document("https://mdot.ms.gov/documents/old.pdf", "Old", "same")])
        new = snapshot(documents=[document("https://mdot.ms.gov/documents/new.pdf", "New", "same")])
        changes = monitor.compare_snapshots(old, new)
        self.assertEqual(1, len(changes["documents_moved"]))
        self.assertFalse(changes["documents_added"])
        self.assertFalse(changes["documents_removed"])

    def test_dynamic_document_url_change_is_relinked_not_removed(self):
        old_document = document("https://pwdocs.example/old", "Training", "same")
        old_document["identity"] = "dynamic:Training"
        new_document = document("https://pwdocs.example/new", "Training", "same")
        new_document["identity"] = "dynamic:Training"
        changes = monitor.compare_snapshots(
            snapshot(documents=[old_document]),
            snapshot(documents=[new_document]),
        )
        self.assertEqual(1, len(changes["documents_relinked"]))
        self.assertFalse(changes["documents_added"])
        self.assertFalse(changes["documents_removed"])

    def test_page_and_link_changes(self):
        old = snapshot("Old", links=[{"url": "https://x/a", "title": "A"}])
        new = snapshot("New", links=[{"url": "https://x/a", "title": "Renamed"}, {"url": "https://x/b", "title": "B"}])
        changes = monitor.compare_snapshots(old, new)
        self.assertTrue(changes["page_text_changed"])
        self.assertEqual(1, len(changes["links_added"]))
        self.assertEqual(1, len(changes["links_renamed"]))
        self.assertTrue(changes["page_text_details"])
        self.assertEqual(3, monitor.change_count(changes))

    def test_folder_added_and_removed(self):
        old = snapshot(folders=[{
            "path": "Roadway Design / Standards / Old Folder",
            "title": "Old Folder", "section": "Roadway Design / Standards / Old Folder",
            "url": monitor.DEFAULT_URL,
        }])
        new = snapshot(folders=[{
            "path": "Roadway Design / Standards / New Folder",
            "title": "New Folder", "section": "Roadway Design / Standards / New Folder",
            "url": monitor.DEFAULT_URL,
        }])
        changes = monitor.compare_snapshots(old, new)
        self.assertEqual("Roadway Design / Standards / New Folder", changes["folders_added"][0]["path"])
        self.assertEqual("Roadway Design / Standards / Old Folder", changes["folders_removed"][0]["path"])

    def test_legacy_snapshot_does_not_report_every_existing_folder_as_added(self):
        old = snapshot()
        old.pop("folders")
        new = snapshot(folders=[{
            "path": "Roadway Design / Standards", "title": "Standards",
            "section": "Roadway Design / Standards", "url": monitor.DEFAULT_URL,
        }])
        changes = monitor.compare_snapshots(old, new)
        self.assertFalse(changes["folders_added"])

    def test_folder_crawl_upgrade_baselines_newly_discovered_content(self):
        old = snapshot("Collapsed", links=[{
            "url": "https://mdot.ms.gov/portal/obsolete_standards",
            "title": "Obsolete Standards",
            "section": "",
        }])
        old["schema_version"] = 1
        new_doc = document("https://mdot.ms.gov/documents/Roadway%20Design/new.pdf", "New", "digest")
        new = snapshot(
            "Expanded",
            documents=[new_doc],
            links=[{"url": new_doc["url"], "title": new_doc["title"], "section": "Roadway Design"}],
            folders=[{
                "path": "Roadway Design", "title": "Roadway Design",
                "section": "Roadway Design", "url": monitor.DEFAULT_URL,
            }],
        )
        new["schema_version"] = 2
        changes = monitor.compare_snapshots(old, new)
        self.assertFalse(changes["page_text_changed"])
        self.assertFalse(changes["documents_added"])
        self.assertFalse(changes["folders_added"])
        self.assertFalse(changes["links_added"])
        self.assertFalse(changes["links_removed"])
        self.assertFalse(changes["links_renamed"])
        self.assertFalse(monitor.has_changes(changes))

    def test_excluded_construction_documents_do_not_appear_removed(self):
        construction = document(
            "https://mdot.ms.gov/documents/Construction/Specifications/spec.pdf",
            "Specification", "old",
        )
        changes = monitor.compare_snapshots(snapshot(documents=[construction]), snapshot())
        self.assertFalse(changes["documents_removed"])
        self.assertFalse(monitor.has_changes(changes))


class DetailedComparisonTests(unittest.TestCase):
    def test_pdf_reports_changed_page_number_and_text(self):
        import pymupdf

        with tempfile.TemporaryDirectory() as directory:
            old_path = Path(directory) / "old.pdf"
            new_path = Path(directory) / "new.pdf"
            for path, second_page_text in ((old_path, "Original requirement"), (new_path, "Revised requirement")):
                document_file = pymupdf.open()
                first = document_file.new_page()
                first.insert_text((72, 72), "Unchanged cover")
                second = document_file.new_page()
                second.insert_text((72, 72), second_page_text)
                document_file.save(path)
                document_file.close()

            details = monitor.describe_pdf_changes(
                monitor.analyze_pdf(old_path),
                monitor.analyze_pdf(new_path),
            )
            self.assertTrue(any("Page 2 changed" in detail for detail in details))
            self.assertTrue(any("Original requirement" in detail for detail in details))
            self.assertTrue(any("Revised requirement" in detail for detail in details))

    def test_dynamic_pdf_link_resolver_selects_only_matching_target(self):
        import pymupdf

        target = "https://pwdocs.mdot.state.ms.us/Resources/Services/ProjectWise/Download.ashx/View?key=new-key"
        document_file = pymupdf.open()
        page = document_file.new_page()
        page.insert_link({
            "kind": pymupdf.LINK_URI,
            "from": pymupdf.Rect(10, 10, 200, 30),
            "uri": target,
        })
        pdf_bytes = document_file.tobytes()
        document_file.close()
        rule = {
            "title": "RWD Workflow Training ORD",
            "source_url": "https://mdot.ms.gov/manual.pdf",
            "match_host": "pwdocs.mdot.state.ms.us",
            "match_path_contains": "/ProjectWise/Download.ashx/View",
        }
        with mock.patch("monitor.urlopen", return_value=io.BytesIO(pdf_bytes)):
            resolved = monitor.resolve_dynamic_document(rule)
        self.assertEqual(target, resolved["url"])
        self.assertEqual("dynamic:RWD Workflow Training ORD", resolved["identity"])

    def test_word_paragraph_changes_are_described(self):
        details = monitor.describe_analysis_changes(
            {"analysis": {"kind": "paragraphs", "paragraphs": ["Keep", "Old paragraph"]}},
            {"analysis": {"kind": "paragraphs", "paragraphs": ["Keep", "New paragraph"]}},
        )
        self.assertTrue(any("Paragraph(s) 2 changed" in detail for detail in details))
        self.assertTrue(any("Old paragraph" in detail and "New paragraph" in detail for detail in details))

    def test_excel_reports_sheet_and_cell(self):
        details = monitor.describe_excel_changes(
            {"cells": [{"sheet": "Pay Items", "cell": "B7", "value": "100"}]},
            {"cells": [{"sheet": "Pay Items", "cell": "B7", "value": "125"}]},
        )
        self.assertEqual(["Pay Items!B7 changed: 100 → 125"], details)

    def test_tolerant_excel_parser_extracts_shared_string_and_formula(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsm"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr(
                    "xl/workbook.xml",
                    """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <sheets><sheet name="Pay Items" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                )
                archive.writestr(
                    "xl/_rels/workbook.xml.rels",
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>""",
                )
                archive.writestr(
                    "xl/sharedStrings.xml",
                    """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                    <si><t>Bridge item</t></si></sst>""",
                )
                archive.writestr(
                    "xl/worksheets/sheet1.xml",
                    """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                    <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>
                    <c r="B1"><f>SUM(1,2)</f><v>3</v></c></row></sheetData></worksheet>""",
                )
            analysis = monitor.analyze_excel(path)
            self.assertIn({"sheet": "Pay Items", "cell": "A1", "value": "Bridge item"}, analysis["cells"])
            self.assertIn({"sheet": "Pay Items", "cell": "B1", "value": "=SUM(1,2)"}, analysis["cells"])

    def test_email_includes_detailed_change_lines(self):
        old = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "old")])
        new = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "new")])
        changes = monitor.compare_snapshots(old, new)
        changes["documents_modified"][0]["details"] = ["Page 7 changed — Added: revised value"]
        new.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
        body = monitor.format_change_email(changes, new)
        self.assertIn("Page 7 changed", body)
        self.assertIn("revised value", body)

    def test_page_text_email_shows_old_and_new_wording(self):
        old = snapshot(text="Design manual effective January 2025")
        new = snapshot(text="Design manual effective July 2026")
        changes = monitor.compare_snapshots(old, new)
        new.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
        body = monitor.format_change_email(changes, new)
        self.assertIn("January 2025", body)
        self.assertIn("July 2026", body)
        self.assertEqual(1, monitor.change_count(changes))


class StorageAndEmailTests(unittest.TestCase):
    def test_atomic_json_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            monitor.atomic_write_json(path, {"ok": True})
            self.assertEqual({"ok": True}, monitor.read_json(path))

    def test_email_escapes_untrusted_titles(self):
        old = snapshot()
        new = snapshot(
            documents=[document("https://mdot.ms.gov/documents/a.pdf", "<script>alert(1)</script>", "x")]
        )
        changes = monitor.compare_snapshots(old, new)
        new.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
        body = monitor.format_change_email(changes, new)
        self.assertNotIn("<script>", body)
        self.assertIn("&lt;script&gt;", body)

    def test_change_preview_is_clearly_labeled(self):
        real_urls = [
            "https://mdot.ms.gov/documents/one.pdf",
            "https://mdot.ms.gov/documents/two.pdf",
            "https://mdot.ms.gov/documents/three.pdf",
        ]
        preview_state = {
            "documents": [
                {"url": url, "title": f"Real document {index}", "sha256": str(index)}
                for index, url in enumerate(real_urls, start=1)
            ]
        }
        with mock.patch("monitor.load_config", return_value={
            "recipients": ["team@example.com"],
            "page_url": monitor.DEFAULT_URL,
        }), mock.patch("monitor.read_json", return_value=preview_state), mock.patch("monitor.send_outlook") as sender:
            monitor.send_change_preview()
            subject, body, recipients = sender.call_args.args
            self.assertIn("TEST PREVIEW", subject)
            self.assertIn("No real MDOT change was detected", body)
            self.assertIn("fictional examples", body)
            self.assertEqual(["team@example.com"], recipients)
            self.assertTrue(all(url in body for url in real_urls))

    def test_local_test_is_available_without_configuration(self):
        parser = monitor.build_parser()
        args = parser.parse_args(["local-test"])
        self.assertEqual("local-test", args.command)
        self.assertEqual(monitor.BASE_DIR / "local-test-report.html", args.output)
        self.assertFalse(args.open)

    def test_recipient_config_rejects_non_list(self):
        old_path = monitor.CONFIG_FILE
        with tempfile.TemporaryDirectory() as directory:
            try:
                monitor.CONFIG_FILE = Path(directory) / "config.json"
                monitor.atomic_write_json(monitor.CONFIG_FILE, {"recipients": "not-a-list"})
                with self.assertRaises(monitor.MonitorError):
                    monitor.load_config()
            finally:
                monitor.CONFIG_FILE = old_path

    def test_config_rejects_malformed_extra_document(self):
        old_path = monitor.CONFIG_FILE
        with tempfile.TemporaryDirectory() as directory:
            try:
                monitor.CONFIG_FILE = Path(directory) / "config.json"
                monitor.atomic_write_json(
                    monitor.CONFIG_FILE,
                    {"recipients": [], "extra_documents": [{"url": "https://example.com/file.pdf"}]},
                )
                with self.assertRaises(monitor.MonitorError):
                    monitor.load_config()
            finally:
                monitor.CONFIG_FILE = old_path

    def test_failure_notice_is_sent_only_to_failure_recipient(self):
        old_runtime = monitor.RUNTIME_FILE
        with tempfile.TemporaryDirectory() as directory:
            try:
                monitor.RUNTIME_FILE = Path(directory) / "runtime.json"
                logger = mock.Mock()
                config = {
                    "recipients": ["team1@example.com", "team2@example.com"],
                    "failure_recipient": "owner@example.com",
                    "page_url": monitor.DEFAULT_URL,
                }
                with mock.patch("monitor.send_outlook") as sender:
                    monitor.record_failure(RuntimeError("first"), config, logger)
                    sender.assert_called_once()
                    self.assertEqual(["owner@example.com"], sender.call_args.args[2])
                runtime = monitor.read_json(monitor.RUNTIME_FILE)
                self.assertEqual(1, runtime["consecutive_failures"])
            finally:
                monitor.RUNTIME_FILE = old_runtime

    def test_failure_without_failure_recipient_is_logged_only(self):
        old_runtime = monitor.RUNTIME_FILE
        with tempfile.TemporaryDirectory() as directory:
            try:
                monitor.RUNTIME_FILE = Path(directory) / "runtime.json"
                with mock.patch("monitor.send_outlook") as sender:
                    monitor.record_failure(
                        RuntimeError("offline"),
                        {"recipients": ["team@example.com"], "page_url": monitor.DEFAULT_URL},
                        mock.Mock(),
                    )
                    sender.assert_not_called()
            finally:
                monitor.RUNTIME_FILE = old_runtime

    def test_recovery_state_clears_without_sending_email(self):
        old_runtime = monitor.RUNTIME_FILE
        with tempfile.TemporaryDirectory() as directory:
            try:
                monitor.RUNTIME_FILE = Path(directory) / "runtime.json"
                monitor.atomic_write_json(
                    monitor.RUNTIME_FILE,
                    {"consecutive_failures": 2, "last_error": "offline"},
                )
                with mock.patch("monitor.send_outlook") as sender:
                    monitor.clear_failure_state(
                        {"recipients": ["team@example.com"]},
                        mock.Mock(),
                    )
                    sender.assert_not_called()
                self.assertEqual(monitor.default_runtime(), monitor.read_json(monitor.RUNTIME_FILE))
            finally:
                monitor.RUNTIME_FILE = old_runtime


class FeatureTests(unittest.TestCase):
    def test_weekly_health_summary_uses_only_failure_recipient(self):
        original_paths = (
            monitor.STATE_FILE, monitor.RUNTIME_FILE, monitor.LOCK_FILE,
            monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR,
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            try:
                monitor.STATE_FILE = root / "state.json"
                monitor.RUNTIME_FILE = root / "runtime.json"
                monitor.LOCK_FILE = root / "monitor.lock"
                monitor.HISTORY_FILE = root / "history.json"
                monitor.HISTORY_DASHBOARD = root / "history.html"
                monitor.HISTORY_REPORTS_DIR = root / "reports"
                current = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "same")])
                current.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
                monitor.atomic_write_json(monitor.STATE_FILE, current)
                runtime = monitor.default_runtime()
                runtime.update({"last_success_at": "2020-01-01T00:00:00+00:00", "last_heartbeat_at": "2020-01-01T00:00:00+00:00"})
                monitor.atomic_write_json(monitor.RUNTIME_FILE, runtime)
                config = {
                    "page_url": monitor.DEFAULT_URL,
                    "recipients": ["whole-team@example.com"],
                    "failure_recipient": "owner@example.com",
                    "extra_documents": [], "dynamic_documents": [], "filters": {}, "recipient_rules": [],
                    "retry_attempts": 1, "retry_delay_seconds": 0,
                    "confirmation_delay_seconds": 0, "heartbeat_days": 7, "history_limit": 10,
                }
                args = argparse.Namespace(dry_run=False, initialize=False, no_confirm=False)
                with mock.patch("monitor.load_config", return_value=config), \
                     mock.patch("monitor.build_snapshot", return_value=current), \
                     mock.patch("monitor.setup_logging", return_value=mock.Mock()), \
                     mock.patch("monitor.send_outlook") as sender, \
                     mock.patch("monitor.cleanup_preview_cache"):
                    self.assertEqual(0, monitor.run_check(args))
                sender.assert_called_once()
                self.assertEqual(["owner@example.com"], sender.call_args.args[2])
            finally:
                (
                    monitor.STATE_FILE, monitor.RUNTIME_FILE, monitor.LOCK_FILE,
                    monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR,
                ) = original_paths

    def test_old_pdf_analysis_without_previews_is_rebuilt_after_upgrade(self):
        old_pdf_analysis = {"kind": "pdf_pages", "pages": [{"page": 1, "visual_sha256": "x"}]}
        self.assertFalse(monitor.analysis_can_be_reused(old_pdf_analysis, store_previews=True))
        self.assertTrue(monitor.analysis_can_be_reused(old_pdf_analysis, store_previews=False))
        self.assertTrue(monitor.analysis_can_be_reused({"kind": "lines", "lines": ["x"]}, store_previews=True))

    def test_retry_operation_recovers_from_temporary_failure(self):
        operation = mock.Mock(side_effect=[monitor.MonitorError("temporary"), "ok"])
        logger = mock.Mock()
        with mock.patch("monitor.time.sleep") as sleeper:
            result = monitor.retry_operation(operation, 3, 2, logger, "test")
        self.assertEqual("ok", result)
        sleeper.assert_called_once_with(2)
        self.assertIn("temporary", logger.warning.call_args.args)

    def test_snapshot_retries_only_the_failed_document(self):
        page = {
            "text": "Standards",
            "folders": [],
            "links": [{
                "url": "https://mdot.ms.gov/documents/a.pdf",
                "title": "A",
                "section": "Manuals",
            }],
        }
        with tempfile.TemporaryDirectory() as directory:
            downloaded_path = Path(directory) / "a.pdf"
            downloaded_path.write_bytes(b"pdf")
            downloaded = {
                "url": page["links"][0]["url"],
                "title": "A",
                "section": "Manuals",
                "sha256": "digest",
                "size": 3,
                "content_type": "application/pdf",
                "content_disposition": "",
                "_temp_path": str(downloaded_path),
            }
            with mock.patch("monitor.render_page", return_value="rendered") as renderer, \
                 mock.patch("monitor.parse_rendered_page", return_value=page), \
                 mock.patch("monitor.download_document", side_effect=[monitor.MonitorError("timeout"), downloaded]) as downloader, \
                 mock.patch("monitor.analyze_document", return_value={"kind": "lines", "lines": []}), \
                 mock.patch("monitor.time.sleep"):
                result = monitor.build_snapshot(
                    monitor.DEFAULT_URL,
                    mock.Mock(),
                    store_previews=False,
                    retry_attempts=3,
                    retry_delay_seconds=0,
                )
        self.assertEqual(1, renderer.call_count)
        self.assertEqual(2, downloader.call_count)
        self.assertEqual(1, len(result["documents"]))

    def test_snapshot_retains_last_good_document_after_all_retries_fail(self):
        item = {
            "url": "https://mdot.ms.gov/documents/document.pdf",
            "title": "Training",
            "section": "Dynamic",
            "identity": "dynamic:Training",
        }
        previous = dict(item, sha256="last-good", size=123, analysis={"kind": "lines", "lines": ["known"]})
        page = {"text": "Standards", "folders": [], "links": [item]}
        with mock.patch("monitor.render_page", return_value="rendered") as renderer, \
             mock.patch("monitor.parse_rendered_page", return_value=page), \
             mock.patch("monitor.download_document", side_effect=monitor.MonitorError("HTTP Error 504")) as downloader, \
             mock.patch("monitor.time.sleep"):
            result = monitor.build_snapshot(
                monitor.DEFAULT_URL,
                mock.Mock(),
                previous_snapshot=snapshot(documents=[previous]),
                store_previews=False,
                retry_attempts=3,
                retry_delay_seconds=0,
            )
        self.assertEqual(1, renderer.call_count)
        self.assertEqual(3, downloader.call_count)
        self.assertEqual("last-good", result["documents"][0]["sha256"])
        self.assertEqual("Training", result["document_errors"][0]["title"])

    def test_document_warning_goes_only_to_failure_recipient(self):
        added = document("https://mdot.ms.gov/documents/new.pdf", "New standard", "new")
        current = snapshot(documents=[added])
        current.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
        current["document_errors"] = [{"title": "Training", "error": "HTTP Error 504"}]
        config = {"failure_recipient": "owner@example.com", "page_url": monitor.DEFAULT_URL}
        with mock.patch("monitor.send_outlook") as sender:
            monitor.notify_document_errors(current, config, mock.Mock())
        sender.assert_called_once()
        self.assertEqual(["owner@example.com"], sender.call_args.args[2])
        self.assertIn("HTTP Error 504", sender.call_args.args[1])
        team_body = monitor.format_change_email(
            monitor.compare_snapshots(snapshot(), current),
            current,
        )
        self.assertIn("New standard", team_body)
        self.assertNotIn("HTTP Error 504", team_body)

    def test_document_filters_match_section_title_and_extension(self):
        item = {
            "title": "Bridge Design Manual",
            "section": "Bridge Standards",
            "url": "https://mdot.ms.gov/documents/bridge/manual.pdf",
        }
        self.assertTrue(monitor.item_matches_filters(item, {
            "sections": ["bridge"], "titles": ["design"], "extensions": ["pdf"],
        }))
        self.assertFalse(monitor.item_matches_filters(item, {"sections": ["roadway"]}))

    def test_filtered_recipient_route_receives_only_matching_changes(self):
        bridge = document("https://mdot.ms.gov/documents/bridge.pdf", "Bridge Manual", "b")
        bridge["section"] = "Bridge"
        roadway = document("https://mdot.ms.gov/documents/road.pdf", "Road Manual", "r")
        roadway["section"] = "Roadway"
        changes = monitor.compare_snapshots(snapshot(), snapshot(documents=[bridge, roadway]))
        batches = monitor.notification_batches(changes, {
            "recipients": [],
            "recipient_rules": [{
                "name": "Bridge team",
                "recipients": ["bridge@example.com"],
                "filters": {"sections": ["Bridge"]},
            }],
        })
        self.assertEqual(1, len(batches))
        self.assertEqual(["bridge@example.com"], batches[0][0])
        self.assertEqual([bridge], batches[0][1]["documents_added"])

    def test_pdf_comparison_exposes_before_and_after_previews(self):
        import pymupdf

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old_path, new_path = root / "old.pdf", root / "new.pdf"
            for path, value in ((old_path, "Six inches"), (new_path, "Eight inches")):
                with pymupdf.open() as pdf:
                    page = pdf.new_page()
                    page.insert_text((72, 72), value)
                    pdf.save(path)
            old_doc = document("https://mdot.ms.gov/documents/a.pdf", "Manual", "old")
            new_doc = document("https://mdot.ms.gov/documents/a.pdf", "Manual", "new")
            old_doc["analysis"] = monitor.analyze_pdf(old_path, root / "old-previews")
            new_doc["analysis"] = monitor.analyze_pdf(new_path, root / "new-previews")
            changes = monitor.compare_snapshots(snapshot(documents=[old_doc]), snapshot(documents=[new_doc]))
            self.assertEqual(1, len(changes["documents_modified"][0]["previews"]))
            current = snapshot(documents=[new_doc])
            current.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
            body = monitor.format_change_email(changes, current, image_mode="data")
            self.assertIn("BEFORE — PAGE 1", body)
            self.assertIn("data:image/png;base64,", body)

    def test_history_writes_dashboard_and_change_report(self):
        old_paths = (monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            try:
                monitor.HISTORY_FILE = root / "history.json"
                monitor.HISTORY_DASHBOARD = root / "history.html"
                monitor.HISTORY_REPORTS_DIR = root / "reports"
                changes = monitor.compare_snapshots(snapshot(), snapshot(documents=[
                    document("https://mdot.ms.gov/documents/a.pdf", "New Manual", "new")
                ]))
                current = snapshot()
                current.update({"generated_at": "now", "page_url": monitor.DEFAULT_URL})
                event = monitor.record_history("change", "One change", ["1 document added"], 10, changes, current)
                self.assertTrue(Path(str(event["report"])).is_file())
                self.assertIn("One change", monitor.HISTORY_DASHBOARD.read_text(encoding="utf-8"))
            finally:
                monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR = old_paths

    def test_confirmation_suppresses_transient_change_without_email(self):
        original_paths = (
            monitor.STATE_FILE, monitor.RUNTIME_FILE, monitor.LOCK_FILE,
            monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR,
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            try:
                monitor.STATE_FILE = root / "state.json"
                monitor.RUNTIME_FILE = root / "runtime.json"
                monitor.LOCK_FILE = root / "monitor.lock"
                monitor.HISTORY_FILE = root / "history.json"
                monitor.HISTORY_DASHBOARD = root / "history.html"
                monitor.HISTORY_REPORTS_DIR = root / "reports"
                baseline = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "old")])
                baseline.update({"generated_at": "old", "page_url": monitor.DEFAULT_URL})
                changed = snapshot(documents=[document("https://mdot.ms.gov/documents/a.pdf", "A", "new")])
                changed.update({"generated_at": "new", "page_url": monitor.DEFAULT_URL})
                monitor.atomic_write_json(monitor.STATE_FILE, baseline)
                config = {
                    "page_url": monitor.DEFAULT_URL, "recipients": ["team@example.com"],
                    "failure_recipient": "", "extra_documents": [], "dynamic_documents": [],
                    "filters": {}, "recipient_rules": [], "retry_attempts": 1,
                    "retry_delay_seconds": 0, "confirmation_delay_seconds": 1,
                    "heartbeat_days": 0, "history_limit": 10,
                }
                args = argparse.Namespace(dry_run=False, initialize=False, no_confirm=False)
                with mock.patch("monitor.load_config", return_value=config), \
                     mock.patch("monitor.build_snapshot", side_effect=[changed, baseline]), \
                     mock.patch("monitor.setup_logging", return_value=mock.Mock()), \
                     mock.patch("monitor.time.sleep"), \
                     mock.patch("monitor.send_outlook") as sender, \
                     mock.patch("monitor.cleanup_preview_cache"):
                    self.assertEqual(0, monitor.run_check(args))
                    sender.assert_not_called()
                events = monitor.read_json(monitor.HISTORY_FILE)
                self.assertEqual("suppressed", events[0]["kind"])
            finally:
                (
                    monitor.STATE_FILE, monitor.RUNTIME_FILE, monitor.LOCK_FILE,
                    monitor.HISTORY_FILE, monitor.HISTORY_DASHBOARD, monitor.HISTORY_REPORTS_DIR,
                ) = original_paths

    def test_new_commands_are_available(self):
        parser = monitor.build_parser()
        self.assertEqual("status", parser.parse_args(["status"]).command)
        self.assertEqual("history", parser.parse_args(["history"]).command)
        self.assertEqual("update", parser.parse_args(["update", "--check-only"]).command)
        config_args = parser.parse_args(["config", "set-heartbeat-days", "7"])
        self.assertEqual("set-heartbeat-days", config_args.config_action)


if __name__ == "__main__":
    unittest.main()
