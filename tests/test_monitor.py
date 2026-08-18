import io
import tempfile
import unittest
from unittest import mock
from urllib.error import URLError
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


def snapshot(text="Page", documents=None, links=None):
    return {
        "page_text": text,
        "documents": documents or [],
        "links": links or [],
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

    def test_page_and_link_changes(self):
        old = snapshot("Old", links=[{"url": "https://x/a", "title": "A"}])
        new = snapshot("New", links=[{"url": "https://x/a", "title": "Renamed"}, {"url": "https://x/b", "title": "B"}])
        changes = monitor.compare_snapshots(old, new)
        self.assertTrue(changes["page_text_changed"])
        self.assertEqual(1, len(changes["links_added"]))
        self.assertEqual(1, len(changes["links_renamed"]))


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


if __name__ == "__main__":
    unittest.main()
