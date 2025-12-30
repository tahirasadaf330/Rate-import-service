import unittest
from unittest.mock import patch, MagicMock
import date_verification
import pandas as pd
from datetime import datetime, date, timezone
from pathlib import Path
import json
import tempfile
import shutil

class TestDateVerification(unittest.TestCase):
    def test_first_attachment_path(self):
        # Patch Path.is_absolute to simulate absolute and relative paths
        with patch.object(Path, 'is_absolute', return_value=True):
            meta = {'attachments': ['/abs/path/file.xlsx']}
            result = date_verification._first_attachment_path(meta)
            self.assertEqual(str(result).replace('\\', '/').replace('\\', '/'), '/abs/path/file.xlsx')
        with patch.object(Path, 'is_absolute', return_value=False):
            meta = {'attachments': ['file.xlsx'], 'directory': '/dir'}
            result = date_verification._first_attachment_path(meta)
            self.assertEqual(str(result).replace('\\', '/').replace('\\', '/'), '/dir/file.xlsx')
        meta = {'attachments': []}
        self.assertIsNone(date_verification._first_attachment_path(meta))
        meta = {'attachments': ['']}
        self.assertIsNone(date_verification._first_attachment_path(meta))
        meta = {}
        self.assertIsNone(date_verification._first_attachment_path(meta))

    def test_df_preview_records(self):
        df = pd.DataFrame({'A': [1, None], 'B': ['x', 'y']})
        preview = date_verification._df_preview_records(df, limit=1)
        self.assertIsInstance(preview, list)
        self.assertEqual(len(preview), 1)
        df_empty = pd.DataFrame()
        self.assertEqual(date_verification._df_preview_records(df_empty), [])

    def test_parse_iso_utc_dt(self):
        dt_str = '2025-11-20T12:00:00Z'
        dt = date_verification._parse_iso_utc_dt(dt_str)
        self.assertIsInstance(dt, datetime)
        self.assertIsNone(date_verification._parse_iso_utc_dt(None))
        self.assertIsNone(date_verification._parse_iso_utc_dt('bad'))
        self.assertIsNone(date_verification._parse_iso_utc_dt(123))

    def test_has_native_datetimes(self):
        df = pd.DataFrame({'A': [datetime.now(), datetime.now(), datetime.now(), datetime.now(), datetime.now()]})
        self.assertTrue(date_verification._has_native_datetimes(df))
        df = pd.DataFrame({'A': ['x', 'y', 'z']})
        self.assertFalse(date_verification._has_native_datetimes(df))

    def test_ingest_files_db_first_marks_meta_approved_and_inserts_ingest_row(self):
        root = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: shutil.rmtree(root, ignore_errors=True))

        folder = root / "f1"
        folder.mkdir(parents=True, exist_ok=True)
        meta_path = folder / "metadata.json"
        meta = {
            "sender": "rates@example.com",
            "subject": "subj",
            "directory": str(folder),
            "attachments": ["file.xlsx"],
        }
        meta_path.write_text(json.dumps(meta), encoding="utf-8")

        def _fake_upsert(**kwargs):
            # DB-first should approve using DB format
            self.assertEqual(kwargs.get("status"), "approved")
            self.assertEqual(kwargs.get("date_format"), "DD-MM-YYYY")
            self.assertFalse(bool(kwargs.get("is_format_auto_detected")))
            return 999

        with patch("date_verification.fetch_authorized_sender_date_format", return_value="DD-MM-YYYY"), \
             patch("date_verification.insert_or_update_ingest_file", side_effect=_fake_upsert) as mock_upsert, \
             patch("date_verification.mark_processing_stage", return_value=None):
            scanned, inserted, skipped = date_verification.ingest_files_for_manual_date(root)

        self.assertEqual(scanned, 1)
        self.assertEqual(inserted, 1)
        self.assertEqual(skipped, 0)
        self.assertTrue(mock_upsert.called)

        meta2 = json.loads(meta_path.read_text(encoding="utf-8"))
        self.assertTrue(meta2.get("date_verification_ingestion"))
        self.assertTrue(meta2.get("date_verification_ingestion_status"))
        self.assertEqual(meta2.get("date_format_identified"), "DD-MM-YYYY")

    def test_ingest_files_autodetect_sets_meta_and_ingest_row_when_no_db_format(self):
        root = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: shutil.rmtree(root, ignore_errors=True))

        folder = root / "f2"
        folder.mkdir(parents=True, exist_ok=True)
        meta_path = folder / "metadata.json"
        xlsx = folder / "file.xlsx"
        xlsx.write_bytes(b"")  # dummy placeholder (we patch excel read)

        meta = {
            "sender": "rates@example.com",
            "subject": "subj",
            "directory": str(folder),
            "attachments": [str(xlsx.name)],
            "receivedDateTime_raw": "2025-11-20T12:00:00Z",
            "processed_at_utc": "2025-11-20T12:00:01Z",
        }
        meta_path.write_text(json.dumps(meta), encoding="utf-8")

        df_native = pd.DataFrame({"A": [datetime.now(timezone.utc)] * 6})

        def _fake_upsert(**kwargs):
            # should pass approved fields for autodetected excel
            self.assertEqual(kwargs.get("date_format"), "YYYY-MM-DD")
            self.assertEqual(kwargs.get("status"), "approved")
            self.assertTrue(bool(kwargs.get("is_format_auto_detected")))
            return 123

        with patch("date_verification.fetch_authorized_sender_date_format", return_value=None), \
             patch("date_verification._read_excel_native", return_value=df_native), \
             patch("date_verification._has_native_datetimes", return_value=True), \
             patch("date_verification.insert_or_update_ingest_file", side_effect=_fake_upsert), \
             patch("date_verification.upsert_authorized_sender_date_format", return_value=True) as mock_auth_upsert, \
             patch("date_verification.mark_processing_stage", return_value=None):
            scanned, inserted, skipped = date_verification.ingest_files_for_manual_date(root)

        self.assertEqual(scanned, 1)
        self.assertEqual(inserted, 1)
        self.assertEqual(skipped, 0)

        meta2 = json.loads(meta_path.read_text(encoding="utf-8"))
        self.assertTrue(meta2.get("date_verification_ingestion"))
        self.assertTrue(meta2.get("date_verification_ingestion_status"))
        self.assertEqual(meta2.get("date_format_identified"), "YYYY-MM-DD")
        self.assertTrue(mock_auth_upsert.called)


if __name__ == '__main__':
    unittest.main()
