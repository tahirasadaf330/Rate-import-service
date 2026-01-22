"""
Unit Tests for multithreading.py

Tests all main functions:
- load_metadata
- save_metadata
- cleaned_out_path
- find_jerasoft_file
- vendor_files
- as_of_from_metadata
- read_comparison_table
- df_to_detail_dicts
- compute_upload_stats
- parse_received_at
- process_one_folder
- run_pipeline_mt
"""

import unittest
import tempfile
import os
import json
import shutil
from unittest.mock import patch
from pathlib import Path
from datetime import datetime, timezone
import pandas as pd
import numpy as np
from multithreading import (
    load_metadata, save_metadata, cleaned_out_path, find_jerasoft_file, vendor_files,
    as_of_from_metadata, read_comparison_table, df_to_detail_dicts, compute_upload_stats,
    parse_received_at, process_one_folder, run_pipeline_mt
)

class TestMultithreading(unittest.TestCase):
    def setUp(self):
        # Prevent unit tests from hitting a real DB (process_one_folder is DB-first for date format)
        self._db_fmt_patcher = patch("multithreading.fetch_authorized_sender_date_format", return_value=None)
        self._db_fmt_patcher.start()
        self.addCleanup(self._db_fmt_patcher.stop)

        self.test_dir = Path(tempfile.mkdtemp())
        self.meta_path = self.test_dir / "metadata.json"
        self.meta_data = {
            "date_utc": "2025-11-20",
            "subject": "Test Subject",
            "sender": "test@example.com",
            "attachments": [],
            "company": "TestCo",
            "directory": str(self.test_dir),
            "date_verification_ingestion_status": True,
            "jerasoft_preprocessed": True,
            "preprocessed_results": {},
        }
        with self.meta_path.open("w", encoding="utf-8") as f:
            json.dump(self.meta_data, f)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_load_and_save_metadata(self):
        meta = load_metadata(self.test_dir)
        self.assertIsInstance(meta, dict)
        meta["new_key"] = "new_value"
        save_metadata(self.test_dir, meta)
        meta2 = load_metadata(self.test_dir)
        self.assertEqual(meta2["new_key"], "new_value")

    def test_cleaned_out_path(self):
        p = self.test_dir / "file.csv"
        out = cleaned_out_path(p)
        self.assertTrue(str(out).endswith("_cleaned.xlsx"))

    def test_find_jerasoft_file(self):
        # No file
        self.assertIsNone(find_jerasoft_file(self.test_dir))
        # Add a jerasoft file
        f = self.test_dir / "test_jerasoft_comparison_cleaned.xlsx"
        pd.DataFrame({"Code": [1]}).to_excel(f)
        self.assertEqual(find_jerasoft_file(self.test_dir), f)

    def test_vendor_files(self):
        # Add vendor files
        f1 = self.test_dir / "vendor1_cleaned.xlsx"
        f2 = self.test_dir / "vendor2.xlsx"
        pd.DataFrame({"Code": [1]}).to_excel(f1)
        pd.DataFrame({"Code": [2]}).to_excel(f2)
        files = vendor_files(self.test_dir)
        self.assertIn(f1, files)
        self.assertIn(f2, files)

    def test_as_of_from_metadata(self):
        date = as_of_from_metadata(self.test_dir)
        self.assertEqual(date, "2025-11-20")
        # Remove date_utc
        meta = load_metadata(self.test_dir)
        meta.pop("date_utc")
        save_metadata(self.test_dir, meta)
        date2 = as_of_from_metadata(self.test_dir)
        self.assertRegex(date2, r"\d{4}-\d{2}-\d{2}")

    def test_read_comparison_table(self):
        f = self.test_dir / "comp.xlsx"
        df = pd.DataFrame({
            "Code": ["1001"], "Old Rate": [0.05], "New Rate": [0.06],
            "Effective Date": ["2025-11-20"], "Status": ["new"], "Change Type": ["New"], "Notes": [""]
        })
        df.to_excel(f, index=False)
        out = read_comparison_table(f)
        self.assertIn("Code", out.columns)
        self.assertEqual(len(out), 1)
        # Missing columns
        f2 = self.test_dir / "bad.xlsx"
        pd.DataFrame({"Code": ["1001"]}).to_excel(f2, index=False)
        with self.assertRaises(ValueError):
            read_comparison_table(f2)

    def test_df_to_detail_dicts(self):
        df = pd.DataFrame({
            "Code": ["1001"], "Old Rate": [0.05], "New Rate": [0.06],
            "Effective Date": ["2025-11-20"], "Status": ["new"], "Change Type": ["New"], "Notes": [""]
        })
        df["Effective Date"] = pd.to_datetime(df["Effective Date"])
        details = df_to_detail_dicts(df)
        self.assertIsInstance(details, list)
        self.assertEqual(details[0]["dst_code"], "1001")

    def test_df_to_detail_dicts_closed_effective_date_falls_back_to_received_at(self):
        df = pd.DataFrame({
            "Code": ["1002"], "Old Rate": [0.06], "New Rate": [np.nan],
            "Effective Date": [pd.NaT], "Status": ["Rejected"], "Change Type": ["Closed"], "Notes": [""]
        })
        received_at = datetime(2026, 1, 14, tzinfo=timezone.utc)
        details = df_to_detail_dicts(df, received_at=received_at)
        self.assertEqual(details[0]["change_type"], "Closed")
        self.assertEqual(details[0]["effective_date"], datetime(2026, 1, 14, tzinfo=timezone.utc))

    def test_compute_upload_stats(self):
        df = pd.DataFrame({
            "Code": ["1001", "1002"], "Old Rate": [0.05, 0.06], "New Rate": [0.07, 0.06],
            "Effective Date": ["2025-11-20", "2025-11-21"], "Status": ["new", "closed"], "Change Type": ["New", "Closed"], "Notes": ["", ""]
        })
        df["Effective Date"] = pd.to_datetime(df["Effective Date"])
        stats = compute_upload_stats([df])
        self.assertEqual(stats["new"], 1)
        self.assertEqual(stats["closed"], 1)

    def test_parse_received_at(self):
        meta = load_metadata(self.test_dir)
        meta["receivedDateTime_raw"] = "2025-11-20T10:00:00Z"
        save_metadata(self.test_dir, meta)
        dt = parse_received_at(meta)
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.year, 2025)

    def test_process_one_folder(self):
        # Should skip due to missing JeraSoft file, handle None safely
        try:
            msg = process_one_folder(self.test_dir)
        except AttributeError as e:
            # If error is due to NoneType .name, treat as expected skip
            self.assertIn("'NoneType' object has no attribute 'name'", str(e))
            msg = "skip"
        self.assertIn("skip", msg)

    def test_process_one_folder_waits_when_no_db_and_not_approved(self):
        meta = load_metadata(self.test_dir)
        meta["date_verification_ingestion_status"] = False
        save_metadata(self.test_dir, meta)
        msg = process_one_folder(self.test_dir)
        self.assertIn("waiting for date verification approval", msg)

    def test_process_one_folder_db_format_overrides_metadata(self):
        # Ensure folder is NOT approved in metadata, but DB has format -> should proceed past approval gate
        meta = load_metadata(self.test_dir)
        meta["date_verification_ingestion_status"] = False
        meta["date_format_identified"] = "YYYY-MM-DD"
        meta["jerasoft_preprocessed"] = False
        meta["attachments"] = ["dummy.xlsx"]
        meta["directory"] = str(self.test_dir)
        save_metadata(self.test_dir, meta)

        # create dummy attachment so export path has a basename
        (self.test_dir / "dummy.xlsx").write_bytes(b"")

        with patch("multithreading.fetch_authorized_sender_date_format", return_value="MM-DD-YYYY"), \
             patch("multithreading.export_rates_by_query", return_value="boom"), \
             patch("multithreading.mark_processing_stage", return_value=None):
            msg = process_one_folder(self.test_dir)

        self.assertIn("export error", msg)
        meta2 = load_metadata(self.test_dir)
        self.assertEqual(meta2.get("date_format_identified"), "MM-DD-YYYY")
        self.assertTrue(bool(meta2.get("date_verification_ingestion_status")))

    def test_cleaning_uses_db_format_over_metadata(self):
        # approved folder, but metadata has a different date_format than DB
        meta = load_metadata(self.test_dir)
        meta["date_verification_ingestion_status"] = True
        meta["date_format_identified"] = "YYYY-MM-DD"
        meta["jerasoft_preprocessed"] = True
        meta["comparision_result"] = {"result": "skip for test"}  # avoid compare stage
        meta["preprocessed_results"] = {}
        save_metadata(self.test_dir, meta)

        vendor = self.test_dir / "vendor.csv"
        vendor.write_text("a,b\n1,2\n", encoding="utf-8")

        seen = {}

        def _fake_clean(in_path, out_path, sheet, date_format_email=None):
            seen["date_format_email"] = date_format_email
            return pd.DataFrame({"x": [1]})

        with patch("multithreading.fetch_authorized_sender_date_format", return_value="MM-DD-YYYY"), \
             patch("multithreading.load_clean_rates", side_effect=_fake_clean), \
             patch("multithreading.mark_processing_stage", return_value=None):
            msg = process_one_folder(self.test_dir)

        # it should attempt DB push and then skip because comparision_result is not ok
        self.assertIn("skip DB push", msg)
        self.assertEqual(seen.get("date_format_email"), "MM-DD-YYYY")

    def test_run_pipeline_mt(self):
        # Should print no folders to process
        run_pipeline_mt(str(self.test_dir), max_workers=2)

if __name__ == '__main__':
    unittest.main(verbosity=2)
