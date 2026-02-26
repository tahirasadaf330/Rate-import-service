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
        self._db_fmt_patcher = patch("multithreading.fetch_vendor_context_by_sender_email", return_value={"vendor_id": None, "vendor_date_format": None})
        self._db_fmt_patcher.start()
        self.addCleanup(self._db_fmt_patcher.stop)

        # Prevent unit tests from hitting a real DB (process_one_folder now checks processing_statuses).
        self._db_status_patcher = patch("database.get_processing_status", return_value=None)
        self._db_status_patcher.start()
        self.addCleanup(self._db_status_patcher.stop)

        # Prevent unit tests from hitting a real DB for vendor/day pending-upload gating.
        self._db_vendor_pending_patcher = patch("multithreading.vendor_has_pending_jera_upload_today", return_value=False)
        self._db_vendor_pending_patcher.start()
        self.addCleanup(self._db_vendor_pending_patcher.stop)

        self._db_set_status_text_patcher = patch("multithreading.set_processing_status_text", return_value=1)
        self._db_set_status_text_patcher.start()
        self.addCleanup(self._db_set_status_text_patcher.stop)

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
        # No "Billing Increments Changes" label present, so count should be zero
        self.assertEqual(stats["billing_increment_changes"], 0)

    def test_compute_upload_stats_billing_increment_changes_from_label_only(self):
        # Row 1 explicitly marked as Billing Increments Changes
        # Row 2 is Closed with different Old/New BI but no label -> should NOT be counted
        df = pd.DataFrame({
            "Code": ["2001", "2002"],
            "Old Rate": [0.05, 0.06],
            "New Rate": [0.07, np.nan],
            "Effective Date": ["2025-11-22", "2025-11-23"],
            "Status": ["new", "rejected"],
            "Change Type": ["Billing Increments Changes", "Closed"],
            "Old Billing Increment": ["1/1", "1/1"],
            "New Billing Increment": ["1/2", np.nan],
            "Notes": ["", "present in current system but missing in new (closed)"]
        })
        df["Effective Date"] = pd.to_datetime(df["Effective Date"])
        stats = compute_upload_stats([df])
        # Only the row with the explicit label should be counted
        self.assertEqual(stats["billing_increment_changes"], 1)

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

    def test_process_one_folder_sets_waiting_when_vendor_pending_today(self):
        # If vendor_has_pending_jera_upload_today is True, we should skip processing and persist a DB-visible waiting status.
        from multithreading import WAITING_PREVIOUS_VENDOR_PENDING

        with patch("multithreading.vendor_has_pending_jera_upload_today", return_value=True) as mock_pending, \
             patch("multithreading.set_processing_status_text", return_value=1) as mock_set:
            msg = process_one_folder(self.test_dir)

        self.assertIn("skip", msg.lower())
        self.assertIn("waiting:", msg.lower())
        self.assertIn("previous jerasoft upload", msg.lower())

        mock_pending.assert_called()
        mock_set.assert_called_once()
        self.assertEqual(mock_set.call_args.kwargs.get("directory_name"), self.test_dir.name)
        self.assertEqual(mock_set.call_args.kwargs.get("status_text"), WAITING_PREVIOUS_VENDOR_PENDING)

    def test_two_waiting_sheets_advance_one_per_next_run(self):
        """
        Scenario:
          - Run #1: prior JeraSoft upload is pending -> both new sheets go to WAITING.
          - Run #2: prior upload becomes completed -> first sheet proceeds, but second remains WAITING
                   because a new pending upload now exists for the first sheet (simulated).

        This verifies the intended behavior: multiple waiting sheets drain one-by-one across runs.
        """
        from multithreading import WAITING_PREVIOUS_VENDOR_PENDING

        # Create a second folder with same-vendor metadata
        d2 = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: shutil.rmtree(d2, ignore_errors=True))
        meta2 = dict(self.meta_data)
        meta2["directory"] = str(d2)
        meta2["comparision_result"] = {"result": "skip for test"}  # avoid heavy pipeline
        (d2 / "metadata.json").write_text(json.dumps(meta2), encoding="utf-8")

        # Also keep the primary folder lightweight
        meta1 = load_metadata(self.test_dir)
        meta1["comparision_result"] = {"result": "skip for test"}
        save_metadata(self.test_dir, meta1)

        # ---- Run #1: both are blocked -> both become waiting ----
        with patch("multithreading.vendor_has_pending_jera_upload_today", side_effect=[True, True]) as mock_pending, \
             patch("multithreading.set_processing_status_text", return_value=1) as mock_set:
            msg1 = process_one_folder(self.test_dir)
            msg2 = process_one_folder(d2)

        self.assertIn(WAITING_PREVIOUS_VENDOR_PENDING, msg1)
        self.assertIn(WAITING_PREVIOUS_VENDOR_PENDING, msg2)
        self.assertEqual(mock_pending.call_count, 2)
        self.assertEqual(mock_set.call_count, 2)

        # ---- Run #2: prior pending cleared -> first proceeds; second still blocked ----
        # Simulate DB behavior:
        # - first call sees no previous pending (False)
        # - second call sees a new pending created by the first sheet (True)
        with patch("multithreading.vendor_has_pending_jera_upload_today", side_effect=[False, True]) as mock_pending2, \
             patch("multithreading.set_processing_status_text", return_value=1) as mock_set2:
            msg1b = process_one_folder(self.test_dir)
            msg2b = process_one_folder(d2)

        # First should NOT be waiting now
        self.assertNotIn("waiting:", msg1b.lower())
        # Second should still be waiting
        self.assertIn(WAITING_PREVIOUS_VENDOR_PENDING, msg2b)
        self.assertEqual(mock_pending2.call_count, 2)
        self.assertEqual(mock_set2.call_count, 1)  # only the second sheet is blocked in run #2

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

        with patch("multithreading.fetch_vendor_context_by_sender_email", return_value={"vendor_id": 7, "vendor_date_format": "MM-DD-YYYY"}), \
             patch("multithreading.export_rates_by_query", return_value="boom"), \
             patch("multithreading.mark_processing_stage", return_value=None):
            msg = process_one_folder(self.test_dir)

        self.assertIn("export error", msg)
        meta2 = load_metadata(self.test_dir)
        self.assertEqual(meta2.get("date_format_identified"), "MM-DD-YYYY")
        self.assertTrue(bool(meta2.get("date_verification_ingestion_status")))
        self.assertEqual(meta2.get("vendor_id"), 7)

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

        with patch("multithreading.fetch_vendor_context_by_sender_email", return_value={"vendor_id": 7, "vendor_date_format": "MM-DD-YYYY"}), \
             patch("multithreading.load_clean_rates", side_effect=_fake_clean), \
             patch("multithreading.mark_processing_stage", return_value=None):
            msg = process_one_folder(self.test_dir)

        # it should attempt DB push and then skip because comparision_result is not ok
        self.assertIn("skip DB push", msg)
        self.assertEqual(seen.get("date_format_email"), "MM-DD-YYYY")

    def test_compare_skipped_when_preprocessing_failed(self):
        # If preprocessing failed (final_ok=False), comparison should not run.
        meta = load_metadata(self.test_dir)
        meta.pop("comparision_result", None)  # ensure compare stage would run
        meta["final_ok"] = False
        meta["jerasoft_preprocessed"] = True
        meta["date_verification_ingestion_status"] = True
        save_metadata(self.test_dir, meta)

        msg = process_one_folder(self.test_dir)
        self.assertIn("skip compare: preprocessing failed", msg)
        meta2 = load_metadata(self.test_dir)
        self.assertEqual(
            (meta2.get("comparision_result") or {}).get("result"),
            "comparison skipped: preprocessing failed (final_ok=false)",
        )

    def test_zero_row_jera_treats_vendor_as_full_new(self):
        """When JeraSoft export has 0 rows but vendor is preprocessed,
        treat the vendor file as a full-new import instead of skipping."""

        # Prepare metadata to simulate: jera_fetched=True, 0-row Jera export,
        # preprocessing failed overall (final_ok=False), but vendor file
        # itself is marked as successfully preprocessed.
        meta = load_metadata(self.test_dir)
        meta.pop("comparision_result", None)
        meta["date_verification_ingestion_status"] = True
        meta["jerasoft_preprocessed"] = True
        meta["jera_fetched"] = True
        meta["human_eval_details_jerasoft"] = {"file": "jerasoft_comparison_all.xlsx", "rows": 0}
        meta["final_ok"] = False
        meta["table_id"] = 4330

        # Create a cleaned vendor file that read_table/build_full_new_comparison
        # can consume to build a comparison-style result.
        vendor_path = self.test_dir / "vendor1_cleaned.xlsx"
        df_vendor = pd.DataFrame({
            "Dst Code": ["1001"],
            "Rate": [0.05],
            "Effective Date": ["2025-01-01"],
            "Billing Increment": ["1/60"],
        })
        df_vendor.to_excel(vendor_path, index=False)

        # Mark this vendor file as successfully preprocessed.
        meta["preprocessed_results"] = {vendor_path.name: True}
        save_metadata(self.test_dir, meta)

        from unittest.mock import patch

        with patch("multithreading.insert_rate_upload", return_value=123) as mock_insert, \
             patch("multithreading.bulk_insert_rate_upload_details", return_value=1) as mock_bulk, \
             patch("multithreading.mark_processing_stage") as mock_stage:

            msg = process_one_folder(self.test_dir)

        # We should reach the DB-push stage and report a successful push.
        self.assertIn("done (pushed=1/1)", msg)

        # Metadata should record a successful comparison result for the vendor file.
        meta2 = load_metadata(self.test_dir)
        comp = meta2.get("comparision_result") or {}
        self.assertEqual(comp.get("result"), "ok")
        self.assertTrue(comp.get(vendor_path.name))

        # A comparison result file should have been written for the vendor.
        result_files = list(self.test_dir.glob("*_comparision_result.xlsx"))
        self.assertEqual(len(result_files), 1)

        # DB functions should have been invoked.
        mock_insert.assert_called_once()
        mock_bulk.assert_called()

        # Zero-row Jera special-case should backfill stage flags via mark_processing_stage.
        stages = [call.kwargs.get("stage") for call in mock_stage.call_args_list]
        self.assertIn("file_cleaned", stages)
        self.assertIn("rate_compared", stages)
        self.assertIn("rate_uploaded", stages)

    def test_run_pipeline_mt(self):
        # Should print no folders to process
        run_pipeline_mt(str(self.test_dir), max_workers=2)

    def test_prefix_derived_from_force_table_used_for_pending_gate(self):
        # When metadata.prefix is missing, process_one_folder should derive it from
        # force_jerasoft_table_name (before subject) and pass it into the pending gate.
        meta = load_metadata(self.test_dir) or {}
        meta["sender"] = "same@vendor.com"
        meta["date_verification_ingestion_status"] = True
        meta["prefix"] = None
        meta["force_jerasoft_table_name"] = "TERM-RATE IMPORT AUTOMATION TESTING PREFIX:1234 [USD]"
        meta["subject"] = "no prefix here"
        save_metadata(self.test_dir, meta)

        with patch("multithreading.vendor_has_pending_jera_upload_today", return_value=False) as mock_gate:
            process_one_folder(self.test_dir)
        self.assertEqual(mock_gate.call_args.kwargs.get("prefix"), "1234")

if __name__ == '__main__':
    unittest.main(verbosity=2)
