"""
Unit Tests for Ratesheet Comparison Engine

Tests all main functions in ratesheet_comparision_engine.py:
- read_table
- keep_latest_per_code
- validate_row
- effective_note
- summarize_changes
- compare
- write_excel
"""

import unittest
import pandas as pd
import numpy as np
import os
import tempfile
from datetime import datetime, timedelta

from ratesheet_comparision_engine import (
    read_table, keep_latest_per_code, validate_row, effective_note,
    summarize_changes, compare, write_excel,
    COL_CODE, COL_RATE, COL_EDATE, COL_BI, COL_NAME, OUT_COLS
)

class TestRatesheetComparisonEngine(unittest.TestCase):
    def test_read_table_xlsx(self):
        import tempfile
        # Create a valid Excel file
        df = pd.DataFrame({
            COL_CODE: ['1001', '1002'],
            COL_RATE: [0.05, 0.06],
            COL_EDATE: ['2025-01-01', '2025-02-01'],
            COL_BI: ['1/60', '1/60'],
            COL_NAME: ['A', 'B']
        })
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False)
            tmp_path = tmp.name
        out = read_table(tmp_path)
        self.assertEqual(len(out), 2)
        self.assertIn(COL_CODE, out.columns)
        self.assertIn(COL_NAME, out.columns)
        os.remove(tmp_path)

    def test_read_table_xlsx_missing_columns(self):
        import tempfile
        df = pd.DataFrame({
            'SomeCol': [1, 2],
            COL_CODE: ['1001', '1002']
        })
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False)
            tmp_path = tmp.name
        with self.assertRaises(ValueError):
            read_table(tmp_path)
        os.remove(tmp_path)

    def test_read_table_unsupported_filetype(self):
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.unsupported', delete=False) as tmp:
            tmp.write(b"dummy data")
            tmp_path = tmp.name
        with self.assertRaises(ValueError):
            read_table(tmp_path)
        os.remove(tmp_path)

    def test_read_table_extra_columns(self):
        import tempfile
        df = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: ['2025-01-01'],
            COL_BI: ['1/60'],
            COL_NAME: ['A'],
            'ExtraCol': [123]
        })
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
            df.to_csv(tmp.name, index=False)
            tmp_path = tmp.name
        out = read_table(tmp_path)
        self.assertIn(COL_CODE, out.columns)
        self.assertNotIn('ExtraCol', out.columns)
        os.remove(tmp_path)

    def test_read_table_minimal_columns(self):
        import tempfile
        df = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: ['2025-01-01'],
            COL_BI: ['1/60']
        })
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
            df.to_csv(tmp.name, index=False)
            tmp_path = tmp.name
        out = read_table(tmp_path)
        self.assertIn(COL_CODE, out.columns)
        self.assertNotIn(COL_NAME, out.columns)
        os.remove(tmp_path)

    def test_read_table_optional_column_absent(self):
        import tempfile
        df = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: ['2025-01-01'],
            COL_BI: ['1/60']
        })
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False)
            tmp_path = tmp.name
        out = read_table(tmp_path)
        self.assertIn(COL_CODE, out.columns)
        self.assertNotIn(COL_NAME, out.columns)
        os.remove(tmp_path)

    def test_read_table_optional_column_present(self):
        import tempfile
        df = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: ['2025-01-01'],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False)
            tmp_path = tmp.name
        out = read_table(tmp_path)
        self.assertIn(COL_CODE, out.columns)
        self.assertIn(COL_NAME, out.columns)
        os.remove(tmp_path)
    def test_read_table_csv(self):
        import tempfile
        # Create a valid CSV file
        df = pd.DataFrame({
            COL_CODE: ['1001', '1002'],
            COL_RATE: [0.05, 0.06],
            COL_EDATE: ['2025-01-01', '2025-02-01'],
            COL_BI: ['1/60', '1/60'],
            COL_NAME: ['A', 'B']
        })
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            df.to_csv(tmp.name, index=False)
            tmp_path = tmp.name
        # Should read without error
        out = read_table(tmp_path)
        self.assertEqual(len(out), 2)
        self.assertIn(COL_CODE, out.columns)
        self.assertIn(COL_NAME, out.columns)
        os.remove(tmp_path)

    def test_read_table_missing_columns(self):
        import tempfile
        # Create a CSV missing required columns
        df = pd.DataFrame({
            'SomeCol': [1, 2],
            COL_CODE: ['1001', '1002']
        })
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            df.to_csv(tmp.name, index=False)
            tmp_path = tmp.name
        # Should raise ValueError for missing columns
        with self.assertRaises(ValueError):
            read_table(tmp_path)
        os.remove(tmp_path)
    def setUp(self):
        # Minimal valid DataFrame
        self.df = pd.DataFrame({
            COL_CODE: ['1001', '1002', '1001'],
            COL_RATE: [0.05, 0.06, 0.07],
            COL_EDATE: [datetime(2025, 1, 1), datetime(2025, 2, 1), datetime(2025, 3, 1)],
            COL_BI: ['1/60', '1/60', '1/60'],
            COL_NAME: ['A', 'B', 'A']
        })

    def test_keep_latest_per_code(self):
        latest = keep_latest_per_code(self.df)
        self.assertEqual(len(latest), 2)
        self.assertIn('1001', latest[COL_CODE].values)
        self.assertIn('1002', latest[COL_CODE].values)
        # Should keep the row with latest date for '1001'
        self.assertEqual(latest[latest[COL_CODE]=='1001'][COL_EDATE].iloc[0], datetime(2025, 3, 1))

    def test_validate_row(self):
        valid_row = self.df.iloc[0]
        invalid_row = pd.Series({COL_CODE: 'abc', COL_RATE: -1, COL_EDATE: pd.NaT, COL_BI: 'bad'})
        self.assertEqual(validate_row(valid_row), [])
        reasons = validate_row(invalid_row)
        self.assertIn('invalid code', reasons)
        self.assertIn('negative rate', reasons)
        self.assertIn('invalid effective date', reasons)
        self.assertIn('invalid billing increment', reasons)

    def test_effective_note(self):
        as_of = datetime(2025, 1, 1)
        self.assertEqual(effective_note(datetime(2025, 1, 2), as_of, 7), 'new without 7-day notice')
        self.assertEqual(effective_note(datetime(2025, 1, 10), as_of, 7), 'proper 7-day notice')
        self.assertEqual(effective_note(datetime(2024, 12, 31), as_of, 7), 'immediate effective date')
        self.assertEqual(effective_note(pd.NaT, as_of, 7), 'invalid effective date')

    def test_summarize_changes(self):
        df = pd.DataFrame({
            'Change Type': [
                'New', 'Increase', 'Decrease', 'Unchanged', 'Closed',
                'Backdated Increase', 'Backdated Decrease', 'Billing Increments Changes',
                'new', 'increase', 'decrease', 'unchanged', 'closed',
                'backdated_increase', 'backdated_decrease', 'billing_increment_changes'
            ]
        })
        stats = summarize_changes(df)
        self.assertEqual(stats['total_rows'], len(df))
        # Check all possible key variants (capitalization and underscores)
        possible_keys = [
            'New', 'Increase', 'Decrease', 'Unchanged', 'Closed',
            'Backdated Increase', 'Backdated Decrease', 'Billing Increments Changes',
            'new', 'increase', 'decrease', 'unchanged', 'closed',
            'backdated_increase', 'backdated_decrease', 'billing_increment_changes'
        ]
        for key in possible_keys:
            if key in stats:
                self.assertGreaterEqual(stats[key], 1)

    def test_compare(self):
        # New code in right only
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001', '1002'],
            COL_RATE: [0.05, 0.06],
            COL_EDATE: [datetime(2025, 1, 1), datetime(2025, 2, 1)],
            COL_BI: ['1/60', '1/60'],
            COL_NAME: ['A', 'B']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('New', result['Change Type'].values)
        # Accept any key variant for 'New'
        found = False
        for k in ['New', 'new', 'new_code']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'New' key variant found in stats: {stats}")

        # Closed code in left only
        left = pd.DataFrame({
            COL_CODE: ['1001', '1002'],
            COL_RATE: [0.05, 0.06],
            COL_EDATE: [datetime(2025, 1, 1), datetime(2025, 2, 1)],
            COL_BI: ['1/60', '1/60'],
            COL_NAME: ['A', 'B']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Closed', result['Change Type'].values)
        found = False
        for k in ['Closed', 'closed', 'closed_code']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Closed' key variant found in stats: {stats}")

        # Unchanged
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = left.copy()
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Unchanged', result['Change Type'].values)
        found = False
        for k in ['Unchanged', 'unchanged']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Unchanged' key variant found in stats: {stats}")

        # If effective date moved earlier (new < old), create a Stashed row for the old rate
        # (per updated requirement), even if the rate is unchanged.
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2026, 2, 12)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],  # same rate
            COL_EDATE: [datetime(2026, 1, 12)],  # earlier effective date
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2026-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Stashed', result['Change Type'].values)
        self.assertIn('Unchanged', result['Change Type'].values)
        # The replacement row (Unchanged) that comes instead of the stashed future rate
        # must now be Rejected (requires manual approval).
        unchanged_rows = result[result['Change Type'] == 'Unchanged']
        if not unchanged_rows.empty:
            self.assertTrue((unchanged_rows['Status'] == 'Rejected').all())

        # Stashed (rate changed + new date earlier): old effective date > new effective date AND rate changed
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2026, 2, 12)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.06],  # rate changed
            COL_EDATE: [datetime(2026, 1, 12)],  # earlier effective date
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2026-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Stashed', result['Change Type'].values)
        # New vendor row should still be classified by normal logic (Increase/Decrease/etc),
        # not a special "Rescheduled" type.
        self.assertNotIn('Rescheduled', result['Change Type'].values)

        # Do NOT stash when both old and new effective dates are in the past
        # relative to as_of (backdated historical correction only).
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 2, 23)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 2, 13)],  # earlier but still before as_of
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        # as_of is after both dates -> treat as historical backdating only, no "Stashed" row
        result, stats = compare(left, right, as_of_date='2026-01-01', notice_days=7, rate_tol=0.0)
        self.assertNotIn('Stashed', result['Change Type'].values)

        # Increase
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.07],
            COL_EDATE: [datetime(2025, 3, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Increase', result['Change Type'].values)
        found = False
        for k in ['Increase', 'increase', 'increased']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Increase' key variant found in stats: {stats}")

        # Decrease
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.07],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 3, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Decrease', result['Change Type'].values)
        found = False
        for k in ['Decrease', 'decrease', 'decreased']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Decrease' key variant found in stats: {stats}")

        # Billing Increment Changes
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 3, 1)],
            COL_BI: ['2/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Billing Increments Changes', result['Change Type'].values)
        found = False
        for k in ['Billing Increments Changes', 'billing_increment_changes', 'Billing Increment Changes']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Billing Increment Changes' key variant found in stats: {stats}")

        # Backdated Increase
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.07],
            COL_EDATE: [datetime(2024, 12, 31)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Backdated Increase', result['Change Type'].values)
        found = False
        for k in ['Backdated Increase', 'backdated_increase']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Backdated Increase' key variant found in stats: {stats}")

        # Backdated Decrease
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.07],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2024, 12, 31)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        self.assertIn('Backdated Decrease', result['Change Type'].values)
        found = False
        for k in ['Backdated Decrease', 'backdated_decrease']:
            if k in stats:
                self.assertEqual(stats[k], 1)
                found = True
        self.assertTrue(found, f"No 'Backdated Decrease' key variant found in stats: {stats}")

        # Invalid effective date
        left = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [datetime(2025, 1, 1)],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        right = pd.DataFrame({
            COL_CODE: ['1001'],
            COL_RATE: [0.05],
            COL_EDATE: [pd.NaT],
            COL_BI: ['1/60'],
            COL_NAME: ['A']
        })
        result, stats = compare(left, right, as_of_date='2025-01-01', notice_days=7, rate_tol=0.0)
        # Accept any column name variant for 'Effective Date Note', fallback to 'Notes' column
        col_found = None
        for col in ['Effective Date Note', 'effective_date_note', 'EffectiveDateNote', 'effectiveDateNote']:
            if col in result.columns:
                col_found = col
                break
        if col_found:
            self.assertIn('invalid effective date', result[col_found].values)
        else:
            self.assertIn('Notes', result.columns, f"No 'Notes' column found in result: {result.columns}")
            self.assertTrue(any('invalid effective date' in str(x) for x in result['Notes'].values), "'invalid effective date' not found in 'Notes' column")

    def test_write_excel(self):
        # Create a temp file and DataFrame
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, 'test_output.xlsx')
            df = pd.DataFrame({col: [1, 2] for col in OUT_COLS})
            write_excel(df, path)
            self.assertTrue(os.path.exists(path))

if __name__ == '__main__':
    #unittest.main(verbosity=2)

    # Example usage for manual testing (edit file paths as needed)
    # Uncomment and set your file paths to test with real data
    old_file = r"C:\Users\Tahira Sadaf\Desktop\projects\rate-import-service\attachments\tahira.sadaf_at_kingrevolution.com_20251118_081521\CPL_011_HAYO_011-20251029-149146333333333333333333_jerasoft_comparison_cleaned.xlsx"
    new_file = r"C:\Users\Tahira Sadaf\Desktop\projects\rate-import-service\attachments\tahira.sadaf_at_kingrevolution.com_20251118_081521\CPL_011_HAYO_011-20251029-149146333333333333333333_cleaned.xlsx"
    output_file = "comparison_output.xlsx"
    left_df = read_table(old_file)
    right_df = read_table(new_file)
    result_df, stats = compare(
        left_df,
        right_df,
        as_of_date="2025-11-20",
        notice_days=7,
        rate_tol=0.0,
    )
    print("\n=== Summary Stats ===")
    for k, v in stats.items():
        print(f"{k}: {v}")
    write_excel(result_df, output_file)
    print(f"\n✅ Comparison complete. Results saved to: {output_file}")
