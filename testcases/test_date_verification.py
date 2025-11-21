import unittest
from unittest.mock import patch, MagicMock
import date_verification
import pandas as pd
from datetime import datetime, date, timezone
from pathlib import Path
import json

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


if __name__ == '__main__':
    unittest.main()
