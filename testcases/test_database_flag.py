import unittest
from unittest.mock import patch, MagicMock
import database_flag
from pathlib import Path
import json
from datetime import datetime, date

class TestDatabaseFlag(unittest.TestCase):
    @patch('database_flag.mark_ingest_processed')
    def test_finalize_processed_flags_all_true(self, mock_mark_ingest_processed):
        # Simulate two files, both with valid metadata and all results_pushed True
        paths_map = {'/tmp/f1.csv': 'fmt1', '/tmp/f2.csv': 'fmt2'}
        meta = {'results_pushed': {'a': True, 'b': True}}
        # Patch Path methods for each file
        def fake_exists(self):
            return str(self).endswith('metadata.json')
        def fake_open(self, *args, **kwargs):
            return unittest.mock.mock_open(read_data=json.dumps(meta))()
        def fake_parent(self):
            return type(self)(str(self).rsplit('/', 1)[0])
        def fake_resolve(self):
            return self
        with patch.object(Path, 'exists', fake_exists), \
             patch.object(Path, 'open', fake_open), \
             patch.object(Path, 'parent', property(fake_parent)), \
             patch.object(Path, 'resolve', fake_resolve):
            mock_mark_ingest_processed.return_value = 2
            result = database_flag.finalize_processed_flags(paths_map)
            self.assertEqual(result, (2, 2, 2))

    @patch('database_flag.mark_ingest_processed')
    def test_finalize_processed_flags_not_all_true(self, mock_mark_ingest_processed):
        # Not all results_pushed True, not eligible
        meta = {'results_pushed': {'a': True, 'b': False}}
        with patch('pathlib.Path.exists', return_value=True), \
             patch('builtins.open', unittest.mock.mock_open(read_data=json.dumps(meta))):
            paths_map = {'/tmp/f1.csv': 'fmt1'}
            mock_mark_ingest_processed.return_value = 0
            result = database_flag.finalize_processed_flags(paths_map)
            self.assertEqual(result[1], 0)
            self.assertEqual(result[2], 0)

    def test_folder_is_today_or_newer(self):
        today = date.today().isoformat()
        meta = {'receivedDateTime_raw': today + 'T00:00:00Z'}
        self.assertTrue(database_flag._folder_is_today_or_newer(meta))
        meta = {'date_utc': today}
        self.assertTrue(database_flag._folder_is_today_or_newer(meta))
        meta = {'receivedDateTime_raw': '2000-01-01T00:00:00Z'}
        self.assertFalse(database_flag._folder_is_today_or_newer(meta))
        meta = {'date_utc': '2000-01-01'}
        self.assertFalse(database_flag._folder_is_today_or_newer(meta))
        meta = {}
        self.assertTrue(database_flag._folder_is_today_or_newer(meta))

    @patch('database_flag.upsert_processing_status')
    @patch('database_flag._folder_is_today_or_newer', return_value=True)
    def test_seed_processing_status_rows(self, mock_folder_is_today_or_newer, mock_upsert):
        # Simulate one directory with valid metadata.json
        meta = {'internet_message_id': 'id1', 'directory': 'dir1', 'sender': 'a@b.com', 'subject': 'subj', 'receivedDateTime_raw': '2025-11-20T00:00:00Z'}
        root = Path('attachments')
        dir1 = root / 'dir1'
        meta_path = dir1 / 'metadata.json'
        def fake_exists(self):
            return str(self) in [str(root), str(dir1), str(meta_path)]
        def fake_is_dir(self):
            return str(self) in [str(root), str(dir1)]
        def fake_open(self, *args, **kwargs):
            return unittest.mock.mock_open(read_data=json.dumps(meta))()
        def fake_iterdir(self):
            if str(self) == str(root):
                return [dir1]
            return []
        def fake_resolve(self):
            return self
        def fake_expanduser(self):
            return self
        with patch.object(Path, 'exists', fake_exists), \
             patch.object(Path, 'is_dir', fake_is_dir), \
             patch.object(Path, 'open', fake_open), \
             patch.object(Path, 'iterdir', fake_iterdir), \
             patch.object(Path, 'resolve', fake_resolve), \
             patch.object(Path, 'expanduser', fake_expanduser):
            mock_upsert.return_value = 1
            result = database_flag.seed_processing_status_rows('attachments')
            self.assertEqual(result, (1, 1, 0))

    @patch('database_flag.get_processing_status')
    @patch('database_flag.get_conn')
    def test_reset_processing_flags_success(self, mock_get_conn, mock_get_status):
        # Simulate successful reset
        mock_get_status.return_value = {'directory_name': 'dir1'}
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 1
        with patch('pathlib.Path.exists', return_value=True), \
             patch('builtins.open', unittest.mock.mock_open(read_data=json.dumps({'jera_fetched': True}))):
            result = database_flag.reset_processing_flags('dir1', 'attachments')
            self.assertTrue(result)

    @patch('database_flag.get_processing_status')
    @patch('database_flag.get_conn')
    def test_reset_processing_flags_no_status(self, mock_get_conn, mock_get_status):
        # No processing status found
        mock_get_status.return_value = None
        result = database_flag.reset_processing_flags('dir1', 'attachments')
        self.assertFalse(result)

    @patch('database_flag.get_processing_status')
    @patch('database_flag.get_conn')
    def test_reset_processing_flags_no_rows_updated(self, mock_get_conn, mock_get_status):
        # No rows updated
        mock_get_status.return_value = {'directory_name': 'dir1'}
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 0
        with patch('pathlib.Path.exists', return_value=True), \
             patch('builtins.open', unittest.mock.mock_open(read_data=json.dumps({'jera_fetched': False}))):
            result = database_flag.reset_processing_flags('dir1', 'attachments')
            self.assertFalse(result)

    @patch('database_flag.get_processing_status')
    @patch('database_flag.get_conn')
    def test_reset_processing_flags_exception(self, mock_get_conn, mock_get_status):
        # Exception during reset
        mock_get_status.side_effect = Exception('fail')
        result = database_flag.reset_processing_flags('dir1', 'attachments')
        self.assertFalse(result)

if __name__ == '__main__':
    unittest.main()
