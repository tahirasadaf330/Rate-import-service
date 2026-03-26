import unittest
from unittest.mock import patch, MagicMock
import database
import psycopg2
from datetime import datetime
from database import _parse_iso_utc
from pathlib import Path

class TestDatabaseModule(unittest.TestCase):
    def test_short_failed_status_cases(self):
        self.assertEqual(database._short_failed_status(None), "failed")
        self.assertEqual(database._short_failed_status("   "), "failed")
        self.assertEqual(
            database._short_failed_status("some reason"),
            "failed: some reason",
        )

        long_msg = "x" * 400
        status_text = database._short_failed_status(long_msg)
        self.assertTrue(status_text.startswith("failed: "))
        self.assertTrue(status_text.endswith("..."))
        self.assertLessEqual(len(status_text), 255)

    @patch('database.get_conn')
    def test_get_pending_jera_uploads(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [(1, 'subj', 'email', 2, datetime.now(), 10, datetime.now(), datetime.now())]
        mock_cursor.description = [(col,) for col in ['id','subject','sender_email','jera_table_id','processed_at','total_rows','created_at','updated_at']]
        result = database.get_pending_jera_uploads()
        self.assertIsInstance(result, list)
        self.assertEqual(result[0]['id'], 1)

    @patch('database.get_conn')
    def test_vendor_has_pending_jera_upload_today_cases(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor

        # Found a pending row
        mock_cursor.fetchall.return_value = [("PRM",)]
        self.assertTrue(database.vendor_has_pending_jera_upload_today(sender_email="a@b.com", trunk="PRM", exclude_internet_message_id="mid-1"))

        # Not found
        mock_cursor.fetchall.return_value = []
        self.assertFalse(database.vendor_has_pending_jera_upload_today(sender_email="a@b.com", trunk="PRM"))

        # Null DB trunk should not fallback to subject parsing.
        mock_cursor.fetchall.return_value = [(None,)]
        self.assertFalse(database.vendor_has_pending_jera_upload_today(sender_email="a@b.com", trunk="PRM"))

        # Empty sender should be False (and not query DB)
        self.assertFalse(database.vendor_has_pending_jera_upload_today(sender_email=""))

    @patch('database.get_conn')
    def test_fetch_vendor_context_by_sender_email_includes_is_mapped(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = (7, 'YYYY-MM-DD', True, True)

        result = database.fetch_vendor_context_by_sender_email('a@b.com')

        self.assertEqual(result['vendor_id'], 7)
        self.assertEqual(result['vendor_date_format'], 'YYYY-MM-DD')
        self.assertTrue(result['is_partial'])
        self.assertTrue(result['is_mapped'])
        executed_sql = mock_cursor.execute.call_args_list[0].args[0]
        self.assertIn('is_header_mapping_set', executed_sql)

    @patch('database.get_conn')
    def test_fetch_active_vendor_header_mapping_returns_latest_active(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = (22, 3)
        mock_cursor.fetchall.return_value = [
            ('dst_code', 'Rates', 8, 'IBIS codes'),
            ('rate', 'Rates', 8, 'TDE'),
            ('effective_date', 'Rates', 8, 'Start Date'),
            ('billing_increment', 'Rates', 8, 'Price Status'),
        ]

        result = database.fetch_active_vendor_header_mapping(7)

        self.assertEqual(result['vendor_header_mapping_id'], 22)
        self.assertEqual(result['vendor_id'], 7)
        self.assertEqual(result['version'], 3)
        self.assertEqual(result['source_sheet'], 'Rates')
        self.assertEqual(result['header_row_index'], 8)
        self.assertEqual(len(result['field_mappings']), 4)
        self.assertEqual(result['field_mappings'][0]['source_header'], 'IBIS codes')

    @patch('database.get_conn')
    def test_set_processing_status_text_cases(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 1

        affected = database.set_processing_status_text(directory_name="dir", status_text="waiting: something")
        self.assertEqual(affected, 1)
        mock_conn.commit.assert_called()

        # Missing key should raise
        with self.assertRaises(ValueError):
            database.set_processing_status_text(status_text="x")

        # Empty status text should raise
        with self.assertRaises(ValueError):
            database.set_processing_status_text(directory_name="dir", status_text="")

    @patch('database.get_conn')
    def test_update_jera_upload_status(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        database.update_jera_upload_status(1, 'success')
        database.update_jera_upload_status(2, 'failed', {'error': 'fail'})
        database.update_jera_upload_status(3, 'pending')
        mock_conn.commit.assert_called()

    @patch('database.get_conn')
    def test_get_jera_upload_history(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [(1, 'subj', 'email', 2, True, 'success', None, datetime.now(), 10, datetime.now(), datetime.now())]
        mock_cursor.description = [(col,) for col in ['id','subject','sender_email','jera_table_id','is_rate_approved_by_admin','jera_upload_status','jera_upload_result','jera_uploaded_at','total_rows','processed_at','created_at']]
        result = database.get_jera_upload_history()
        self.assertIsInstance(result, list)
        self.assertEqual(result[0]['id'], 1)

    @patch('database.get_conn')
    @patch('database.set_jera_upload_flag')
    def test_mark_rate_upload_for_jera(self, mock_set_flag, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [42]
        mock_set_flag.return_value = None
        result = database.mark_rate_upload_for_jera('subject', 'email', 99)
        self.assertEqual(result, 42)
        mock_cursor.fetchone.return_value = None
        result_none = database.mark_rate_upload_for_jera('subject', 'email', 99)
        self.assertIsNone(result_none)

    @patch('database.get_conn')
    def test_mark_comparison_file_for_bulk_upload(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [55]
        mock_cursor.execute.return_value = None
        result = database.mark_comparison_file_for_bulk_upload('file.csv', 'subject', 'email', 88)
        self.assertEqual(result, 55)
        mock_cursor.fetchone.return_value = None
        result_none = database.mark_comparison_file_for_bulk_upload('file.csv', 'subject', 'email', 88)
        self.assertIsNone(result_none)

    @patch('database.get_conn')
    @patch('database.Path')
    def test_auto_update_status_on_import_flag_change(self, mock_path, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 2
        mock_cursor.fetchall.return_value = [(1, 'subj', 'email')]
        mock_folder = MagicMock()
        mock_folder.is_dir.return_value = True
        mock_folder.name = 'email_folder'
        mock_metadata = MagicMock()
        mock_metadata.exists.return_value = True
        mock_path.return_value.exists.return_value = True
        mock_path.return_value.iterdir.return_value = [mock_folder]
        mock_folder.__truediv__.return_value = mock_metadata
        with patch('builtins.open', unittest.mock.mock_open(read_data='{"table_id": 123}')):
            result = database.auto_update_status_on_import_flag_change()
            self.assertIsInstance(result, int)

    @patch('database.get_conn')
    def test_parse_iso_utc_cases(self, mock_get_conn):
        from database import _parse_iso_utc
        # Valid ISO UTC
        self.assertIsInstance(_parse_iso_utc('2025-09-28T12:34:56Z'), datetime)
        # Invalid string returns None
        self.assertIsNone(_parse_iso_utc('not-a-date'))
        # None input returns None
        self.assertIsNone(_parse_iso_utc(None))
        # Non-string input returns None
        self.assertIsNone(_parse_iso_utc(123))
        # Accepts ISO without Z, returns datetime
        self.assertIsInstance(_parse_iso_utc('2025-09-28T12:34:56+00:00'), datetime)

    @patch('database.get_conn')
    def test__parse_iso_utc_all_cases(self, mock_get_conn):
        from database import _parse_iso_utc
        # Valid ISO UTC
        self.assertIsInstance(_parse_iso_utc('2025-09-28T12:34:56Z'), datetime)
        # Accepts ISO without Z, returns datetime
        self.assertIsInstance(_parse_iso_utc('2025-09-28T12:34:56+00:00'), datetime)
        # Accepts ISO without timezone, returns datetime
        self.assertIsInstance(_parse_iso_utc('2025-09-28 12:34:56'), datetime)
        # Also check exact value for ISO without timezone
        self.assertEqual(_parse_iso_utc('2025-09-28 12:34:56'), datetime(2025, 9, 28, 12, 34, 56))
        # Invalid string returns None
        self.assertIsNone(_parse_iso_utc('not-a-date'))
        # None input returns None
        self.assertIsNone(_parse_iso_utc(None))
        # Non-string input returns None
        self.assertIsNone(_parse_iso_utc(123))
        # Invalid format
        self.assertIsNone(_parse_iso_utc('2025/09/28T12:34:56Z'))
        # Partial date returns datetime
        self.assertEqual(_parse_iso_utc('2025-09-28'), datetime(2025, 9, 28))

    @patch('valid_emails.get_verified_senders', return_value=['rates@saifglobal.net'])
    @patch('database.get_conn')
    def test_insert_rejected_email_cases(self, mock_get_conn, mock_get_verified):
        from database import insert_rejected_email
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate no duplicate found
        mock_cursor.fetchone.return_value = None
        # Simulate insert returning id
        mock_cursor.fetchone.return_value = [1]
        # Use an authorized sender email from VERIFIED_SENDERS
        result = insert_rejected_email('rates@saifglobal.net', 'subj', 'cat', 'notes', None, None)
        self.assertEqual(result, 1)
        # With internet_message_id
        result = insert_rejected_email('rates@saifglobal.net', 'subj', 'cat', 'notes', None, None, 'msg-id')
        self.assertEqual(result, 1)
        # Test missing category (should return 1, not None)
        result = insert_rejected_email('rates@saifglobal.net', 'subj', None, 'notes', None, None)
        self.assertEqual(result, 1)

    @patch('database.execute_values')
    @patch('database.get_conn')
    def test_insert_rejected_emails_cases(self, mock_execute_values, mock_get_conn):
        from database import insert_rejected_emails
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        class Conn:
            encoding = 'utf-8'
        mock_cursor.connection = Conn()
        rows = [
            {"sender_email": "a@b.com", "subject": "subj", "category": "cat", "notes": "n", "received_at": None, "processed_at": None, "internet_message_id": "mid-1"},
            {"sender_email": "b@c.com", "subject": "subj2", "category": "cat2", "notes": "n2", "received_at": None, "processed_at": None, "internet_message_id": "mid-2"}
        ]
        with patch('psycopg2.extras.execute_values', mock_execute_values):
            insert_rejected_emails(rows)
            self.assertEqual(insert_rejected_emails([]), [])
            # Test missing category (should raise ValueError)
            with self.assertRaises(ValueError):
                insert_rejected_emails([{"sender_email": "a@b.com", "subject": "subj", "notes": "n", "received_at": None, "processed_at": None}])
    def test_atomic_write_json_cases(self):
        from database import _atomic_write_json
        import tempfile, os
        # Valid path and data
        with tempfile.NamedTemporaryFile(delete=False) as tf:
            path = Path(tf.name)
        _atomic_write_json(path, {'a': 1})
        self.assertTrue(path.exists())
        # Data not serializable
        with self.assertRaises(TypeError):
            _atomic_write_json(path, {'a': set([1,2])})
        # Permission error (simulate by passing directory)
        with self.assertRaises(Exception):
            _atomic_write_json(Path('/'), {'a': 1})

    @patch('database.get_conn')
    def test_push_failed_emails_json_to_db_cases(self, mock_get_conn):
        from database import push_failed_emails_json_to_db
        import tempfile, json, os
        # Valid JSON file with new entries
        with tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='utf-8') as tf:
            json.dump({"buckets": {"cat": [{"sender": "a@b.com", "subject": "subj", "already_pushed": False,
                                               "receivedDateTime": "2025-09-28T12:34:56Z",
                                               "logged_at_utc": "2025-09-28T12:34:56Z",
                                               "internetMessageId": "msg-id-1",
                                               "details": {"info": "x"}}]}}, tf)
            tf.close()
            with patch('database.insert_rejected_email') as mock_insert:
                mock_insert.return_value = 1
                push_failed_emails_json_to_db(tf.name)
                mock_insert.assert_called_once()
                self.assertEqual(mock_insert.call_args.kwargs.get("internet_message_id"), "msg-id-1")

            # JSON entry should only be marked after a successful insert
            with open(tf.name, "r", encoding="utf-8") as f:
                updated = json.load(f)
            self.assertTrue(updated["buckets"]["cat"][0].get("already_pushed"))
            os.remove(tf.name)

        # Insert failure should NOT mark already_pushed (so it can retry)
        with tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='utf-8') as tf:
            json.dump({"buckets": {"cat": [{"sender": "a@b.com", "subject": "subj", "already_pushed": False,
                                               "receivedDateTime": "2025-09-28T12:34:56Z",
                                               "logged_at_utc": "2025-09-28T12:34:56Z",
                                               "internetMessageId": "msg-id-2",
                                               "details": {"info": "x"}}]}}, tf)
            tf.close()
            with patch('database.insert_rejected_email', side_effect=Exception("db down")):
                res = push_failed_emails_json_to_db(tf.name)
                self.assertGreaterEqual(res[2], 1)

            with open(tf.name, "r", encoding="utf-8") as f:
                updated = json.load(f)
            self.assertFalse(bool(updated["buckets"]["cat"][0].get("already_pushed")))
            os.remove(tf.name)
        # JSON file missing
        self.assertEqual(push_failed_emails_json_to_db('notfound.json'), (0,0,0))
        # Malformed JSON
        with tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='utf-8') as tf:
            tf.write('{bad json}')
            tf.close()
            self.assertEqual(push_failed_emails_json_to_db(tf.name)[2], 1)
            os.remove(tf.name)
        # All entries already pushed
        with tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='utf-8') as tf:
            json.dump({"buckets": {"cat": [{"already_pushed": True}]}}, tf)
            tf.close()
            self.assertEqual(push_failed_emails_json_to_db(tf.name)[0], 0)
            os.remove(tf.name)

    @patch('database.get_conn')
    def test_insert_rate_upload_cases(self, mock_get_conn):
        from database import insert_rate_upload
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1]
        # All valid
        insert_rate_upload(sender_email='a@b.com', subject='subj', totals={'total_rows': 1}, jera_table_id=1, comparison_file_path='file.csv')
        # Some None
        insert_rate_upload(sender_email=None, subject=None, totals=None, jera_table_id=None, comparison_file_path=None)
        # Invalid totals
        insert_rate_upload(sender_email='a@b.com', subject='subj', totals={'bad': 1}, jera_table_id=1, comparison_file_path='file.csv')
        # With internet_message_id
        insert_rate_upload(sender_email='a@b.com', subject='subj', totals={'total_rows': 1}, jera_table_id=1, comparison_file_path='file.csv', internet_message_id='msg-id')

    @patch('database.get_conn')
    @patch('database.execute_values', return_value=[(1,)])
    def test_bulk_insert_rate_upload_details_cases(self, mock_execute_values, mock_get_conn):
        from database import bulk_insert_rate_upload_details
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_connection = MagicMock()
        mock_connection.encoding = 'utf-8'
        mock_cursor.connection = mock_connection
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        details = [{
            "dst_code": "code",
            "rate_existing": 1.0,
            "rate_new": 2.0,
            "effective_date": None,
            "change_type": "inc",
            "status": "active",
            "notes": "n",
            "old_billing_increment": None,
            "new_billing_increment": None,
            "code_name": "name"
        }]
        bulk_insert_rate_upload_details(1, details)
        bulk_insert_rate_upload_details(1, [], batch_size=1)

    @patch('database.get_conn')
    def test_fetch_vendor_contact_emails_cases(self, mock_get_conn):
        from database import fetch_vendor_contact_emails
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [('a@b.com',), (' b@c.com ',), ('',)]

        emails = fetch_vendor_contact_emails()
        self.assertIn('a@b.com', emails)
        self.assertIn('b@c.com', emails)

        # No emails
        mock_cursor.fetchall.return_value = []
        self.assertEqual(fetch_vendor_contact_emails(), [])

    @patch('database.get_conn')
    def test_insert_or_update_ingest_file_cases(self, mock_get_conn):
        from database import insert_or_update_ingest_file
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1]
        # All valid
        insert_or_update_ingest_file(email_address='a@b.com', subject='subj', received_at=datetime.now(), processed_at=datetime.now(), file_path='file.csv', internet_message_id='msg-id', preview_cache={'a':1}, error_message=None, status='pending', date_format='fmt', approved_at=None, is_format_auto_detected=True)
        # Some None
        insert_or_update_ingest_file(email_address=None, subject=None, received_at=None, processed_at=None, file_path='file.csv')
        # file_path missing
        with self.assertRaises(ValueError):
            insert_or_update_ingest_file(email_address='a@b.com', subject='subj', received_at=datetime.now(), processed_at=datetime.now(), file_path=None)

    @patch('database.get_conn')
    def test_bulk_upsert_ingest_files_cases(self, mock_get_conn):
        from database import bulk_upsert_ingest_files
        with patch('database.execute_values') as mock_execute_values:
            mock_conn = MagicMock()
            mock_cursor = MagicMock()
            mock_connection = MagicMock()
            mock_connection.encoding = 'utf-8'
            mock_cursor.connection = mock_connection
            mock_get_conn.return_value.__enter__.return_value = mock_conn
            mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
            rows = [
                {"email_address": "a@b.com", "subject": "subj", "received_at": None, "processed_at": None, "file_path": "f1", "internet_message_id": "mid-1", "preview_cache": None, "error_message": None},
                {"email_address": "b@c.com", "subject": "subj2", "received_at": None, "processed_at": None, "file_path": "f2", "internet_message_id": "mid-2", "preview_cache": None, "error_message": None}
            ]
            bulk_upsert_ingest_files(rows)
            self.assertEqual(bulk_upsert_ingest_files([]), 0)
            with self.assertRaises(ValueError):
                bulk_upsert_ingest_files([{"email_address": "a@b.com"}])

    @patch('database.get_conn')
    def test_fetch_approved_unprocessed_paths_map_cases(self, mock_get_conn):
        from database import fetch_approved_unprocessed_paths_map
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [('file.csv', 'fmt')]
        # limit None
        result = fetch_approved_unprocessed_paths_map()
        self.assertIn('file.csv', result)
        # limit integer
        result = fetch_approved_unprocessed_paths_map(1)
        self.assertIn('file.csv', result)
        # No matching rows
        mock_cursor.fetchall.return_value = []
        self.assertEqual(fetch_approved_unprocessed_paths_map(), {})

    @patch('database.get_conn')
    def test_mark_ingest_processed_cases(self, mock_get_conn):
        from database import mark_ingest_processed
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 2
        # Valid file_paths
        self.assertEqual(mark_ingest_processed(['file.csv', 'other.csv']), 2)
        # Empty file_paths
        self.assertEqual(mark_ingest_processed([]), 0)
        # processed True/False
        mark_ingest_processed(['file.csv'], False)

    @patch('database.get_conn')
    def test_upsert_processing_status_cases(self, mock_get_conn):
        from database import upsert_processing_status
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1]
        # All valid
        upsert_processing_status(internet_message_id='id', directory_name='dir')
        # Missing internet_message_id
        with self.assertRaises(ValueError):
            upsert_processing_status(internet_message_id=None, directory_name='dir')
        # Missing directory_name
        with self.assertRaises(ValueError):
            upsert_processing_status(internet_message_id='id', directory_name=None)

        executed_sql = mock_cursor.execute.call_args_list[0].args[0]
        executed_params = mock_cursor.execute.call_args_list[0].args[1]
        self.assertIn("COALESCE(%s, FALSE)", executed_sql)
        self.assertIn("WHEN %s IS NULL THEN processing_statuses.is_reprocessing_enabled", executed_sql)
        self.assertEqual(executed_params[-2:], (None, None))

    @patch('database.get_conn')
    def test_ensure_row_by_directory_cases(self, mock_get_conn):
        from database import ensure_row_by_directory
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate row exists
        mock_cursor.fetchone.side_effect = [[1], None]
        ensure_row_by_directory(directory_name='dir1')
        executed_sql = mock_cursor.execute.call_args_list[1].args[0]
        executed_params = mock_cursor.execute.call_args_list[1].args[1]
        self.assertIn("WHEN %s IS NULL THEN is_reprocessing_enabled", executed_sql)
        self.assertEqual(executed_params[-3:-1], (None, None))
        # Simulate row does not exist, requires internet_message_id
        mock_cursor.fetchone.side_effect = [None, [2]]
        ensure_row_by_directory(directory_name='dir2', internet_message_id='id2')
        insert_sql = mock_cursor.execute.call_args_list[3].args[0]
        self.assertIn("COALESCE(%s, FALSE)", insert_sql)
        # Test missing directory_name
        with self.assertRaises(ValueError):
            ensure_row_by_directory(directory_name='')
        # Test missing internet_message_id for insert
        mock_cursor.fetchone.side_effect = [None]
        with self.assertRaises(ValueError):
            ensure_row_by_directory(directory_name='dir3')

    @patch('database.upsert_rejected_email_from_processing_failure')
    @patch('database.get_conn')
    def test_mark_processing_stage_cases(self, mock_get_conn, mock_upsert):
        from database import mark_processing_stage, _short_failed_status
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 1
        # When a row is marked failed, processing_statuses is selected and
        # upsert_rejected_email_from_processing_failure is called with context.
        mock_cursor.fetchone.return_value = ('msg-id', 'a@b.com', 'subj', datetime.now())

        # Valid stage, final_status True (no rejected_emails sync)
        mark_processing_stage(directory_name='dir', stage='rate_uploaded', final_status=True)
        self.assertFalse(mock_upsert.called)

        # Non-final progress updates must preserve any failure status prefix.
        mark_processing_stage(directory_name='dir', stage='jera_fetched')

        # Valid stage, final_status False (with and without error message)
        mark_processing_stage(directory_name='dir', stage='rate_uploaded', final_status=False)
        mark_processing_stage(directory_name='dir', stage='rate_uploaded', final_status=False, error_message='some reason')
        self.assertTrue(mock_upsert.called)

        long_error = 'Ambiguous columns for Dst Code: ' + ('x' * 400)
        expected_status = _short_failed_status(long_error)
        mark_processing_stage(directory_name='dir', stage='rate_uploaded', final_status=False, error_message=long_error)

        update_calls = [call for call in mock_cursor.execute.call_args_list if 'UPDATE processing_statuses SET' in call.args[0]]
        self.assertTrue(update_calls)
        self.assertEqual(update_calls[-1].args[1][0], expected_status)
        self.assertLessEqual(len(update_calls[-1].args[1][0]), 255)
        success_update_sql = next(call.args[0] for call in update_calls if "status = CASE WHEN" in call.args[0])
        self.assertIn("status LIKE 'failed%%'", success_update_sql)

        self.assertEqual(mock_upsert.call_args.kwargs.get('status_text'), expected_status)

        # Missing directory_name and internet_message_id
        with self.assertRaises(ValueError):
            mark_processing_stage(stage='rate_uploaded', final_status=True)

    @patch('database.get_conn')
    def test_backfill_ingest_files_jera_table_from_metadata_cases(self, mock_get_conn):
        from database import backfill_ingest_files_jera_table_from_metadata
        import tempfile, json, os, shutil

        # Prepare a temporary folder with metadata.json
        temp_dir = tempfile.mkdtemp()
        try:
            file_path = os.path.join(temp_dir, 'file.csv')
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write('')
            meta_path = os.path.join(temp_dir, 'metadata.json')
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump({"best_table_name": "TEST_TABLE"}, f)

            mock_conn = MagicMock()
            mock_cursor = MagicMock()
            mock_get_conn.return_value.__enter__.return_value = mock_conn
            mock_conn.cursor.return_value.__enter__.return_value = mock_cursor

            # First connection (ALTER TABLE) ignores fetchall.
            # Second connection: SELECT ingest_files rows to backfill.
            mock_cursor.fetchall.return_value = [(1, file_path)]
            mock_cursor.rowcount = 1

            updated = backfill_ingest_files_jera_table_from_metadata()
            self.assertGreaterEqual(updated, 1)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    @patch('database.get_conn')
    def test_ensure_internet_message_id_links_cases(self, mock_get_conn):
        from database import ensure_internet_message_id_links
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor

        ensure_internet_message_id_links()

        # At least one ALTER statement should have been attempted
        self.assertTrue(mock_cursor.execute.called)

    @patch('database.get_conn')
    def test_get_processing_status_cases(self, mock_get_conn):
        from database import get_processing_status
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1]*15
        # Valid directory_name
        self.assertIsInstance(get_processing_status(directory_name='dir'), dict)
        # Valid internet_message_id
        self.assertIsInstance(get_processing_status(internet_message_id='id'), dict)
        # Missing both
        with self.assertRaises(ValueError):
            get_processing_status()
        # No matching row
        mock_cursor.fetchone.return_value = None
        self.assertIsNone(get_processing_status(directory_name='dir'))

    @patch('database.get_conn')
    def test_update_reprocessing_enabled_cases(self, mock_get_conn):
        from database import update_reprocessing_enabled
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 1
        # Valid directory_name
        update_reprocessing_enabled(directory_name='dir', is_reprocessing_enabled=True)
        # Valid internet_message_id
        update_reprocessing_enabled(internet_message_id='id', is_reprocessing_enabled=False)
        # Missing both
        with self.assertRaises(ValueError):
            update_reprocessing_enabled(is_reprocessing_enabled=True)

    @patch('database.get_conn')
    def test_get_reprocessing_enabled_directories_cases(self, mock_get_conn):
        from database import get_reprocessing_enabled_directories
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate DB rows
        mock_cursor.fetchall.return_value = [('dir1',), ('dir2',)]
        dirs = get_reprocessing_enabled_directories()
        self.assertIn('dir1', dirs)
        self.assertIn('dir2', dirs)
        executed_sql = mock_cursor.execute.call_args_list[0].args[0]
        self.assertNotIn("COALESCE(status, '') = 'processing'", executed_sql)
        # Test empty result
        mock_cursor.fetchall.return_value = []
        self.assertEqual(get_reprocessing_enabled_directories(), [])

    @patch('database.get_conn')
    def test_get_failed_directories_for_reprocessing_cases(self, mock_get_conn):
        from database import get_failed_directories_for_reprocessing
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [(1, 'id', 'dir', 'a@b.com', 'subj', datetime.now(), 'failed', True, datetime.now(), datetime.now())]
        # limit None
        result = get_failed_directories_for_reprocessing()
        self.assertIsInstance(result, list)
        executed_sql = mock_cursor.execute.call_args_list[0].args[0]
        self.assertNotIn("COALESCE(status, '') = 'processing'", executed_sql)
        # limit integer
        result = get_failed_directories_for_reprocessing(1)
        self.assertIsInstance(result, list)
        # No matching rows
        mock_cursor.fetchall.return_value = []
        self.assertEqual(get_failed_directories_for_reprocessing(), [])

    @patch('database.get_conn')
    def test_get_or_create_invalid_subject_cases(self, mock_get_conn):
        from database import get_or_create_invalid_subject
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate row exists
        mock_cursor.fetchone.side_effect = [[1], None, [2]]
        self.assertEqual(get_or_create_invalid_subject('a@b.com'), 1)
        # Simulate row does not exist, insert
        self.assertEqual(get_or_create_invalid_subject('b@c.com'), 2)

    @patch('database.get_conn')
    def test_insert_invalid_subject_detail_cases(self, mock_get_conn):
        from database import insert_invalid_subject_detail
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1]
        # All valid
        insert_invalid_subject_detail(1, 'subj', 'table')
        # Some None
        insert_invalid_subject_detail(1, 'subj', None)

    @patch('database.get_conn')
    def test_find_invalid_subject_detail_cases(self, mock_get_conn):
        from database import find_invalid_subject_detail
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.fetchone.return_value = [1, 'table']
        # Row exists
        result = find_invalid_subject_detail(1, 'subj')
        self.assertEqual(result, (1, 'table'))
        # Row does not exist
        mock_cursor.fetchone.return_value = None
        self.assertIsNone(find_invalid_subject_detail(1, 'subj'))

    @patch('database.get_conn')
    def test_set_jera_upload_flag_cases(self, mock_get_conn):
        from database import set_jera_upload_flag
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_cursor.rowcount = 1
        set_jera_upload_flag(1, True, 2)
        set_jera_upload_flag(1, False, None)
        mock_cursor.rowcount = 0
        set_jera_upload_flag(99, True, 2)

    @patch('database.get_conn')
    def test_get_pending_jera_uploads_cases(self, mock_get_conn):
        from database import get_pending_jera_uploads
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate DB rows
        mock_cursor.fetchall.return_value = [
            (1, 'subj', 'email', 2, None, 10, None, None)
        ]
        mock_cursor.description = [(col,) for col in ['id','subject','sender_email','jera_table_id','processed_at','total_rows','created_at','updated_at']]
        result = get_pending_jera_uploads()
        self.assertIsInstance(result, list)
        mock_cursor.fetchall.return_value = []
        self.assertEqual(get_pending_jera_uploads(), [])

    @patch('database.get_conn')
    def test_update_jera_upload_status_cases(self, mock_get_conn):
        from database import update_jera_upload_status
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate status update
        update_jera_upload_status(1, 'success')
        update_jera_upload_status(2, 'failed', {'error': 'fail'})
        update_jera_upload_status(3, 'pending')
        mock_conn.commit.assert_called()

    @patch('database.get_conn')
    def test_get_jera_upload_history_cases(self, mock_get_conn):
        from database import get_jera_upload_history
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate DB rows
        mock_cursor.fetchall.return_value = [
            (1, 'subj', 'email', 2, True, 'pending', None, None, 10, None, None)
        ]
        mock_cursor.description = [(col,) for col in ['id','subject','sender_email','jera_table_id','is_rate_approved_by_admin','jera_upload_status','jera_upload_result','jera_uploaded_at','total_rows','processed_at','created_at']]
        result = get_jera_upload_history(10)
        self.assertIsInstance(result, list)
        mock_cursor.fetchall.return_value = []
        self.assertEqual(get_jera_upload_history(1), [])

    @patch('database.get_conn')
    def test_mark_rate_upload_for_jera_cases(self, mock_get_conn):
        from database import mark_rate_upload_for_jera
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate found row
        mock_cursor.fetchone.return_value = [1]
        with patch('database.set_jera_upload_flag') as mock_set_flag:
            mock_set_flag.return_value = None
            rid = mark_rate_upload_for_jera('subj', 'email', 2)
            self.assertEqual(rid, 1)
        # Simulate not found
        mock_cursor.fetchone.return_value = None
        rid = mark_rate_upload_for_jera('notfound', 'email', 2)
        self.assertIsNone(rid)

    @patch('database.get_conn')
    def test_mark_comparison_file_for_bulk_upload_cases(self, mock_get_conn):
        from database import mark_comparison_file_for_bulk_upload
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        # Simulate found row
        mock_cursor.fetchone.return_value = [1]
        mock_cursor.description = [(col,) for col in ['id']]
        with patch('database.Json') as mock_json:
            mock_json.return_value = '{}'
            rid = mark_comparison_file_for_bulk_upload('file', 'subj', 'email', 2)
            self.assertEqual(rid, 1)
        # Simulate not found
        mock_cursor.fetchone.return_value = None
        rid = mark_comparison_file_for_bulk_upload('file', 'notfound', 'email', 2)
        self.assertIsNone(rid)
