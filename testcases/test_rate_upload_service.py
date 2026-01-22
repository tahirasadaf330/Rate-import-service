import unittest
from unittest.mock import patch, MagicMock
import rate_upload_service

class TestRateUploadService(unittest.TestCase):
    @patch('rate_upload_service.get_conn')
    def test_get_pending_bulk_uploads(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_cursor.fetchall.return_value = [
            (1, 'subject', 'sender', 123, 'pending_bulk', None, 10, '2025-01-01', 'file.xlsx')
        ]
        mock_cursor.description = [
            ('id',), ('subject',), ('sender_email',), ('jera_table_id',), ('jera_upload_status',),
            ('jera_upload_result',), ('total_rows',), ('created_at',), ('comparison_file_path',)
        ]
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        uploads = rate_upload_service.get_pending_bulk_uploads()
        self.assertEqual(len(uploads), 1)
        self.assertEqual(uploads[0]['id'], 1)

    @patch('rate_upload_service.get_conn')
    def test_update_bulk_upload_status(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_cursor.rowcount = 1
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        # On success: jera_upload_result should be NULL/None (we store nothing on success)
        rate_upload_service.update_bulk_upload_status(1, 'completed', {'result': 'ok'})
        args1 = mock_cursor.execute.call_args_list[-1][0][1]
        self.assertEqual(args1[0], 'completed')
        self.assertIsNone(args1[1])

        # On failure: jera_upload_result should be valid JSON string
        rate_upload_service.update_bulk_upload_status(1, 'failed', {'error': 'fail'})
        args2 = mock_cursor.execute.call_args_list[-1][0][1]
        self.assertEqual(args2[0], 'failed')
        self.assertIsInstance(args2[1], str)
        self.assertIn('"error"', args2[1])

        # On processing: jera_upload_result should be NULL/None
        rate_upload_service.update_bulk_upload_status(1, 'processing')
        args3 = mock_cursor.execute.call_args_list[-1][0][1]
        self.assertEqual(args3[0], 'processing')
        self.assertIsNone(args3[1])

    @patch('rate_upload_service.get_conn')
    def test_verify_upload_status(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_cursor.fetchone.return_value = (1, 'completed', '2025-01-01', 'ok')
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        rate_upload_service.verify_upload_status(1)

    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_success(self, mock_export, mock_update, mock_fetch, mock_bulk):
        mock_fetch.return_value = [
            {"dst_code": "1001", "rate_new": 0.05, "effective_date": "2025-01-01", "status": "Accepted"}
        ]
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 1, 'files_id': 'fileid123'}
        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 5, 'comparison_file_path': 'file.xlsx'
        }
        result = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(result)
        self.assertTrue(mock_export.called)

    @patch('rate_upload_service.stash_future_rates_df_to_jerasoft')
    @patch('rate_upload_service.wait_for_imported_rates_visible')
    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_stash_future_rates_called_on_success(self, mock_export, mock_update, mock_fetch, mock_bulk, mock_wait, mock_stash_future):
        mock_fetch.return_value = [
            {"dst_code": "1001", "code_name": "X", "rate_new": 0.05, "effective_date": "2025-01-01", "status": "Accepted"}
        ]
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 1, 'files_id': 'fileid123'}
        mock_wait.return_value = {"status": "ready", "found": 1, "need": 1, "total_pairs": 1, "waited_sec": 0.0}
        mock_stash_future.return_value = {"status": "success", "stashed": 1, "errors": []}

        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 1, 'comparison_file_path': 'file.xlsx'
        }
        with patch.dict('os.environ', {'JERASOFT_ENABLE_STASH_FUTURE_RATES': '1'}):
            ok = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(ok)
        self.assertTrue(mock_stash_future.called)

    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_closed_rows_are_sent_as_close_keyword(self, mock_export, mock_update, mock_fetch, mock_bulk):
        # DB rows include a Closed row (Accepted) with a numeric rate_new; service should convert it to "close"
        mock_fetch.return_value = [
            {"dst_code": "43", "code_name": "AUSTRIA", "rate_new": None, "effective_date": "2026-01-15", "new_billing_increment": "1/1",
             "status": "Accepted", "change_type": "Closed", "notes": "present in current but missing in new"},
        ]
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 1, 'files_id': 'fileid123'}

        upload = {
            'id': 1, 'jera_table_id': 4330, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 1, 'comparison_file_path': 'file.xlsx'
        }
        ok = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(ok)

        # Verify the DF passed to bulk_upload_df_to_jerasoft has "close" keyword
        passed_df = mock_bulk.call_args.kwargs["df"]
        self.assertIn("New Rate", passed_df.columns)
        self.assertIn("Change Type", passed_df.columns)

        closed_row = passed_df[passed_df["Change Type"].astype(str).str.lower().eq("closed")].iloc[0]
        self.assertEqual(str(closed_row["New Rate"]).strip().lower(), "close")

    @patch('rate_upload_service.stash_rates_df_to_jerasoft')
    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_stashed_rows_are_applied_via_api(self, mock_export, mock_update, mock_fetch, mock_bulk, mock_stash):
        mock_fetch.return_value = [
            {"dst_code": "43", "code_name": "AUSTRIA", "rate_new": 0.56, "effective_date": "2026-01-21", "new_billing_increment": "1/1",
             "status": "Accepted", "change_type": "Stashed", "notes": "stash this rate"},
        ]
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 0, 'files_id': 'fileid123'}
        mock_stash.return_value = {"status": "success", "stashed": 1, "not_found": 0, "errors": []}

        upload = {
            'id': 1, 'jera_table_id': 4330, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 1, 'comparison_file_path': 'file.xlsx'
        }
        ok = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(ok)
        self.assertTrue(mock_stash.called)

    @patch('rate_upload_service.stash_rates_df_to_jerasoft')
    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_stash_only_is_success(self, mock_export, mock_update, mock_fetch, mock_bulk, mock_stash):
        # Only 1 accepted row and it's Stashed → bulk import should be "skipped/no_accepted_rates"
        mock_fetch.return_value = [
            {"dst_code": "234708", "code_name": "X", "rate_new": None, "effective_date": "2026-01-23", "new_billing_increment": None,
             "status": "Accepted", "change_type": "Stashed", "notes": "stash old future rate"},
        ]
        mock_bulk.return_value = {"status": "skipped", "reason": "no_accepted_rates", "original_rows": 0}
        mock_stash.return_value = {"status": "success", "stashed": 1, "not_found": 0, "errors": []}

        upload = {
            'id': 205, 'jera_table_id': 4330, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 1, 'comparison_file_path': 'file.xlsx'
        }
        ok = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(ok)
        self.assertTrue(mock_stash.called)

    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_failure(self, mock_export, mock_update, mock_fetch, mock_bulk):
        mock_fetch.return_value = [
            {"dst_code": "1001", "rate_new": 0.05, "effective_date": "2025-01-01", "status": "Accepted"}
        ]
        mock_bulk.return_value = {'status': 'error', 'error': 'fail'}
        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 5, 'comparison_file_path': 'file.xlsx'
        }
        result = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertFalse(result)
        self.assertTrue(mock_export.called)

    @patch('rate_upload_service.wait_for_imported_rates_visible')
    @patch('rate_upload_service.bulk_upload_df_to_jerasoft')
    @patch('rate_upload_service.fetch_rate_upload_details_for_upload')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('rate_upload_service.export_db_upload_dataframe')
    def test_process_bulk_upload_waits_for_import_apply_when_enabled(self, mock_export, mock_update, mock_fetch, mock_bulk, mock_wait):
        mock_fetch.return_value = [
            {"dst_code": "1001", "code_name": "X", "rate_new": 0.05, "effective_date": "2025-01-01", "status": "Accepted"}
        ]
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 1, 'files_id': 'fileid123'}
        mock_wait.return_value = {"status": "ready", "found": 1, "need": 1, "total_pairs": 1, "waited_sec": 0.0}

        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 1, 'comparison_file_path': 'file.xlsx'
        }
        with patch.dict('os.environ', {'JERASOFT_WAIT_IMPORT_APPLY': '1'}):
            ok = rate_upload_service.process_bulk_upload(upload, dry_run=False)
        self.assertTrue(ok)
        self.assertTrue(mock_wait.called)

    @patch('rate_upload_service.get_pending_bulk_uploads')
    def test_list_pending_uploads(self, mock_get_pending):
        mock_get_pending.return_value = [
            {'id': 1, 'subject': 'subject', 'sender_email': 'sender', 'jera_table_id': 123,
             'total_rows': 5, 'created_at': '2025-01-01'}
        ]
        rate_upload_service.list_pending_uploads()

    @patch('rate_upload_service.get_pending_bulk_uploads')
    @patch('rate_upload_service.process_bulk_upload')
    def test_process_all_uploads(self, mock_process, mock_get_pending):
        mock_get_pending.return_value = [
            {'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
             'total_rows': 5, 'comparison_file_path': 'file.xlsx'}
        ]
        mock_process.return_value = True
        results = rate_upload_service.process_all_uploads(dry_run=True)
        self.assertEqual(results['success'], 1)
        self.assertEqual(results['failed'], 0)

if __name__ == '__main__':
    unittest.main()
