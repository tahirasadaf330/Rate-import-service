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
        rate_upload_service.update_bulk_upload_status(1, 'completed', {'result': 'ok'})
        rate_upload_service.update_bulk_upload_status(1, 'failed', {'error': 'fail'})
        rate_upload_service.update_bulk_upload_status(1, 'processing')

    @patch('rate_upload_service.get_conn')
    def test_verify_upload_status(self, mock_get_conn):
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_cursor.fetchone.return_value = (1, 'completed', '2025-01-01', 'ok')
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor
        mock_get_conn.return_value.__enter__.return_value = mock_conn
        rate_upload_service.verify_upload_status(1)

    @patch('rate_upload_service.bulk_upload_comparison_to_jerasoft')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('os.path.exists')
    def test_process_bulk_upload_success(self, mock_exists, mock_update, mock_bulk):
        mock_exists.return_value = True
        mock_bulk.return_value = {'status': 'success', 'filtered_rows': 5, 'files_id': 'fileid123'}
        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 5, 'comparison_file_path': 'file.xlsx'
        }
        result = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertTrue(result)

    @patch('rate_upload_service.bulk_upload_comparison_to_jerasoft')
    @patch('rate_upload_service.update_bulk_upload_status')
    @patch('os.path.exists')
    def test_process_bulk_upload_failure(self, mock_exists, mock_update, mock_bulk):
        mock_exists.return_value = True
        mock_bulk.return_value = {'status': 'error', 'error': 'fail'}
        upload = {
            'id': 1, 'jera_table_id': 123, 'subject': 'subject', 'sender_email': 'sender',
            'total_rows': 5, 'comparison_file_path': 'file.xlsx'
        }
        result = rate_upload_service.process_bulk_upload(upload, dry_run=True)
        self.assertFalse(result)

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
