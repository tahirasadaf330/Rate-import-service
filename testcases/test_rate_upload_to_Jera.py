import unittest
from unittest.mock import patch, MagicMock
import pandas as pd
import rate_upload_to_Jera

class TestRateUploadToJera(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame({
            'Code': ['1', '44'],
            'New Rate': [0.05, 0.03],
            'Effective Date': ['2024-01-01', '2024-01-01'],
            'New Billing Increment': ['60/60', '1/1'],
            'Status': ['Accepted', 'Rejected']
        })

    def test_parse_billing_increment(self):
        self.assertEqual(rate_upload_to_Jera._parse_billing_increment('60/60'), (60, 60))
        self.assertEqual(rate_upload_to_Jera._parse_billing_increment('1/1'), (1, 1))
        self.assertEqual(rate_upload_to_Jera._parse_billing_increment('bad'), (None, None))
        self.assertEqual(rate_upload_to_Jera._parse_billing_increment(None), (None, None))

    def test_fmt_date(self):
        self.assertEqual(rate_upload_to_Jera._fmt_date('2024-01-01'), '2024-01-01')
        self.assertIsNone(rate_upload_to_Jera._fmt_date(None))
        self.assertIsNone(rate_upload_to_Jera._fmt_date(''))

    @patch('rate_upload_to_Jera._rpc_call')
    def test_search_rate_id(self, mock_rpc):
        mock_rpc.return_value = [{'id': 123, 'effective_from': '2024-01-01'}]
        rid = rate_upload_to_Jera._search_rate_id(1, '1', '2024-01-01')
        self.assertEqual(rid, 123)
        mock_rpc.return_value = []
        rid = rate_upload_to_Jera._search_rate_id(1, '1', '2024-01-01')
        self.assertIsNone(rid)

    @patch('rate_upload_to_Jera._rpc_call')
    def test_create_rate(self, mock_rpc):
        mock_rpc.return_value = {'id': 456}
        rid = rate_upload_to_Jera._create_rate(1, '1', 0.05, '2024-01-01', 60, 60)
        self.assertEqual(rid, 456)
        mock_rpc.return_value = [{}]
        with self.assertRaises(RuntimeError):
            rate_upload_to_Jera._create_rate(1, '1', 0.05, '2024-01-01', 60, 60)

    @patch('rate_upload_to_Jera._rpc_call')
    def test_update_rate(self, mock_rpc):
        mock_rpc.return_value = None
        # Should not raise
        rate_upload_to_Jera._update_rate(123, value=0.05, min_vol=60, pay_int=60, status='active')

    @patch('rate_upload_to_Jera._rpc_call')
    def test_push_comparison_to_jerasoft(self, mock_rpc):
        rate_upload_to_Jera.J_API_KEY = 'dummykey'
        mock_rpc.side_effect = [[], {'id': 789}, None]
        df = pd.DataFrame({
            'Code': ['1'],
            'New Rate': [0.05],
            'Effective Date': ['2024-01-01'],
            'New Billing Increment': ['60/60'],
            'Status': ['Accepted']
        })
        result = rate_upload_to_Jera.push_comparison_to_jerasoft(df, 1, dry_run=True)
        self.assertIn('processed', result)
        self.assertEqual(result['processed'], 1)

    @patch('rate_upload_to_Jera.upload_file_to_jerasoft')
    @patch('rate_upload_to_Jera._rpc_call')
    @patch('os.path.exists')
    @patch('os.remove')
    def test_bulk_import_rates(self, mock_remove, mock_exists, mock_rpc, mock_upload):
        rate_upload_to_Jera.J_API_KEY = 'dummykey'
        mock_upload.return_value = 'fileid123'
        mock_rpc.side_effect = [
            {'id': 'queueid123'},
            {'result': 'ok'}
        ]
        mock_exists.return_value = True
        df = pd.DataFrame({
            'Code': ['1'],
            'New Rate': [0.05],
            'Effective Date': ['2024-01-01'],
            'Status': ['Accepted']
        })
        df['code'] = df['Code']
        df['code_name'] = df['Code']
        df['value'] = df['New Rate']
        df['effective_from'] = pd.to_datetime(df['Effective Date']).dt.strftime('%Y-%m-%d')
        result = rate_upload_to_Jera.bulk_import_rates(df, 1, dry_run=True)
        self.assertIn('status', result)
        self.assertEqual(result['status'], 'dry_run')

    @patch('rate_upload_to_Jera.bulk_import_rates')
    @patch('pandas.read_excel')
    def test_bulk_upload_comparison_to_jerasoft(self, mock_read_excel, mock_bulk):
        mock_read_excel.return_value = pd.DataFrame({
            'Code': ['1'],
            'New Rate': [0.05],
            'Effective Date': ['2024-01-01'],
            'Status': ['Accepted']
        })
        mock_bulk.return_value = {'status': 'success', 'rows_uploaded': 1}
        result = rate_upload_to_Jera.bulk_upload_comparison_to_jerasoft('file.xlsx', 1)
        self.assertEqual(result['status'], 'success')
        self.assertEqual(result['rows_uploaded'], 1)

if __name__ == '__main__':
    unittest.main()
