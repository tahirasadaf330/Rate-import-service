import unittest
from unittest.mock import patch, MagicMock
import jerasoft
import pandas as pd
from datetime import datetime
import os

class TestJerasoft(unittest.TestCase):
    def test_normalize_prefix(self):
        self.assertEqual(jerasoft.normalize_prefix('1001'), '1001')
        self.assertEqual(jerasoft.normalize_prefix(' 1001 '), '1001')
        self.assertEqual(jerasoft.normalize_prefix(None), None)
        self.assertEqual(jerasoft.normalize_prefix('abc'), None)
        self.assertEqual(jerasoft.normalize_prefix(''), None)

    def test_table_prefix_from_name(self):
        self.assertEqual(jerasoft.table_prefix_from_name('TERM Quickcom PRM trunk PREFIX:1001 USD'), '1001')
        self.assertEqual(jerasoft.table_prefix_from_name('PRFX-33'), '33')
        self.assertIsNone(jerasoft.table_prefix_from_name('NoPrefixHere'))

    def test_table_has_prefix(self):
        self.assertTrue(jerasoft.table_has_prefix('TERM Quickcom PRM trunk PREFIX:1001 USD', '1001'))
        self.assertFalse(jerasoft.table_has_prefix('TERM Quickcom PRM trunk PREFIX:1001 USD', '9999'))
        self.assertTrue(jerasoft.table_has_prefix('PRFX-33', '33'))
        self.assertFalse(jerasoft.table_has_prefix('PRFX-33', '44'))

    def test_normalize(self):
        self.assertEqual(jerasoft.normalize('  Foo.Bar  ##'), 'foo bar')
        self.assertEqual(jerasoft.normalize('Foo   Bar'), 'foo bar')
        self.assertEqual(jerasoft.normalize(''), '')

    def test_fuzzy_score(self):
        self.assertGreaterEqual(jerasoft.fuzzy_score('foo', 'foo'), 0.99)
        self.assertLess(jerasoft.fuzzy_score('foo', 'bar'), 0.5)
        self.assertGreater(jerasoft.fuzzy_score('foo bar', 'bar foo'), 0.5)

    def test_name_starts_with_term(self):
        self.assertTrue(jerasoft.name_starts_with_term('TERM Quickcom'))
        self.assertFalse(jerasoft.name_starts_with_term('Other Quickcom'))

    def test_name_contains_company(self):
        self.assertTrue(jerasoft.name_contains_company('TERM Quickcom', 'quickcom'))
        self.assertFalse(jerasoft.name_contains_company('TERM Other', 'quickcom'))

    def test_extract_company_keyword(self):
        self.assertEqual(jerasoft.extract_company_keyword('Quickcom tel PRM trunk Prefix:1001 USD'), 'quickcom')
        self.assertEqual(jerasoft.extract_company_keyword(''), '')

    @patch('jerasoft._session.post')
    def test__post_json_success(self, mock_post):
        mock_resp = MagicMock()
        mock_resp.json.return_value = {'result': 'ok'}
        mock_resp.headers = {'Content-Type': 'application/json'}
        mock_resp.raise_for_status = MagicMock()
        mock_post.return_value = mock_resp
        result = jerasoft._post_json('http://fakeurl', {'foo': 'bar'})
        self.assertEqual(result, {'result': 'ok'})

    @patch('jerasoft._session.post')
    def test__post_json_jsondecodeerror(self, mock_post):
        mock_resp = MagicMock()
        mock_resp.json.side_effect = jerasoft.JSONDecodeError('Expecting value', '', 0)
        mock_resp.text = 'bad json'
        mock_resp.headers = {'Content-Type': 'application/json', 'Content-Length': '8'}
        mock_resp.raise_for_status = MagicMock()
        mock_post.return_value = mock_resp
        with self.assertRaises(RuntimeError):
            jerasoft._post_json('http://fakeurl', {'foo': 'bar'})

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_exact(self, mock_fetch):
        mock_fetch.return_value = [{'id': 1, 'name': 'TERM Quickcom'}]
        self.assertEqual(jerasoft.get_table_id_by_name('TERM Quickcom'), 1)

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_fuzzy(self, mock_fetch):
        mock_fetch.return_value = [{'id': 2, 'name': 'TERM Quickcom'}]
        self.assertIsNone(jerasoft.get_table_id_by_name('TERM QuickcomX'))

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_not_found(self, mock_fetch):
        mock_fetch.return_value = [{'id': 3, 'name': 'TERM Other'}]
        self.assertIsNone(jerasoft.get_table_id_by_name('TERM Quickcom'))

    @patch('jerasoft._post_json')
    def test_fetch_all_tables(self, mock_post):
        mock_post.side_effect = [
            {'result': [{'id': 1, 'name': 'TERM Quickcom'}]},
            {'result': []}
        ]
        tables = jerasoft.fetch_all_tables(api_key='dummy')
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0]['id'], 1)

    @patch('jerasoft._post_json')
    def test_fetch_active_current_future_rates(self, mock_post):
        mock_post.side_effect = [
            {'result': [{'code': '001', 'code_name': 'US', 'value': 0.1, 'effective_from': '2025-01-01', 'min_volume': 1, 'pay_interval': 60}]},
            {'result': []}
        ]
        df = jerasoft.fetch_active_current_future_rates(1, api_key='dummy')
        self.assertIsInstance(df, pd.DataFrame)
        self.assertEqual(df.shape[0], 1)
        self.assertIn('Dst Code', df.columns)

    def test_save_rates_to_excel(self):
        df = pd.DataFrame({'Dst Code': ['001'], 'Rate': [0.1]})
        path = jerasoft.save_rates_to_excel(df, 'test_rates.xlsx')
        self.assertTrue(path.endswith('test_rates.xlsx'))
        self.assertTrue(os.path.exists(path))
        os.remove(path)

    @patch('jerasoft.get_table_id_by_name')
    @patch('jerasoft.fetch_active_current_future_rates')
    @patch('jerasoft.save_rates_to_excel')
    def test_export_rates_by_query_force_table(self, mock_save, mock_fetch, mock_get_id):
        mock_get_id.return_value = 42
        mock_fetch.return_value = pd.DataFrame({'Dst Code': ['001'], 'Rate': [0.1]})
        mock_save.return_value = 'dummy.xlsx'
        result = jerasoft.export_rates_by_query(
            target_query='Quickcom tel PRM trunk Prefix:1001 USD',
            output_path='dummy.xlsx',
            subject='Quickcom tel PRM trunk Prefix:1001 USD',
            force_table_name='TERM Quickcom',
            api_key='dummy',
        )
        self.assertEqual(result['table_id'], 42)
        self.assertEqual(result['saved_to'], 'dummy.xlsx')

    def test_is_valid_subject(self):
        # Valid subjects (4 bracketed parts)
        self.assertTrue(jerasoft.is_valid_subject('[COMPANY] [TRUNK] [PREFIX] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[A] [B] [C] [D]'))
        self.assertTrue(jerasoft.is_valid_subject('[TELMOBIL] [TRUNK] [44128] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[QUICKCOM] [STANDARD] [Prefix 002] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[ALLIP][CLI][1072#][USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[TELEGEEKS] [CC] [333] [USD]'))
        # Invalid subjects (not 4 bracketed parts)
        self.assertFalse(jerasoft.is_valid_subject('[TELMOBIL] [44128] [USD]'))  # only 3
        self.assertFalse(jerasoft.is_valid_subject('[COMPANY] [TRUNK] [PREFIX]'))  # only 3
        self.assertFalse(jerasoft.is_valid_subject('COMPANY TRUNK PREFIX USD'))  # no brackets
        self.assertFalse(jerasoft.is_valid_subject(''))  # empty
        self.assertFalse(jerasoft.is_valid_subject(None))  # None

    def test_is_valid_subject_examples(self):
        # Valid (4 bracketed parts)
        self.assertTrue(jerasoft.is_valid_subject('[CALLCARIBE INC] [NCLI] [NONE] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[TELEGEEKS] [CC] [333] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[ALLIP][CLI][1072#][USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[Vasudev Global Pte Ltd.] [Prime] [11] [USD]'))
        self.assertTrue(jerasoft.is_valid_subject('[CIMA TELECOM] [HAYOTEL] [75433] [USD]'))
        # Invalid (not 4 bracketed parts)
        self.assertFalse(jerasoft.is_valid_subject('[IPBTEL] [HAYO CC] [53956 [USD]'))  # missing closing bracket
        self.assertFalse(jerasoft.is_valid_subject('[42com] [HAYO INT PRM USD] [] [USD]'))  # empty 3rd part
        self.assertFalse(jerasoft.is_valid_subject('HAYO:: CLI PARTIAL RATE SHEET || PREFIX 111'))  # no brackets
        self.assertFalse(jerasoft.is_valid_subject('[ADC][WHS][][USD]'))  # empty 3rd part
        self.assertFalse(jerasoft.is_valid_subject('[TELMOBIL] [44128] [USD]'))  # only 3 parts
        self.assertFalse(jerasoft.is_valid_subject('Pricelist_11/23/2025 8:13:06 AM_HAYO (Premium)_Full'))  # no brackets

    def test_print_subject_validity(self):
        subjects = [
            '[CALLCARIBE INC] [NCLI] [NONE] [USD]',
            '[IPBTEL] [HAYO CC] [53956 [USD]',
            '[42com] [HAYO INT PRM USD] [] [USD]',
            'HAYO:: CLI PARTIAL RATE SHEET || PREFIX 111',
            '[TELEGEEKS] [CC] [333] [USD]',
            '[ALLIP][CLI][1072#][USD]',
            '[Vasudev Global Pte Ltd.] [Prime] [11] [USD]',
            '[ADC][WHS][][USD]',
            '[TELMOBIL] [44128] [USD]',
            '[CIMA TELECOM] [HAYOTEL] [75433] [USD]',
            'Pricelist_11/23/2025 8:13:06 AM_HAYO (Premium)_Full',
        ]
        for subj in subjects:
            print(f"{subj!r} => {'VALID' if jerasoft.is_valid_subject(subj) else 'INVALID'}")

    def test_is_valid_subject_expected_results(self):
        cases = [
            ('[CALLCARIBE INC] [NCLI] [NONE] [USD]', True),
            ('[IPBTEL] [HAYO CC] [53956 [USD]', False),
            ('[42com] [HAYO INT PRM USD] [] [USD]', False),
            ('HAYO:: CLI PARTIAL RATE SHEET || PREFIX 111', False),
            ('[TELEGEEKS] [CC] [333] [USD]', True),
            ('[ALLIP][CLI][1072#][USD]', True),
            ('[Vasudev Global Pte Ltd.] [Prime] [11] [USD]', True),
            ('[ADC][WHS][][USD]', False),
            ('[TELMOBIL] [44128] [USD]', False),
            ('[CIMA TELECOM] [HAYOTEL] [75433] [USD]', True),
            ('Pricelist_11/23/2025 8:13:06 AM_HAYO (Premium)_Full', False),
        ]
        for subj, expected in cases:
            self.assertEqual(jerasoft.is_valid_subject(subj), expected, f"Subject: {subj!r} should be {'VALID' if expected else 'INVALID'}")

if __name__ == '__main__':
    unittest.main()
