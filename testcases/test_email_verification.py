import unittest
from unittest.mock import patch, MagicMock
import email_verification
import json
from pathlib import Path
from datetime import datetime, timezone
class TestEmailVerification(unittest.TestCase):
    def test_get_token_timeout_and_errors(self):
        from unittest.mock import patch, MagicMock
        import requests
        # Simulate timeout
        with patch('requests.post') as mock_post, patch('logging.getLogger') as mock_logger:
            mock_post.side_effect = requests.exceptions.Timeout('timeout')
            logger = MagicMock()
            mock_logger.return_value = logger
            with self.assertRaises(requests.exceptions.Timeout):
                email_verification.get_token('tenant', 'client', 'secret')
            logger.error.assert_any_call('Token request timed out: timeout')

        # Simulate request error
        with patch('requests.post') as mock_post, patch('logging.getLogger') as mock_logger:
            mock_post.side_effect = requests.exceptions.RequestException('fail')
            logger = MagicMock()
            mock_logger.return_value = logger
            with self.assertRaises(requests.exceptions.RequestException):
                email_verification.get_token('tenant', 'client', 'secret')
            logger.error.assert_any_call("Token request failed: RequestException('fail')")

        # Simulate non-200 response
        with patch('requests.post') as mock_post, patch('logging.getLogger') as mock_logger:
            resp = MagicMock()
            resp.status_code = 400
            resp.json.return_value = {'error': 'invalid'}
            resp.text = 'bad request'
            mock_post.return_value = resp
            logger = MagicMock()
            mock_logger.return_value = logger
            with self.assertRaises(RuntimeError):
                email_verification.get_token('tenant', 'client', 'secret')
            logger.error.assert_any_call("Token error: " + json.dumps({'error': 'invalid'}, indent=2))

        # Simulate missing access_token
        with patch('requests.post') as mock_post, patch('logging.getLogger') as mock_logger:
            resp = MagicMock()
            resp.status_code = 200
            resp.json.return_value = {'not_access_token': 'nope'}
            mock_post.return_value = resp
            logger = MagicMock()
            mock_logger.return_value = logger
            with self.assertRaises(RuntimeError):
                email_verification.get_token('tenant', 'client', 'secret')
            logger.error.assert_any_call("Token response missing access_token: {'not_access_token': 'nope'}")
    def test_load_failed_log_and_atomic_write(self):
        # Should load default schema if file missing
        with patch('email_verification.FAILED_EMAILS_PATH', Path('nonexistent_failed_emails.json')):
            log = email_verification._load_failed_log()
            self.assertIsInstance(log, dict)
            self.assertIn('version', log)
            # Should write and reload
            email_verification._atomic_write_failed_log(log)
            loaded = email_verification._load_failed_log()
            self.assertEqual(loaded['version'], 1)

    def test_entry_exists(self):
        bucket = [{'id': 'abc'}, {'internetMessageId': 'xyz'}]
        self.assertTrue(email_verification._entry_exists(bucket, 'abc'))
        self.assertTrue(email_verification._entry_exists(bucket, 'xyz'))
        self.assertFalse(email_verification._entry_exists(bucket, 'notfound'))

    def test_log_failed_email(self):
        payload = {'id': 'testid', 'internetMessageId': 'testmsgid'}
        with patch('email_verification._atomic_write_failed_log') as mock_write:
            email_verification.log_failed_email('unverified_sender', payload)
            mock_write.assert_called()

    def test_normalize_subject(self):
        # None input
        self.assertIsNone(email_verification._normalize_subject(None))
        # Empty string
        self.assertIsNone(email_verification._normalize_subject(''))
        # Only spaces returns empty string
        self.assertEqual(email_verification._normalize_subject('   '), '')
        # Normal subject
        self.assertEqual(email_verification._normalize_subject('Test Subject'), 'Test Subject')
        # Subject with special characters
        self.assertEqual(email_verification._normalize_subject('Test:Subject!'), 'Test Subject')
        # Subject with multiple separators
        self.assertEqual(email_verification._normalize_subject('A:B;C|D,E/F\G'), 'A B C D E F G')
        # Subject with unicode normalization
        self.assertEqual(email_verification._normalize_subject('Tést Sübjèct'), 'T st S bj ct')
        # Subject with excessive whitespace
        self.assertEqual(email_verification._normalize_subject('   Test    Subject   '), 'Test Subject')
        # Subject with only special characters
        self.assertEqual(email_verification._normalize_subject('!@#$%^&*()'), '')

    def test_first_3letter_currency(self):
        self.assertEqual(email_verification._first_3letter_currency(['USD', 'EUR', 'GBP']), 'GBP')
        self.assertEqual(email_verification._first_3letter_currency(['foo', 'bar']), 'BAR')

    def test_normalize_output(self):
        out = email_verification._normalize_output('Comp', 'Trunk', '123', 'USD')
        self.assertEqual(out['company'], 'Comp')
        self.assertEqual(out['trunk'], 'Trunk')
        self.assertEqual(out['prefix'], '123')  # prefix stored as string
        self.assertEqual(out['currency'], 'USD')
        self.assertIsNone(email_verification._normalize_output('', '', '', ''))

    def test_extract_anyorder(self):
        subj = 'Acme trunk prefix 123 USD'
        result = email_verification._extract_anyorder(subj)
        self.assertIsInstance(result, dict)
        self.assertEqual(result['currency'], 'USD')
        self.assertIsNone(email_verification._extract_anyorder('no currency here'))

    def test_validate_subject(self):
        # Strict 4-bracket subjects are valid
        v1 = email_verification.validate_subject('[SIGMA] [Gold] [3333] [USD]')
        self.assertIsInstance(v1, dict)
        self.assertEqual(v1["company"], "SIGMA")
        self.assertEqual(v1["trunk"], "Gold")
        self.assertEqual(v1["prefix"], "3333")
        self.assertEqual(v1["currency"], "USD")

        # Prefix numeric extraction (preserve leading zeros)
        v2 = email_verification.validate_subject('[QUICKCOM] [CC] [Prefix 004] [USD]')
        self.assertIsInstance(v2, dict)
        self.assertEqual(v2["prefix"], "004")

        # Prefix numeric extraction (suffix allowed)
        v3 = email_verification.validate_subject('[Asia Access Telecom] [Hayo Telecom IN] [1117#] [USD]')
        self.assertIsInstance(v3, dict)
        self.assertEqual(v3["prefix"], "1117")

        # Prefix textual "None" now maps to string "NONE"
        v4 = email_verification.validate_subject('[World Hub Communications Pte. Ltd.] [Premium] [None] [USD]')
        self.assertIsInstance(v4, dict)
        self.assertEqual(v4["prefix"], "NONE")

        # Prefix can also be expressed as "PREFIX NONE" (but not other phrases)
        v4b = email_verification.validate_subject('[TITAN INTERNATIONAL] [STANDARD] [PREFIX NONE] [USD]')
        self.assertIsInstance(v4b, dict)
        self.assertEqual(v4b["company"], "TITAN INTERNATIONAL")
        self.assertEqual(v4b["trunk"], "STANDARD")
        self.assertEqual(v4b["prefix"], "NONE")
        self.assertEqual(v4b["currency"], "USD")

        # Freeform format is also valid:
        #   "<company words> <trunk> trunk Prefix:1001 USD"
        v5 = email_verification.validate_subject('Quickcom tel PRM trunk Prefix:1001 USD')
        self.assertIsInstance(v5, dict)
        self.assertEqual(v5["company"], "Quickcom tel")
        self.assertEqual(v5["trunk"], "PRM")
        self.assertEqual(v5["prefix"], "1001")
        self.assertEqual(v5["currency"], "USD")

        v6 = email_verification.validate_subject('Quickcom tel PRM trunk Prefix 0007 usd')
        self.assertIsInstance(v6, dict)
        self.assertEqual(v6["company"], "Quickcom tel")
        self.assertEqual(v6["trunk"], "PRM")
        self.assertEqual(v6["prefix"], "0007")
        self.assertEqual(v6["currency"], "USD")

        # Freeform with textual none should map to prefix "NONE"
        v8 = email_verification.validate_subject('Vendor X PRM trunk Prefix none usd')
        self.assertIsInstance(v8, dict)
        self.assertEqual(v8["company"], "Vendor X")
        self.assertEqual(v8["trunk"], "PRM")
        self.assertEqual(v8["prefix"], "NONE")
        self.assertEqual(v8["currency"], "USD")

        # Invalid: not in strict bracket format
        # (This does NOT match freeform either because it lacks "<TRUNK> trunk" and the prefix/currency framing)
        self.assertIsNone(email_verification.validate_subject('Acme trunk prefix 123 USD'))
        self.assertIsNone(email_verification.validate_subject('bad subject'))

        # Invalid: empty brackets
        self.assertIsNone(email_verification.validate_subject('[ALLIP][][1072][USD]'))
        self.assertIsNone(email_verification.validate_subject('[ALLIP][ ][1072][USD]'))
        self.assertIsNone(email_verification.validate_subject('[][CLI][1072][USD]'))

        # Invalid: missing currency / extra trailing text
        self.assertIsNone(email_verification.validate_subject('[SIGMA] [Gold] [3333]'))
        self.assertIsNone(email_verification.validate_subject('[SIGMA] [Gold] [3333] [USD] 2025-01-01'))
        self.assertIsNone(email_verification.validate_subject('Quickcom tel PRM trunk Prefix:1001 USD 2025-01-01'))  # trailing tokens
        # Freeform: trunk is the token immediately before the keyword "trunk"
        v7 = email_verification.validate_subject('Quickcom tel trunk Prefix:1001 USD')
        self.assertIsInstance(v7, dict)
        self.assertEqual(v7["company"], "Quickcom")
        self.assertEqual(v7["trunk"], "tel")
        self.assertEqual(v7["prefix"], "1001")
        self.assertEqual(v7["currency"], "USD")

        # User-provided invalid subjects (should go to invalid-subject flow)
        self.assertIsNone(email_verification.validate_subject('PHOENOS_RCN-Hayo,2026-01-10'))
        self.assertIsNone(email_verification.validate_subject('[ADC][WHS][][USD]'))  # empty prefix bracket
        self.assertIsNone(email_verification.validate_subject('HAYO:: CLI PARTIAL RATE SHEET || PREFIX 111'))  # no currency and no "trunk" keyword
        self.assertIsNone(email_verification.validate_subject('A-Z Replacement Request - [RICOCHET][WHL][1771][USD]'))  # extra text outside brackets
        self.assertIsNone(email_verification.validate_subject('New A-Z rate update from SwissLink for CLI Services SURCHARG...'))

    def test_subject_ok(self):
        self.assertTrue(email_verification.subject_ok('[SIGMA] [Gold] [3333] [USD]'))
        self.assertTrue(email_verification.subject_ok('Quickcom tel PRM trunk Prefix:1001 USD'))
        self.assertFalse(email_verification.subject_ok('Acme trunk prefix 123 USD'))
        self.assertFalse(email_verification.subject_ok('bad subject'))

    def test_strip_date_time_tokens_for_invalid_subject(self):
        cases = [
            {
                "name": "month_name_date_with_time_and_tz",
                "inp": "ECOCARRIER LATEST OFFER RATES TO HAYO TELECOM, INC - FULL REPLACEMENT with prefix 8300 effective immediately December 12, 2025 14:04 GMT",
                "contains": [
                    # punctuation like ',' and '-' are removed by subject cleanup, so assert on a punctuation-free core
                    "ECOCARRIER LATEST OFFER RATES TO HAYO TELECOM INC FULL REPLACEMENT with prefix 8300 effective immediately",
                ],
                "not_contains": ["December", "2025", "14:04", "14 04", "GMT"],
            },
            {
                "name": "iso_datetime_z",
                "inp": "Offer update prefix 8300 2025-12-01T15:41:59Z",
                "equals": "Offer update prefix 8300",
            },
            {
                "name": "iso_datetime_offset",
                "inp": "Offer update prefix 8300 2025-12-01 15:41:59+05:00",
                "equals": "Offer update prefix 8300",
            },
            {
                "name": "numeric_date_ymd_with_dots",
                "inp": "Acme prefix 123 effective 2025.12.01",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "numeric_date_dmy_with_slashes",
                "inp": "Acme prefix 123 effective 01/12/2025",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "numeric_date_ymd_with_slashes",
                "inp": "Acme prefix 123 effective 2025/12/01",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "compact_yyyymmdd",
                "inp": "Acme prefix 123 effective 20251201",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "day_name_and_month_name_date",
                "inp": "Acme prefix 123 effective Mon December 1 2025 09:00 UTC",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "hyphenated_month_name_date",
                "inp": "Acme prefix 123 effective 01-Dec-2025 09:00",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "month_year_only",
                "inp": "Acme prefix 123 effective Dec 2025",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "am_pm_time",
                "inp": "Acme prefix 123 effective December 1, 2025 9:15 PM GMT",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "hhmmss_compact_time",
                "inp": "Acme prefix 123 effective 2025-12-01 154159 GMT",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "offset_token_alone",
                "inp": "Acme prefix 123 effective 2025-12-01 +05:00",
                "equals": "Acme prefix 123 effective",
            },
            {
                "name": "tid_block_stripped",
                "inp": "New Voice Rates for Hayo Telecom Inc. Platinum [TID:4445188]",
                "equals": "New Voice Rates for Hayo Telecom Inc. Platinum",
            },
            {
                "name": "tid_block_stripped_with_extra_note",
                "inp": "New Voice Rates for Hayo Telecom Inc. Gold (Open RTP) [TID:4445265]",
                "equals": "New Voice Rates for Hayo Telecom Inc. Gold (Open RTP)",
            },
            {
                "name": "tid_block_stripped_simple_gold",
                "inp": "New Voice Rates for Hayo Telecom Inc. Gold [TID:4445272]",
                "equals": "New Voice Rates for Hayo Telecom Inc. Gold",
            },
        ]

        for c in cases:
            with self.subTest(case=c["name"]):
                out = email_verification._strip_date_time_tokens_for_invalid_subject(c["inp"])
                if "equals" in c:
                    self.assertEqual(out, c["equals"])
                for needle in c.get("contains", []):
                    self.assertIn(needle, out)
                for bad in c.get("not_contains", []):
                    self.assertNotIn(bad, out)

    def test_strip_date_time_tokens_for_invalid_subject_iso(self):
        s = "Offer update prefix 8300 2025-12-01T15:41:59Z"
        out = email_verification._strip_date_time_tokens_for_invalid_subject(s)
        self.assertEqual(out, "Offer update prefix 8300")

    def test_dbg(self):
        with patch('email_verification.DEBUG', True):
            with patch('builtins.print') as mock_print:
                email_verification.dbg('test')
                mock_print.assert_called()

if __name__ == '__main__':
    unittest.main()
