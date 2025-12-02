import unittest
from unittest.mock import patch, MagicMock
import email_verification
import json
from pathlib import Path
from datetime import datetime, timezone

class TestEmailVerification(unittest.TestCase):
    def test_get_token_30s_timeout(self):
        import time
        from unittest.mock import patch, MagicMock
        with patch('email_verification.ConfidentialClientApplication') as mock_app, \
             patch('logging.getLogger') as mock_logger:
            instance = mock_app.return_value
            # Simulate a hanging token acquisition
            def slow_acquire(*args, **kwargs):
                time.sleep(31)
            instance.acquire_token_for_client.side_effect = slow_acquire
            logger = MagicMock()
            mock_logger.return_value = logger
            with self.assertRaises(Exception) as cm:
                email_verification.get_token("tenant", "client", "secret")
            self.assertIn("Token acquisition timed out (30 seconds)", str(cm.exception))
            logger.error.assert_any_call("Token acquisition timed out (30 seconds): ")

    def test_get_token_timeout_logging(self):
        with patch('email_verification.ConfidentialClientApplication') as mock_app, \
             patch('logging.getLogger') as mock_logger:
            instance = mock_app.return_value
            logger = MagicMock()
            mock_logger.return_value = logger
            # Simulate 'timeout' error
            instance.acquire_token_for_client.side_effect = Exception("timeout occurred")
            with self.assertRaises(Exception):
                email_verification.get_token("tenant", "client", "secret")
            logger.error.assert_any_call("Token acquisition timed out: timeout occurred")

            # Simulate 'timed out' error
            instance.acquire_token_for_client.side_effect = Exception("Request timed out")
            with self.assertRaises(Exception):
                email_verification.get_token("tenant", "client", "secret")
            logger.error.assert_any_call("Token acquisition timed out: Request timed out")
            logger.error.assert_any_call("Token acquisition timed out: Request timed out")

            # Simulate other error
            instance.acquire_token_for_client.side_effect = Exception("other error")
            with self.assertRaises(Exception):
                email_verification.get_token("tenant", "client", "secret")
            logger.error.assert_any_call("Token acquisition failed: other error")
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
        self.assertEqual(out['prefix'], 123)
        self.assertEqual(out['currency'], 'USD')
        self.assertIsNone(email_verification._normalize_output('', '', '', ''))

    def test_extract_anyorder(self):
        subj = 'Acme trunk prefix 123 USD'
        result = email_verification._extract_anyorder(subj)
        self.assertIsInstance(result, dict)
        self.assertEqual(result['currency'], 'USD')
        self.assertIsNone(email_verification._extract_anyorder('no currency here'))

    def test_validate_subject(self):
        valid = email_verification.validate_subject('Acme trunk prefix 123 USD')
        self.assertIsInstance(valid, dict)
        self.assertIsNone(email_verification.validate_subject('bad subject'))

    def test_subject_ok(self):
        self.assertTrue(email_verification.subject_ok('Acme trunk prefix 123 USD'))
        self.assertFalse(email_verification.subject_ok('bad subject'))

    def test_dbg(self):
        with patch('email_verification.DEBUG', True):
            with patch('builtins.print') as mock_print:
                email_verification.dbg('test')
                mock_print.assert_called()

if __name__ == '__main__':
    unittest.main()
