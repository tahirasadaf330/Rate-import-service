import unittest
from unittest.mock import patch, MagicMock
import email_verification
import json
from pathlib import Path
from datetime import datetime, timezone
class TestEmailVerification(unittest.TestCase):
    def test_bracketed_subject_empty_trunk(self):
        subj = '[IPBTEL][ ][44090][USD]'
        result = email_verification.validate_subject(subj)
        self.assertIsInstance(result, dict)
        self.assertEqual(result['company'], 'ipbtel')
        self.assertEqual(result['trunk'], '')
        self.assertEqual(result['prefix'], 44090)
        self.assertEqual(result['currency'], 'usd')
    def test_bracketed_subject_valid(self):
            subj = '[ALLIP][PREMIUM][1062#][USD]'
            self.assertIsInstance(email_verification.validate_subject(subj), dict)
            self.assertTrue(email_verification.subject_ok(subj))
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
        self.assertEqual(out['company'], 'comp')
        self.assertEqual(out['trunk'], 'trunk')
        self.assertEqual(out['prefix'], 123)
        self.assertEqual(out['currency'], 'usd')
        self.assertIsNone(email_verification._normalize_output('', '', '', ''))

    def test_extract_anyorder(self):
        subj = '[Acme][Trunk][123][USD]'
        result = email_verification.validate_subject(subj)
        self.assertIsInstance(result, dict)
        self.assertEqual(result['currency'], 'usd')
        self.assertIsNone(email_verification.validate_subject('no currency here'))

    def test_validate_subject(self):
        valid = email_verification.validate_subject('[Acme][Trunk][123][USD]')
        self.assertIsInstance(valid, dict)
        self.assertIsNone(email_verification.validate_subject('bad subject'))

    def test_subject_ok(self):
        self.assertTrue(email_verification.subject_ok('[Acme][Trunk][123][USD]'))
        self.assertFalse(email_verification.subject_ok('bad subject'))

    def test_dbg(self):
        with patch('email_verification.DEBUG', True):
            with patch('builtins.print') as mock_print:
                email_verification.dbg('test')
                mock_print.assert_called()

if __name__ == '__main__':
    unittest.main()
