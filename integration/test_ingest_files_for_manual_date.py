"""
Integration test for ingest_files_for_manual_date function
Uses real file system and database credentials from .env
"""

import unittest
import os
import sys
from pathlib import Path
import json
from dotenv import load_dotenv
import pandas as pd

# Load environment variables
load_dotenv()

# Add parent directory to path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))


class TestIngestFilesForManualDateIntegration(unittest.TestCase):
    def setUp(self):
        self.attachments_root = Path(parent_dir) / 'attachments'
        self.test_dirs = []
        # 1. Valid ingestion
        valid_dir = self.attachments_root / 'integration_valid_dir'
        valid_dir.mkdir(parents=True, exist_ok=True)
        file_name = 'valid_file.xlsx'
        file_path = valid_dir / file_name
        df = pd.DataFrame({'A': [pd.Timestamp('2025-11-20')]})
        df.to_excel(file_path, index=False, header=False)
        meta = {
            'attachments': [file_name],
            'directory': str(valid_dir.relative_to(parent_dir)),
            'sender': 'integration@example.com',
            'subject': 'integration test',
            'receivedDateTime_raw': '2025-11-20T12:00:00Z',
            'processed_at_utc': '2025-11-20T13:00:00Z',
        }
        with (valid_dir / 'metadata.json').open('w', encoding='utf-8') as f:
            json.dump(meta, f)
        self.test_dirs.append(valid_dir)

        # 2. Missing metadata.json
        missing_meta_dir = self.attachments_root / 'integration_missing_meta_dir'
        missing_meta_dir.mkdir(parents=True, exist_ok=True)
        file_path = missing_meta_dir / 'missing_meta_file.xlsx'
        df.to_excel(file_path, index=False, header=False)
        self.test_dirs.append(missing_meta_dir)

        # 3. Malformed metadata.json
        malformed_dir = self.attachments_root / 'integration_malformed_dir'
        malformed_dir.mkdir(parents=True, exist_ok=True)
        file_path = malformed_dir / 'malformed_file.xlsx'
        df.to_excel(file_path, index=False, header=False)
        with (malformed_dir / 'metadata.json').open('w', encoding='utf-8') as f:
            f.write('{bad json')
        self.test_dirs.append(malformed_dir)

        # 4. Missing attachment file
        missing_file_dir = self.attachments_root / 'integration_missing_file_dir'
        missing_file_dir.mkdir(parents=True, exist_ok=True)
        meta = {
            'attachments': ['missing_file.xlsx'],
            'directory': str(missing_file_dir.relative_to(parent_dir)),
            'sender': 'integration@example.com',
            'subject': 'integration test',
            'receivedDateTime_raw': '2025-11-20T12:00:00Z',
            'processed_at_utc': '2025-11-20T13:00:00Z',
        }
        with (missing_file_dir / 'metadata.json').open('w', encoding='utf-8') as f:
            json.dump(meta, f)
        self.test_dirs.append(missing_file_dir)

        # 5. Unsupported file type
        unsupported_dir = self.attachments_root / 'integration_unsupported_dir'
        unsupported_dir.mkdir(parents=True, exist_ok=True)
        file_path = unsupported_dir / 'unsupported_file.txt'
        with file_path.open('w') as f:
            f.write('not an excel or csv')
        meta = {
            'attachments': ['unsupported_file.txt'],
            'directory': str(unsupported_dir.relative_to(parent_dir)),
            'sender': 'integration@example.com',
            'subject': 'integration test',
            'receivedDateTime_raw': '2025-11-20T12:00:00Z',
            'processed_at_utc': '2025-11-20T13:00:00Z',
        }
        with (unsupported_dir / 'metadata.json').open('w', encoding='utf-8') as f:
            json.dump(meta, f)
        self.test_dirs.append(unsupported_dir)

        # 6. Already ingested directory
        already_ingested_dir = self.attachments_root / 'integration_already_ingested_dir'
        already_ingested_dir.mkdir(parents=True, exist_ok=True)
        file_name = 'already_ingested.xlsx'
        file_path = already_ingested_dir / file_name
        df.to_excel(file_path, index=False, header=False)
        meta = {
            'attachments': [file_name],
            'directory': str(already_ingested_dir.relative_to(parent_dir)),
            'sender': 'integration@example.com',
            'subject': 'integration test',
            'receivedDateTime_raw': '2025-11-20T12:00:00Z',
            'processed_at_utc': '2025-11-20T13:00:00Z',
            'date_verification_ingestion': True
        }
        with (already_ingested_dir / 'metadata.json').open('w', encoding='utf-8') as f:
            json.dump(meta, f)
        self.test_dirs.append(already_ingested_dir)

    def tearDown(self):
        # Cleanup: remove all test directories and files
        for test_dir in self.test_dirs:
            for item in test_dir.iterdir():
                item.unlink()
            test_dir.rmdir()

    def test_ingest_files_for_manual_date_all_cases(self):
        import date_verification
        result = date_verification.ingest_files_for_manual_date(self.attachments_root)
        self.assertIsInstance(result, tuple)
        scanned, inserted, skipped = result
        # Only valid Excel file should be inserted
        self.assertEqual(inserted, 3)
        # Malformed, missing meta, already ingested should be skipped
        self.assertEqual(skipped, 3)
        # Check metadata.json updated for valid and already ingested
        valid_meta_path = self.test_dirs[0] / 'metadata.json'
        with valid_meta_path.open('r', encoding='utf-8') as f:
            meta = json.load(f)
        self.assertTrue(meta.get('date_verification_ingestion'))
        already_meta_path = self.test_dirs[5] / 'metadata.json'
        with already_meta_path.open('r', encoding='utf-8') as f:
            meta = json.load(f)
        self.assertTrue(meta.get('date_verification_ingestion'))

if __name__ == '__main__':
    unittest.main()
