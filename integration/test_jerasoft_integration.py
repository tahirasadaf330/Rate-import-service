"""
Integration Tests for JeraSoft API

Tests real JeraSoft API integration using credentials from .env file.
"""

import unittest
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import pandas as pd
import time

# Load environment variables
load_dotenv()

# Add parent directory to path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

# Test configuration
TEST_TABLE_ID = 4330
SAFE_TEST_CODES = ['999001', '999002', '999003']


class TestJeraSoftAPIIntegration(unittest.TestCase):
    """Test JeraSoft API integration with real credentials."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Check if JeraSoft API credentials are available
        self.api_key = os.getenv('JERA_SOFT_API_KEY')
        self.api_url = os.getenv('JERASOFT_API_URL')
        
        if not all([self.api_key, self.api_url]):
            self.skipTest("JeraSoft API credentials not found in environment")
    
    def test_jerasoft_api_data_fetch(self):
        """Test fetching data from JeraSoft API."""
        try:
            import jerasoft
            
            # Test fetching tables
            start_time = time.time()
            tables = jerasoft.fetch_all_tables()
            fetch_time = time.time() - start_time
            
            self.assertIsNotNone(tables, "Tables should not be None")
            self.assertIsInstance(tables, list, "Tables should be a list")
            self.assertGreater(len(tables), 0, "Should fetch at least one table")
            
            print(f"JeraSoft API: Fetched {len(tables)} tables in {fetch_time:.2f}s")
            
            # Test table structure
            if len(tables) > 0:
                sample_table = tables[0]
                self.assertIsInstance(sample_table, dict, "Table should be a dictionary")
                self.assertIn('id', sample_table, "Table should have 'id' field")
                
        except Exception as e:
            self.fail(f"JeraSoft API data fetch failed: {str(e)}")
    
    def test_jerasoft_real_upload(self):
        """Test JeraSoft upload functionality with REAL upload."""
        try:
            import rate_upload_to_Jera
            
            # Create test data with safe test codes for table 4330
            test_data = pd.DataFrame([
                {'Code': SAFE_TEST_CODES[0], 'New Rate': 0.0001, 'Effective Date': '2025-12-31'},
                {'Code': SAFE_TEST_CODES[1], 'New Rate': 0.0002, 'Effective Date': '2025-12-31'}
            ])
            
            # Test REAL upload to table 4330
            result = rate_upload_to_Jera.upload_rates_to_jerasoft(
                test_data,
                table_id=TEST_TABLE_ID,
                dry_run=False  # REAL upload
            )
            
            self.assertIsNotNone(result, "Upload result should not be None")
            
            if isinstance(result, dict):
                # Check for successful upload indicators
                if 'status' in result:
                    self.assertEqual(result.get('status'), 'success', "Should confirm successful upload")
                if 'rows_uploaded' in result:
                    self.assertGreater(result.get('rows_uploaded'), 0, "Should upload at least one row")
                if 'files_id' in result:
                    self.assertIsNotNone(result.get('files_id'), "Should have file ID")
                    
            print(f"JeraSoft REAL upload successful: {result}")
            
        except Exception as e:
            self.fail(f"JeraSoft REAL upload failed: {str(e)}")


if __name__ == '__main__':
    unittest.main()