"""
Unit tests for jerasoft.py module.

This module tests all the functions in jerasoft.py including:
- String utility functions
- Prefix handling functions
- API interaction functions
- Data export functions
- Error handling scenarios
"""

import unittest
from unittest.mock import Mock, patch, MagicMock, mock_open
import pandas as pd
from datetime import datetime
import json
import os
from pathlib import Path
import tempfile
import requests
from requests.exceptions import HTTPError, ConnectionError, Timeout

# Import the module under test
import jerasoft


class TestStringUtilities(unittest.TestCase):
    """Test string utility functions."""

    def test_normalize(self):
        """Test the normalize function."""
        test_cases = [
            ("Hello World", "hello world"),
            ("TEST.WITH.DOTS", "test with dots"),
            ("Multiple   Spaces", "multiple spaces"),
            ("With###Hashes", "withhashes"),
            ("", ""),
            (None, ""),
            ("  Mixed  .  CASE  ###  ", "mixed case"),
        ]
        
        for input_str, expected in test_cases:
            with self.subTest(input_str=input_str):
                result = jerasoft.normalize(input_str)
                self.assertEqual(result, expected)

    def test_fuzzy_score(self):
        """Test the fuzzy_score function."""
        # Identical strings should score 1.0
        self.assertEqual(jerasoft.fuzzy_score("test", "test"), 1.0)
        
        # Empty strings - fuzzy_score returns 0.85 for empty strings due to the weighted calculation
        score = jerasoft.fuzzy_score("", "")
        self.assertGreater(score, 0.8)
        
        # Similar strings should have high scores
        score = jerasoft.fuzzy_score("quickcom tel", "QUICKCOM TEL")
        self.assertGreater(score, 0.8)
        
        # Completely different strings should have low scores
        score = jerasoft.fuzzy_score("abc", "xyz")
        self.assertLess(score, 0.5)
        
        # Test with actual company names
        score = jerasoft.fuzzy_score(
            "TERM Quickcom tel PRM trunk Prefix:1001 USD",
            "quickcom tel prefix 1001"
        )
        self.assertGreater(score, 0.6)

    def test_name_starts_with_term(self):
        """Test the name_starts_with_term function."""
        test_cases = [
            ("TERM Quickcom", True),
            ("term something", True),
            ("  TERM with spaces", True),
            ("Not a term table", False),
            ("TERMINAL", True),  # This actually matches because it starts with "TERM"
            ("", False),
        ]
        
        for name, expected in test_cases:
            with self.subTest(name=name):
                result = jerasoft.name_starts_with_term(name)
                self.assertEqual(result, expected)

    def test_name_contains_company(self):
        """Test the name_contains_company function."""
        test_cases = [
            ("TERM Quickcom tel", "quickcom", True),
            ("Something QUICKCOM else", "quickcom", True),
            ("No match here", "quickcom", False),
            ("", "test", False),
            ("test company", "", True),  # Empty string is always found in any string
        ]
        
        for name, company, expected in test_cases:
            with self.subTest(name=name, company=company):
                result = jerasoft.name_contains_company(name, company)
                self.assertEqual(result, expected)

    def test_extract_company_keyword(self):
        """Test the extract_company_keyword function."""
        test_cases = [
            ("Quickcom tel PRM", "quickcom"),
            ("Test-Company.Name", "test-company.name"),
            ("123Company", "123company"),
            ("", ""),
            ("   SpacedName   ", "spacedname"),
            ("Special@Chars", "special"),
        ]
        
        for query, expected in test_cases:
            with self.subTest(query=query):
                result = jerasoft.extract_company_keyword(query)
                self.assertEqual(result, expected)


class TestPrefixUtilities(unittest.TestCase):
    """Test prefix handling functions."""

    def test_normalize_prefix(self):
        """Test the normalize_prefix function."""
        test_cases = [
            ("123", "123"),
            (123, "123"),
            ("0123", "123"),  # Leading zeros removed
            ("  456  ", "456"),
            ("", None),
            (None, None),
            ("abc", None),
            ("12.34", None),
        ]
        
        for input_val, expected in test_cases:
            with self.subTest(input_val=input_val):
                result = jerasoft.normalize_prefix(input_val)
                self.assertEqual(result, expected)

    def test_table_prefix_from_name(self):
        """Test the table_prefix_from_name function."""
        test_cases = [
            ("TERM Quickcom PREFIX:1001", "1001"),
            ("Something PRFX-33 test", "33"),
            ("prefix: 456", "456"),
            ("No prefix here", None),
            ("", None),
            ("PREFIX:", None),
        ]
        
        for name, expected in test_cases:
            with self.subTest(name=name):
                result = jerasoft.table_prefix_from_name(name)
                self.assertEqual(result, expected)

    def test_table_has_prefix(self):
        """Test the table_has_prefix function."""
        test_cases = [
            ("TERM Quickcom PREFIX:1001", "1001", True),
            ("TERM Quickcom PREFIX:1001", "001", False),  # Different prefix
            ("TERM Quickcom PREFIX:1001", "1001", True),
            ("No prefix table", "123", False),
            ("", "123", False),
        ]
        
        for name, prefix_code, expected in test_cases:
            with self.subTest(name=name, prefix_code=prefix_code):
                result = jerasoft.table_has_prefix(name, prefix_code)
                self.assertEqual(result, expected)


class TestAPIFunctions(unittest.TestCase):
    """Test API interaction functions."""

    def setUp(self):
        """Set up test fixtures."""
        self.mock_api_url = "http://test-api.com"
        self.mock_api_key = "test-key-123"
        
        self.sample_tables = [
            {
                "id": 1,
                "name": "TERM Quickcom tel PRM trunk Prefix:1001 USD",
                "description": "Test table 1"
            },
            {
                "id": 2,
                "name": "TERM AnotherCo PREFIX:2002",
                "description": "Test table 2"
            },
            {
                "id": 3,
                "name": "NON-TERM Regular table",
                "description": "Not a TERM table"
            }
        ]

    @patch('jerasoft._post_json')
    def test_fetch_all_tables_success(self, mock_post_json):
        """Test successful fetch_all_tables."""
        # Mock pagination - first page has data, second is empty
        mock_post_json.side_effect = [
            {"result": self.sample_tables[:2]},
            {"result": []}  # Empty result indicates end of pagination
        ]
        
        result = jerasoft.fetch_all_tables(
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["id"], 1)
        self.assertEqual(mock_post_json.call_count, 2)

    @patch('jerasoft._post_json')
    def test_fetch_all_tables_api_error(self, mock_post_json):
        """Test fetch_all_tables with API error."""
        mock_post_json.return_value = {"error": "API Error"}
        
        with self.assertRaises(RuntimeError) as cm:
            jerasoft.fetch_all_tables(
                api_url=self.mock_api_url,
                api_key=self.mock_api_key
            )
        
        self.assertIn("API error", str(cm.exception))

    @patch('jerasoft.DEFAULT_API_KEY', None)  # Mock DEFAULT_API_KEY to be None
    def test_fetch_all_tables_no_api_key(self):
        """Test fetch_all_tables without API key."""
        with self.assertRaises(ValueError) as cm:
            jerasoft.fetch_all_tables(api_url=self.mock_api_url, api_key=None)
        
        self.assertIn("Missing API key", str(cm.exception))

    @patch('jerasoft.fetch_all_tables')
    def test_find_best_term_table_success(self, mock_fetch):
        """Test successful find_best_term_table."""
        mock_fetch.return_value = self.sample_tables
        
        table_id, best_table, scored = jerasoft.find_best_term_table(
            target_query="Quickcom tel PRM trunk",
            subject="Quickcom tel PRM trunk Prefix:1001 USD",
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertEqual(table_id, 1)
        self.assertEqual(best_table["name"], "TERM Quickcom tel PRM trunk Prefix:1001 USD")
        self.assertIsInstance(scored, list)
        self.assertGreater(len(scored), 0)

    @patch('jerasoft.fetch_all_tables')
    def test_find_best_term_table_with_prefix_filter(self, mock_fetch):
        """Test find_best_term_table with prefix filtering."""
        mock_fetch.return_value = self.sample_tables
        
        table_id, best_table, scored = jerasoft.find_best_term_table(
            target_query="Quickcom tel PRM trunk",
            subject="Quickcom tel PRM trunk Prefix:1001 USD",
            prefix_code="1001",
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertEqual(table_id, 1)
        self.assertIn("1001", best_table["name"])

    @patch('jerasoft.fetch_all_tables')
    def test_find_best_term_table_no_matches(self, mock_fetch):
        """Test find_best_term_table with no matching tables."""
        mock_fetch.return_value = self.sample_tables
        
        result = jerasoft.find_best_term_table(
            target_query="NonExistentCompany",
            subject="NonExistentCompany test",
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        # Should return tuple with error message as first element
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)
        self.assertIn("No TERM* tables found", result[0])

    def test_find_best_term_table_empty_query(self):
        """Test find_best_term_table with empty query."""
        with self.assertRaises(ValueError) as cm:
            jerasoft.find_best_term_table("", "test")
        
        self.assertIn("target_query must be non-empty", str(cm.exception))

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_exact_match(self, mock_fetch):
        """Test get_table_id_by_name with exact match."""
        mock_fetch.return_value = self.sample_tables
        
        result = jerasoft.get_table_id_by_name(
            "TERM Quickcom tel PRM trunk Prefix:1001 USD",
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertEqual(result, 1)

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_fuzzy_match(self, mock_fetch):
        """Test get_table_id_by_name with fuzzy match."""
        mock_fetch.return_value = self.sample_tables
        
        # Test with slightly different name that should fuzzy match
        with patch('builtins.print'):  # Suppress print output
            result = jerasoft.get_table_id_by_name(
                "TERM Quickcom tel PRM trunk Prefix 1001 USD",  # Missing colon
                api_url=self.mock_api_url,
                api_key=self.mock_api_key
            )
        
        self.assertEqual(result, 1)

    @patch('jerasoft.fetch_all_tables')
    def test_get_table_id_by_name_no_match(self, mock_fetch):
        """Test get_table_id_by_name with no match."""
        mock_fetch.return_value = self.sample_tables
        
        with patch('builtins.print'):  # Suppress print output
            result = jerasoft.get_table_id_by_name(
                "Completely Different Table Name",
                api_url=self.mock_api_url,
                api_key=self.mock_api_key
            )
        
        self.assertIsNone(result)

    @patch('jerasoft._post_json')
    def test_fetch_active_current_future_rates_success(self, mock_post_json):
        """Test successful fetch_active_current_future_rates."""
        mock_rates_data = [
            {
                "code": "1001",
                "code_name": "Test Destination",
                "value": "0.050",
                "effective_from": "2024-01-01 00:00:00",
                "min_volume": "1",
                "pay_interval": "60"
            }
        ]
        
        mock_post_json.side_effect = [
            {"result": mock_rates_data},
            {"result": []}  # End pagination
        ]
        
        result = jerasoft.fetch_active_current_future_rates(
            table_id=1,
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertIsInstance(result, pd.DataFrame)
        self.assertGreater(len(result), 0)
        self.assertIn("Dst Code", result.columns)
        self.assertIn("Rate", result.columns)

    @patch('jerasoft._post_json')
    def test_fetch_active_current_future_rates_json_decode_error(self, mock_post_json):
        """Test fetch_active_current_future_rates with JSON decode error and retry logic."""
        # First call fails with JSON decode error, second succeeds with smaller page size
        mock_post_json.side_effect = [
            RuntimeError("JSON decode failed at pos 1000: Expecting ',' delimiter"),
            {"result": []},  # Success on retry with smaller page
        ]
        
        result = jerasoft.fetch_active_current_future_rates(
            table_id=1,
            api_url=self.mock_api_url,
            api_key=self.mock_api_key,
            page_limit=500
        )
        
        self.assertIsInstance(result, pd.DataFrame)
        # Should have called twice due to retry logic
        self.assertEqual(mock_post_json.call_count, 2)

    @patch('jerasoft._post_json')
    def test_fetch_active_current_future_rates_empty_result(self, mock_post_json):
        """Test fetch_active_current_future_rates with empty result."""
        mock_post_json.return_value = {"result": []}
        
        result = jerasoft.fetch_active_current_future_rates(
            table_id=1,
            api_url=self.mock_api_url,
            api_key=self.mock_api_key
        )
        
        self.assertIsInstance(result, pd.DataFrame)
        self.assertEqual(len(result), 0)
        # Should have expected columns even when empty
        expected_columns = ["Dst Code", "Dst Code Name", "Rate", "Effective Date", "Billing Increment"]
        for col in expected_columns:
            self.assertIn(col, result.columns)


class TestDataExport(unittest.TestCase):
    """Test data export functions."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_df = pd.DataFrame({
            "Dst Code": ["1001", "1002"],
            "Dst Code Name": ["Test Dest 1", "Test Dest 2"],
            "Rate": ["0.050", "0.060"],
            "Effective Date": ["2024-01-01", "2024-01-02"],
            "Billing Increment": ["1/60", "1/60"]
        })

    def test_save_rates_to_excel_success(self):
        """Test successful save_rates_to_excel."""
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, "test_rates.xlsx")
            
            result_path = jerasoft.save_rates_to_excel(self.test_df, output_path)
            
            # Check file was created
            self.assertTrue(os.path.exists(result_path))
            
            # Check we can read it back
            read_df = pd.read_excel(result_path)
            self.assertEqual(len(read_df), 2)
            self.assertEqual(list(read_df.columns), list(self.test_df.columns))

    def test_save_rates_to_excel_creates_directory(self):
        """Test save_rates_to_excel creates parent directories."""
        with tempfile.TemporaryDirectory() as temp_dir:
            nested_path = os.path.join(temp_dir, "nested", "folder", "test_rates.xlsx")
            
            result_path = jerasoft.save_rates_to_excel(self.test_df, nested_path)
            
            # Check file and directories were created
            self.assertTrue(os.path.exists(result_path))
            self.assertTrue(os.path.isdir(os.path.dirname(result_path)))

    def test_save_rates_to_excel_empty_path(self):
        """Test save_rates_to_excel with empty output path."""
        with self.assertRaises(ValueError) as cm:
            jerasoft.save_rates_to_excel(self.test_df, "")
        
        self.assertIn("output_path must be provided", str(cm.exception))

    @patch('jerasoft.get_table_id_by_name')
    @patch('jerasoft.fetch_active_current_future_rates')
    @patch('jerasoft.save_rates_to_excel')
    def test_export_rates_by_query_with_forced_table(self, mock_save, mock_fetch, mock_get_id):
        """Test export_rates_by_query with forced table name."""
        mock_get_id.return_value = 123
        mock_fetch.return_value = self.test_df
        mock_save.return_value = "/path/to/output.xlsx"
        
        result = jerasoft.export_rates_by_query(
            target_query="test query",
            output_path="output.xlsx",
            subject="test subject",
            force_table_name="TERM Test Table"
        )
        
        self.assertIsInstance(result, dict)
        self.assertEqual(result["table_id"], 123)
        self.assertEqual(result["rows"], 2)
        self.assertIn("saved_to", result)

    @patch('jerasoft.find_best_term_table')
    @patch('jerasoft.fetch_active_current_future_rates')
    @patch('jerasoft.save_rates_to_excel')
    def test_export_rates_by_query_normal_flow(self, mock_save, mock_fetch, mock_find):
        """Test export_rates_by_query normal flow."""
        mock_find.return_value = (
            123,
            {"id": 123, "name": "Test Table"},
            [(0.95, {"id": 123, "name": "Test Table"})]
        )
        mock_fetch.return_value = self.test_df
        mock_save.return_value = "/path/to/output.xlsx"
        
        with patch('builtins.print'):  # Suppress debug prints
            result = jerasoft.export_rates_by_query(
                target_query="Quickcom test",
                output_path="output.xlsx",
                subject="Quickcom test subject"
            )
        
        self.assertIsInstance(result, dict)
        self.assertEqual(result["table_id"], 123)
        self.assertEqual(result["rows"], 2)
        self.assertIn("top_candidates", result)

    @patch('jerasoft.get_table_id_by_name')
    def test_export_rates_by_query_forced_table_not_found(self, mock_get_id):
        """Test export_rates_by_query with forced table that doesn't exist."""
        mock_get_id.return_value = None
        
        result = jerasoft.export_rates_by_query(
            target_query="test query",
            output_path="output.xlsx",
            subject="test subject",
            force_table_name="Non-existent Table"
        )
        
        self.assertIsInstance(result, str)
        self.assertIn("not found", result)

    @patch('jerasoft.find_best_term_table')
    def test_export_rates_by_query_no_matching_tables(self, mock_find):
        """Test export_rates_by_query when no tables match."""
        mock_find.return_value = ("No TERM* tables found containing company 'test'.", "", "")
        
        result = jerasoft.export_rates_by_query(
            target_query="NonExistentCompany",
            output_path="output.xlsx",
            subject="test subject"
        )
        
        self.assertIsInstance(result, str)
        self.assertIn("No TERM* tables found", result)


class TestErrorHandling(unittest.TestCase):
    """Test error handling scenarios."""

    @patch('jerasoft._session.post')
    def test_post_json_http_error(self, mock_post):
        """Test _post_json with HTTP error."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = HTTPError("HTTP 500 Error")
        mock_post.return_value = mock_response
        
        with self.assertRaises(HTTPError):
            jerasoft._post_json("http://test.com", {"test": "data"})

    @patch('jerasoft._session.post')
    def test_post_json_invalid_json(self, mock_post):
        """Test _post_json with invalid JSON response."""
        mock_response = Mock()
        mock_response.raise_for_status.return_value = None
        mock_response.json.side_effect = json.JSONDecodeError("Invalid JSON", "doc", 0)
        mock_response.text = "Invalid JSON response"
        mock_response.status_code = 200
        mock_response.headers = {"Content-Type": "application/json"}
        mock_post.return_value = mock_response
        
        with self.assertRaises(RuntimeError) as cm:
            jerasoft._post_json("http://test.com", {"test": "data"})
        
        self.assertIn("JSON decode failed", str(cm.exception))

    @patch('jerasoft._post_json')
    def test_fetch_active_rates_unexpected_response(self, mock_post_json):
        """Test fetch_active_current_future_rates with unexpected response format."""
        mock_post_json.return_value = {"result": "not a list"}
        
        with self.assertRaises(RuntimeError) as cm:
            jerasoft.fetch_active_current_future_rates(
                table_id=1,
                api_url="http://test.com",
                api_key="test-key"
            )
        
        self.assertIn("Unexpected response", str(cm.exception))


class TestEnvironmentDefaults(unittest.TestCase):
    """Test environment variable defaults and configuration."""

    def test_default_values(self):
        """Test default API URL and key handling."""
        # Test that defaults are set correctly
        self.assertIsNotNone(jerasoft.DEFAULT_API_URL)
        self.assertIn("billing.voipsystem.org", jerasoft.DEFAULT_API_URL)
        
        # Test headers
        self.assertIn("Content-Type", jerasoft.DEFAULT_HEADERS)
        self.assertEqual(jerasoft.DEFAULT_HEADERS["Content-Type"], "application/json")

    @patch.dict(os.environ, {"JERASOFT_API_URL": "http://custom-url.com", "JERA_SOFT_API_KEY": "custom-key"})
    def test_environment_override(self):
        """Test that environment variables override defaults."""
        # Reload the module to pick up new env vars
        import importlib
        importlib.reload(jerasoft)
        
        self.assertEqual(jerasoft.DEFAULT_API_URL, "http://custom-url.com")
        self.assertEqual(jerasoft.DEFAULT_API_KEY, "custom-key")


if __name__ == "__main__":
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test classes
    test_classes = [
        TestStringUtilities,
        TestPrefixUtilities,
        TestAPIFunctions,
        TestDataExport,
        TestErrorHandling,
        TestEnvironmentDefaults,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print(f"\n{'='*50}")
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    print(f"Success rate: {((result.testsRun - len(result.failures) - len(result.errors)) / result.testsRun * 100):.1f}%")