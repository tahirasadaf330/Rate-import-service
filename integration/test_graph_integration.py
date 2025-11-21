"""
Integration Tests for Microsoft Graph API

Tests real Microsoft Graph API integration using credentials from .env file.
"""

import unittest
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Add parent directory to path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))


class TestMicrosoftGraphAPIIntegration(unittest.TestCase):
    """Test Microsoft Graph API integration with real credentials."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Check if Microsoft Graph API credentials are available
        self.tenant_id = os.getenv('TENANT_ID')
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        
        if not all([self.tenant_id, self.client_id, self.client_secret]):
            self.skipTest("Microsoft Graph API credentials not found in environment")
    
    def test_graph_api_authentication(self):
        """Test Microsoft Graph API authentication."""
        try:
            import email_verification
            
            # Test authentication
            token = email_verification.get_token(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret,
                verbose=False
            )
            
            self.assertIsNotNone(token, "Access token should not be None")
            self.assertIsInstance(token, str, "Access token should be a string")
            self.assertGreater(len(token), 0, "Access token should not be empty")
            
            # Validate JWT structure
            token_parts = token.split('.')
            self.assertEqual(len(token_parts), 3, "JWT token should have 3 parts")
            
            print(f"Graph API authentication successful: Token length {len(token)} chars")
            
        except Exception as e:
            self.fail(f"Microsoft Graph API authentication failed: {str(e)}")
    
    def test_graph_api_token_validity(self):
        """Test that obtained token has valid structure."""
        try:
            import email_verification
            
            token = email_verification.get_token(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret,
                verbose=False
            )
            
            # Test token characteristics
            self.assertGreater(len(token), 1000, "Token should be reasonably long")
            self.assertLess(len(token), 10000, "Token should not be excessively long")
            self.assertTrue(token.startswith('eyJ'), "JWT token should start with 'eyJ'")
            
            print("Graph API token structure validation passed")
            
        except Exception as e:
            self.fail(f"Graph API token validation failed: {str(e)}")


if __name__ == '__main__':
    unittest.main()