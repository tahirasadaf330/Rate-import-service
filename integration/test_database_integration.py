"""
Integration Tests for Database Connectivity

Tests real database integration using credentials from .env file.
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


class TestDatabaseIntegration(unittest.TestCase):
    """Test database integration with real credentials."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Check if database credentials are available (using correct .env variable names)
        self.db_host = os.getenv('DB_HOST')
        self.db_port = os.getenv('DB_PORT')
        self.db_name = os.getenv('DB_DATABASE')
        self.db_user = os.getenv('DB_USERNAME')
        self.db_password = os.getenv('DB_PASSWORD')
        
        if not all([self.db_host, self.db_port, self.db_name, self.db_user, self.db_password]):
            self.skipTest("Database credentials not found in environment")
    
    def test_database_connection(self):
        """Test real database connection."""
        try:
            import database
            
            # Test connection
            conn = database.get_conn()
            self.assertIsNotNone(conn, "Database connection should not be None")
            
            # Test basic query
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            result = cursor.fetchone()
            
            self.assertEqual(result[0], 1, "Basic query should return 1")
            
            cursor.close()
            conn.close()
            
            print(f"Database integration successful: {self.db_host}:{self.db_port}/{self.db_name}")
            
        except Exception as e:
            self.fail(f"Database integration failed: {str(e)}")
    
    def test_database_schema_access(self):
        """Test access to database schema."""
        try:
            import database
            
            conn = database.get_conn()
            cursor = conn.cursor()
            
            # Test schema access
            cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' LIMIT 5")
            tables = cursor.fetchall()
            
            self.assertGreater(len(tables), 0, "Should find at least one table in public schema")
            
            cursor.close()
            conn.close()
            
            print(f"Found {len(tables)} tables in database schema")
            
        except Exception as e:
            self.fail(f"Database schema access failed: {str(e)}")


if __name__ == '__main__':
    unittest.main()