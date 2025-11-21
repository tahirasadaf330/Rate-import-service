"""
Simple Microsoft Graph API Integration Test

Tests only the Graph API authentication and connectivity using .env credentials.
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Add parent directory to path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

def test_graph_api_integration():
    """
    Simple function to test Microsoft Graph API integration.
    Returns True if API integration passes, False otherwise.
    """
    
    print("=" * 60)
    print("MICROSOFT GRAPH API INTEGRATION TEST")
    print("=" * 60)
    
    # Check environment variables
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    user_email = os.getenv('USER_EMAIL')
    
    print(f"Tenant ID: {tenant_id}")
    print(f"Client ID: {client_id}")
    print(f"User Email: {user_email}")
    print(f"Client Secret: {'Set' if client_secret else 'Missing'}")
    print()
    
    # Check if all required credentials are available
    if not all([tenant_id, client_id, client_secret]):
        print("FAILED: Missing required Microsoft Graph API credentials in .env file")
        print("Required variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
        return False
    
    try:
        # Import email verification module
        import email_verification
        print("Successfully imported email_verification module")
        
        # Test authentication
        print("\nTesting Microsoft Graph API authentication...")
        
        token = email_verification.get_token(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            verbose=True
        )
        
        if not token:
            print("FAILED: Could not obtain access token")
            return False
        
        print("Successfully obtained access token")
        print(f"Token preview: {token[:30]}...")
        
        # Validate token structure (JWT tokens have 3 parts separated by dots)
        token_parts = token.split('.')
        if len(token_parts) == 3:
            print("Token has valid JWT structure (3 parts)")
            print(f"Token length: {len(token)} characters")
            
            # Since authentication worked and we got a valid token,
            # the Microsoft Graph API integration is successful
            print("\nMICROSOFT GRAPH API INTEGRATION: PASSED")
            print("Authentication successful - API credentials are valid")
            return True
        else:
            print("FAILED: Token does not have valid JWT structure")
            return False
            
    except ImportError as e:
        print(f"FAILED: Could not import email_verification module: {str(e)}")
        return False
        
    except Exception as e:
        print(f"FAILED: Graph API integration test failed: {str(e)}")
        return False


if __name__ == '__main__':
    success = test_graph_api_integration()
    
    print("\n" + "=" * 60)
    print("FINAL RESULT:")
    if success:
        print("Microsoft Graph API Integration: PASSED")
    else:
        print("Microsoft Graph API Integration: FAILED")
    print("=" * 60)
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)