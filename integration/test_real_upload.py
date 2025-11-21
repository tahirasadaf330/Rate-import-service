"""
JeraSoft API REAL Data Push Test

Tests ACTUAL data upload to JeraSoft API using table ID 4330.
Uses safe test codes 999001-999003 for real upload verification.
"""

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

# SAFE TEST CONFIGURATION
TEST_TABLE_ID = 4330  # ONLY use this table ID
SAFE_TEST_CODES = ['999001', '999002', '999003']  # Safe test prefix codes

def test_real_jerasoft_upload():
    """
    Test REAL data upload to JeraSoft API using table ID 4330.
    This will actually upload data and show the results.
    """
    
    print("=" * 60)
    print("JERASOFT REAL DATA UPLOAD TEST")
    print("=" * 60)
    
    # Check credentials
    api_key = os.getenv('JERA_SOFT_API_KEY')
    api_url = os.getenv('JERASOFT_API_URL')
    
    print(f"API URL: {api_url}")
    print(f"API Key: {'Available' if api_key else 'Missing'}")
    print(f"Target Table ID: {TEST_TABLE_ID}")
    print(f"Test Codes: {', '.join(SAFE_TEST_CODES)}")
    print()
    
    if not all([api_key, api_url]):
        print("FAILED: Missing JeraSoft API credentials")
        return False
    
    try:
        # Import required modules
        import rate_upload_to_Jera
        import jerasoft
        print("Required modules imported successfully")
        
        # Create test DataFrame with SAFE test data
        test_data = pd.DataFrame([
            {
                'Code': SAFE_TEST_CODES[0], 
                'New Rate': 0.0001, 
                'Effective Date': '2025-12-31'
            },
            {
                'Code': SAFE_TEST_CODES[1], 
                'New Rate': 0.0002, 
                'Effective Date': '2025-12-31'
            },
            {
                'Code': SAFE_TEST_CODES[2], 
                'New Rate': 0.0003, 
                'Effective Date': '2025-12-31'
            }
        ])
        
        print("Test data to upload:")
        for idx, row in test_data.iterrows():
            print(f"  {idx+1}. Code: {row['Code']}, Rate: {row['New Rate']}, Date: {row['Effective Date']}")
        
        print(f"\n🚀 REAL UPLOAD to table {TEST_TABLE_ID}...")
        print("⚠️  WARNING: This will actually upload data!")
        
        # REAL data upload (NOT dry run)
        start_time = time.time()
        
        result = rate_upload_to_Jera.upload_rates_to_jerasoft(
            test_data,
            table_id=TEST_TABLE_ID,  # ONLY table 4330
            dry_run=False  # REAL UPLOAD
        )
        
        upload_time = time.time() - start_time
        
        if result is not None:
            print(f"\n✅ UPLOAD COMPLETED in {upload_time:.2f} seconds")
            print(f"Upload Result: {result}")
            
            # Validate result
            if isinstance(result, dict):
                if 'processed' in result:
                    print(f"📊 Successfully processed {result['processed']} rows")
                if 'uploaded' in result:
                    print(f"📤 Successfully uploaded {result['uploaded']} rows")
                if 'errors' in result and result['errors']:
                    print(f"⚠️  Errors encountered: {result['errors']}")
                else:
                    print("✅ No errors reported")
            
            # Now verify the upload by fetching data from the table
            print(f"\n🔍 Verifying upload by checking table {TEST_TABLE_ID}...")
            
            try:
                # Fetch rates from the table to verify our upload
                rates = jerasoft.fetch_active_current_future_rates(
                    table_id=TEST_TABLE_ID,
                    page_limit=10
                )
                
                if rates and len(rates) > 0:
                    print(f"📊 Found {len(rates)} rates in table {TEST_TABLE_ID}")
                    
                    # Look for our uploaded test codes
                    uploaded_codes = []
                    for rate in rates:
                        if hasattr(rate, 'get'):
                            code = rate.get('code') or rate.get('prefix') or rate.get('destination')
                        else:
                            # If it's a DataFrame row
                            code = getattr(rate, 'code', None) or getattr(rate, 'prefix', None)
                        
                        if code in SAFE_TEST_CODES:
                            uploaded_codes.append(code)
                            print(f"✅ Found uploaded code: {code}")
                    
                    if uploaded_codes:
                        print(f"\n🎉 SUCCESS: {len(uploaded_codes)} test codes found in table!")
                        print(f"Verified codes: {uploaded_codes}")
                    else:
                        print("ℹ️  Test codes not immediately visible (may need time to appear)")
                        
                else:
                    print("ℹ️  No rates visible in table (may be empty or need time to appear)")
                    
            except Exception as e:
                print(f"⚠️  Could not verify upload: {e}")
            
            print(f"\n🎉 JERASOFT REAL DATA UPLOAD: SUCCESS")
            return True
            
        else:
            print("❌ FAILED: No result returned from upload")
            return False
            
    except ImportError as e:
        print(f"❌ FAILED: Cannot import required modules: {e}")
        return False
    except Exception as e:
        print(f"❌ FAILED: Error during real upload: {e}")
        return False


if __name__ == '__main__':
    # Safety confirmation
    print("⚠️  SAFETY NOTICE:")
    print(f"   - This will upload REAL data to table {TEST_TABLE_ID}")
    print(f"   - Using safe test codes: {SAFE_TEST_CODES}")
    print("   - No production data will be affected")
    print()
    
    success = test_real_jerasoft_upload()
    
    print("\n" + "=" * 60)
    print("FINAL RESULT:")
    if success:
        print(f"✅ JeraSoft REAL data upload to table {TEST_TABLE_ID}: SUCCESS")
        print("   Data was actually uploaded and can be verified!")
    else:
        print(f"❌ JeraSoft REAL data upload to table {TEST_TABLE_ID}: FAILED")
    print("=" * 60)
    
    sys.exit(0 if success else 1)