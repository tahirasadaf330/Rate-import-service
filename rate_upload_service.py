#!/usr/bin/env python3
"""
Rate Upload Service for JeraSoft Rate Imports

This service processes rate_uploads marked for bulk upload and uploads
entire comparison files (not individual rates) to JeraSoft.

Usage:
    python rate_upload_service.py --once    # Process once and exit
    python rate_upload_service.py --list    # List pending uploads
    python rate_upload_service.py --dry-run # Show what would be uploaded
"""

import os
import sys
import time
import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

from dotenv import load_dotenv
from database import get_conn
from database import fetch_rate_upload_details_for_upload
from rate_upload_to_Jera import bulk_upload_df_to_jerasoft
import pandas as pd

# Load environment
load_dotenv()

def get_pending_bulk_uploads() -> List[Dict[str, Any]]:
    """Get rate_uploads marked for bulk upload."""
    sql = """
        SELECT 
            id, subject, sender_email, jera_table_id,
            jera_upload_status, jera_upload_result,
            total_rows, created_at, comparison_file_path
        FROM rate_uploads 
        WHERE is_rate_approved_by_admin = TRUE 
        AND jera_upload_status = 'pending_bulk'
        AND jera_table_id IS NOT NULL
        ORDER BY created_at ASC
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()
        
        columns = [desc[0] for desc in cur.description]
        return [dict(zip(columns, row)) for row in rows]

def update_bulk_upload_status(upload_id: int, status: str, result: Dict[str, Any] = None):
    """Update the upload status and result."""
    
    # Only store error messages as string when status is failed, otherwise null
    error_message = None
    if status == 'failed' and result:
        if isinstance(result, dict):
            # Extract error message from result dict
            error_message = result.get('error', str(result))
        else:
            error_message = str(result)
    
    sql = """
        UPDATE rate_uploads SET
            jera_upload_status = %s,
            jera_upload_result = %s,
            jera_uploaded_at = CASE WHEN %s = 'completed' THEN NOW() ELSE jera_uploaded_at END,
            updated_at = NOW()
        WHERE id = %s
    """
    
    try:
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql, (status, error_message, status, upload_id))
            rows_affected = cur.rowcount
            conn.commit()
            
            if status == 'completed':
                print(f"✅ SUCCESS: Upload {upload_id} completed successfully")
            elif status == 'failed':
                print(f"❌ FAILED: Upload {upload_id} failed - {error_message}")
            else:
                print(f"🔄 Database update: upload_id={upload_id}, status='{status}', rows_affected={rows_affected}")
            
            if rows_affected == 0:
                print(f"⚠️ WARNING: No rows were updated for upload_id {upload_id}")
                
    except Exception as e:
        print(f"❌ Database update failed for upload_id {upload_id}: {e}")
        import traceback
        traceback.print_exc()
        raise

def verify_upload_status(upload_id: int):
    """Verify the current status of an upload in the database."""
    sql = """
        SELECT id, jera_upload_status, jera_uploaded_at, 
               jera_upload_result as error_message
        FROM rate_uploads 
        WHERE id = %s
    """
    
    try:
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql, (upload_id,))
            row = cur.fetchone()
            
            if row:
                print(f"📊 Database Status Check for upload_id {upload_id}:")
                print(f"   Current Status: {row[1]}")
                print(f"   Uploaded At: {row[2]}")
                print(f"   Result Status: {row[3]}")
            else:
                print(f"❌ Upload ID {upload_id} not found in database")
                
    except Exception as e:
        print(f"❌ Status verification failed: {e}")

def process_bulk_upload(upload: Dict[str, Any], dry_run: bool = False) -> bool:
    """
    Process a single bulk upload.
    
    Args:
        upload: Upload record from database
        dry_run: If True, don't actually upload
        
    Returns:
        True if successful, False otherwise
    """
    upload_id = upload['id']
    table_id = upload['jera_table_id']
    
    print(f"\n🔄 Processing upload ID {upload_id}")
    print(f"   Subject: {upload['subject']}")
    print(f"   Sender: {upload['sender_email']}")
    print(f"   Target table: {table_id}")
    print(f"   Total rows: {upload['total_rows']}")
    
    try:
        # Build upload "sheet" from DB rows (reflects UI-approved statuses)
        statuses_env = os.getenv("JERASOFT_UPLOAD_STATUSES", "Accepted")
        accepted_statuses = tuple(s.strip() for s in (statuses_env or "").split(",") if s.strip()) or ("Accepted",)

        details = fetch_rate_upload_details_for_upload(upload_id, statuses=accepted_statuses)
        df = pd.DataFrame(details or [])
        if df.empty:
            print("⚠️ No accepted rows found in DB for this upload_id")
            update_bulk_upload_status(upload_id, 'failed', {
                'error': 'no_accepted_rows_in_db',
                'accepted_statuses': list(accepted_statuses),
                'processed_at': datetime.now().isoformat()
            })
            return False

        # Map DB column names -> uploader expected names
        rename = {
            "dst_code": "Code",
            "rate_new": "New Rate",
            "effective_date": "Effective Date",
            "new_billing_increment": "New Billing Increment",
            "status": "Status",
            "code_name": "Dst Code Name",
            "notes": "Notes",
            "change_type": "Change Type",
        }
        df = df.rename(columns=rename)
        
        # Mark as processing
        if not dry_run:
            update_bulk_upload_status(upload_id, 'processing')
        
        # Perform bulk upload
        result = bulk_upload_df_to_jerasoft(
            df=df,
            table_id=table_id,
            accepted_statuses=accepted_statuses,
            dry_run=dry_run,
        )
        
        if result.get('status') == 'success':
            print(f"✅ Bulk upload completed successfully!")
            print(f"   Rows uploaded: {result.get('filtered_rows', 0)}")
            print(f"   Files ID: {result.get('files_id', 'N/A')}")
            
            if not dry_run:
                update_bulk_upload_status(upload_id, 'completed', {
                    **result,
                    'processed_at': datetime.now().isoformat()
                })
                # Verify the update was successful
                verify_upload_status(upload_id)
            return True
        
        else:
            print(f"❌ Bulk upload failed: {result.get('error', 'Unknown error')}")
            if not dry_run:
                update_bulk_upload_status(upload_id, 'failed', {
                    **result,
                    'processed_at': datetime.now().isoformat()
                })
            return False
            
    except Exception as e:
        print(f"❌ Exception during bulk upload: {e}")
        import traceback
        traceback.print_exc()
        
        if not dry_run:
            update_bulk_upload_status(upload_id, 'failed', {
                'error': str(e),
                'exception_type': type(e).__name__,
                'processed_at': datetime.now().isoformat()
            })
        return False

def list_pending_uploads():
    """List all pending bulk uploads."""
    uploads = get_pending_bulk_uploads()
    
    if not uploads:
        print("📝 No pending bulk uploads found.")
        return
    
    print(f"📝 Found {len(uploads)} pending bulk upload(s):")
    print("-" * 80)
    
    for upload in uploads:
        print(f"ID: {upload['id']}")
        print(f"Subject: {upload['subject']}")
        print(f"Sender: {upload['sender_email']}")
        print(f"Table ID: {upload['jera_table_id']}")
        print(f"Rows: {upload['total_rows']}")
        print(f"Created: {upload['created_at']}")
        print("-" * 40)

def show_upload_history():
    """Show recent upload history with statuses."""
    sql = """
        SELECT id, subject, sender_email, jera_table_id,
               jera_upload_status, jera_uploaded_at, 
               jera_upload_result as error_message,
               created_at
        FROM rate_uploads 
        WHERE is_rate_approved_by_admin = TRUE 
        ORDER BY created_at DESC 
        LIMIT 10
    """
    
    try:
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()
            
            if not rows:
                print("📝 No upload records found.")
                return
            
            print(f"📊 Recent Upload History ({len(rows)} records):")
            print("-" * 100)
            
            for row in rows:
                upload_id, subject, sender, table_id, status, uploaded_at, result_status, created_at = row
                print(f"ID: {upload_id} | Status: {status} | Result: {result_status}")
                print(f"Subject: {subject[:50]}...")
                print(f"Sender: {sender} | Table: {table_id}")
                print(f"Created: {created_at} | Uploaded: {uploaded_at}")
                print("-" * 50)
                
    except Exception as e:
        print(f"❌ Failed to fetch upload history: {e}")

def process_all_uploads(dry_run: bool = False) -> Dict[str, int]:
    """Process all pending bulk uploads."""
    uploads = get_pending_bulk_uploads()
    
    if not uploads:
        print("📝 No pending bulk uploads to process.")
        return {'total': 0, 'success': 0, 'failed': 0}
    
    print(f"🚀 Processing {len(uploads)} bulk upload(s)...")
    if dry_run:
        print("🧪 DRY RUN MODE - No actual uploads will be performed")
    
    results = {'total': len(uploads), 'success': 0, 'failed': 0}
    
    for upload in uploads:
        success = process_bulk_upload(upload, dry_run=dry_run)
        if success:
            results['success'] += 1
        else:
            results['failed'] += 1
    
    print(f"\n📊 SUMMARY:")
    print(f"   Total: {results['total']}")
    print(f"   Success: {results['success']}")
    print(f"   Failed: {results['failed']}")
    
    return results

def main():
    parser = argparse.ArgumentParser(description="JeraSoft Bulk Upload Service")
    parser.add_argument('--once', action='store_true', 
                       help='Process pending uploads once and exit')
    parser.add_argument('--list', action='store_true',
                       help='List pending uploads')
    parser.add_argument('--history', action='store_true',
                       help='Show recent upload history with statuses')
    parser.add_argument('--dry-run', action='store_true',
                       help='Show what would be uploaded without doing it')
    
    args = parser.parse_args()
    
    print("🎯 JeraSoft Bulk Upload Service")
    print("=" * 50)
    
    if args.list:
        list_pending_uploads()
    elif args.history:
        show_upload_history()
    elif args.dry_run:
        process_all_uploads(dry_run=True)
    elif args.once:
        process_all_uploads()
    else:
        # No arguments provided - check for status updates first, then process
        from database import auto_update_status_on_import_flag_change
        
        # First, check for any manual is_rate_approved_by_admin changes
        updated_count = auto_update_status_on_import_flag_change()
        
        # Then check for pending uploads and process automatically
        uploads = get_pending_bulk_uploads()
        if uploads:
            print(f"📋 Found {len(uploads)} pending bulk upload(s)")
            print("🚀 Processing uploads automatically...")
            process_all_uploads()
        else:
            if updated_count == 0:
                print("📝 No pending bulk uploads found.")
                print("💡 System is up to date!")

if __name__ == "__main__":
    main()