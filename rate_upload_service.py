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
from rate_upload_to_Jera import (
    bulk_upload_df_to_jerasoft,
    stash_rates_df_to_jerasoft,
    stash_future_rates_df_to_jerasoft,
    wait_for_imported_rates_visible,
)
import pandas as pd
import json
import re
import tempfile

# Load environment
load_dotenv()

def _read_metadata_if_present(comparison_file_path: str | None) -> Dict[str, Any] | None:
    """
    Best-effort: load attachments/<folder>/metadata.json (if present).
    """
    cfp = (comparison_file_path or "").strip()
    if not cfp:
        return None
    try:
        folder = Path(cfp).parent
        meta_path = folder / "metadata.json"
        if not meta_path.exists():
            return None
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        return meta if isinstance(meta, dict) else None
    except Exception:
        return None

def _pick_vendor_received_filename(comparison_file_path: str | None) -> str | None:
    """
    Get the vendor's received attachment filename from metadata.json.
    Prefer non-jerasoft_comparison files.
    """
    meta = _read_metadata_if_present(comparison_file_path)
    if not meta:
        return None
    atts = meta.get("attachments")
    if isinstance(atts, list):
        for a in atts:
            if not isinstance(a, str):
                continue
            n = a.lower()
            if "jerasoft_comparison" in n:
                continue
            if a.strip():
                return a.strip()
        for a in atts:
            if isinstance(a, str) and a.strip():
                return a.strip()
    return None

def _safe_stem(filename: str, *, max_len: int = 120) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return (stem or "vendor_file")[:max_len]

def _pick_temp_csv_path_from_vendor_filename(vendor_filename: str) -> str:
    """
    Build a temp CSV path whose basename is just the vendor filename stem:
      HAYOTEL-PREMIUM_A-Z-RN_01-13-2026.csv
    If it exists, add _v2/_v3/... to avoid collisions.
    """
    safe = _safe_stem(vendor_filename, max_len=120)
    tmp = Path(tempfile.gettempdir())
    base = tmp / f"{safe}.csv"
    if not base.exists():
        return str(base)
    for i in range(2, 51):
        alt = tmp / f"{safe}_v{i}.csv"
        if not alt.exists():
            return str(alt)
    return str(tmp / f"{safe}_{int(time.time())}.csv")

def export_db_upload_dataframe(
    df: pd.DataFrame,
    *,
    upload_id: int,
    comparison_file_path: str | None = None,
    suffix: str = "accepted_db_rows",
) -> Path:
    """
    Write the DB-built upload DataFrame to an audit file under attachments.

    If comparison_file_path points into attachments/<folder>/..., we write alongside it.
    Otherwise, we write to ./attachments/.
    """
    base_dir = Path("attachments")
    cfp = (comparison_file_path or "").strip()
    if cfp:
        try:
            p = Path(cfp)
            # if it's just a filename, p.parent will be '.' → keep base_dir as attachments/
            if str(p.parent) not in ("", "."):
                base_dir = p.parent
        except Exception:
            pass

    try:
        base_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # last resort: current dir
        base_dir = Path(".")

    out_path = base_dir / f"upload_{upload_id}_{suffix}.xlsx"
    df.to_excel(out_path, index=False)
    return out_path

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
    
    # `rate_uploads.jera_upload_result` is JSON/JSONB in many DBs.
    # Requirement: when upload is successful, store NOTHING (NULL) in jera_upload_result.
    # We only persist details on failures.
    import json as _json
    payload = None
    if status == "failed" and result is not None:
        payload = result if isinstance(result, dict) else {"result": str(result)}
    jera_result_json = None if payload is None else _json.dumps(payload)
    
    # For console logging only
    error_message = None
    if status == 'failed' and payload is not None:
        if isinstance(payload, dict):
            error_message = payload.get('error', str(payload))
        else:
            error_message = str(payload)
    
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
            cur.execute(sql, (status, jera_result_json, status, upload_id))
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
    comparison_file_path = upload.get("comparison_file_path")
    
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

        # Split out "Stashed" rows: these should be applied via API (rates.update(status="stashed"))
        # so they show as stashed/inactive in Jera UI.
        stashed_df = pd.DataFrame()
        if "Change Type" in df.columns:
            stashed_mask = df["Change Type"].astype(str).str.strip().str.lower().eq("stashed")
            if stashed_mask.any():
                stashed_df = df.loc[stashed_mask].copy()
                df = df.loc[~stashed_mask].copy()

        # JeraSoft "Closed keywords" support:
        # If a row is marked Closed, send the keyword "close" in the Rate (New Rate/value) column so Jera will close it.
        if "Change Type" in df.columns and "New Rate" in df.columns:
            closed_mask = df["Change Type"].astype(str).str.strip().str.lower().eq("closed")
            if closed_mask.any():
                df.loc[closed_mask, "New Rate"] = "close"

        # Write an audit copy of what we are about to upload (DB-built sheet)
        try:
            out_path = export_db_upload_dataframe(
                df,
                upload_id=int(upload_id),
                comparison_file_path=str(comparison_file_path) if comparison_file_path else None,
            )
            print(f"📝 Wrote DB upload sheet: {out_path}")
        except Exception as e:
            print(f"⚠️ Failed to write DB upload sheet for upload_id={upload_id}: {e}")
        
        # Mark as processing
        if not dry_run:
            update_bulk_upload_status(upload_id, 'processing')
        
        # 1) Perform bulk upload (normal + close-keyword rows)
        # Use vendor filename as the uploaded CSV name so Jera Import History shows it.
        temp_csv_path = None
        vendor_file = _pick_vendor_received_filename(str(comparison_file_path) if comparison_file_path else None)
        if vendor_file:
            temp_csv_path = _pick_temp_csv_path_from_vendor_filename(vendor_file)

        # If you want JeraSoft Import History to show "A-Z Stashed" like manual imports,
        # create an import template in JeraSoft with "Stash Future Rates" enabled and set:
        #   JERASOFT_IMPORT_TEMPLATE_ID=<id>
        import_template_id = os.getenv("JERASOFT_IMPORT_TEMPLATE_ID", "").strip()
        import_template_id_int = int(import_template_id) if import_template_id.isdigit() else None

        result = bulk_upload_df_to_jerasoft(
            df=df,
            table_id=table_id,
            accepted_statuses=accepted_statuses,
            import_templates_id=import_template_id_int,
            temp_file_path=temp_csv_path,
            dry_run=dry_run,
        )

        # 1a) Wait for import job to actually apply before doing any post-import API updates.
        # JeraSoft import is async; if we stash too early, the import can overwrite status back to "active".
        if not dry_run and isinstance(result, dict) and result.get("status") == "success":
            enable_wait = os.getenv("JERASOFT_WAIT_IMPORT_APPLY", "1").strip() not in ("0", "false", "False", "no", "NO")
            if enable_wait:
                try:
                    # Use the same "imported_df" filtering as bulk_upload_df_to_jerasoft
                    _wait_df = df.copy()
                    if accepted_statuses and "Status" in _wait_df.columns:
                        _wait_df["Status"] = _wait_df["Status"].astype(str).str.strip()
                        _wait_df = _wait_df[_wait_df["Status"].isin(accepted_statuses)]
                    _wait_df = _wait_df[_wait_df["Code"].astype(str).str.strip().ne("")]
                    _wait_df = _wait_df[_wait_df["New Rate"].notna()]
                    _wait_df = _wait_df[_wait_df["Effective Date"].notna()]
                    if not _wait_df.empty:
                        wait_res = wait_for_imported_rates_visible(_wait_df, int(table_id))
                        print(f"⏳ Import visibility wait: {wait_res}")
                        result = {**result, "import_wait": wait_res}
                except Exception as e:
                    print(f"⚠️ Import wait failed (continuing): {e}")

        # 1b) Optional: "Stash Future Rates" behavior (like Jera UI option)
        # After importing a rate, stash any later ACTIVE rates for the same code in the same table.
        # This is useful for "rescheduled/backdated" scenarios where a new earlier effective date
        # should make already-scheduled future rates inactive (stashed) in Jera UI.
        enable_stash_future = os.getenv("JERASOFT_ENABLE_STASH_FUTURE_RATES", "1").strip() not in ("0", "false", "False", "no", "NO")
        if enable_stash_future and isinstance(result, dict) and result.get("status") == "success":
            try:
                imported_df = df.copy()
                # Match bulk_upload_df_to_jerasoft filtering
                if accepted_statuses and "Status" in imported_df.columns:
                    imported_df["Status"] = imported_df["Status"].astype(str).str.strip()
                    imported_df = imported_df[imported_df["Status"].isin(accepted_statuses)]
                # Skip "Closed" rows: those are handled via close keyword, not future-stashing.
                if "Change Type" in imported_df.columns:
                    ct = imported_df["Change Type"].astype(str).str.strip().str.lower()
                    imported_df = imported_df[~ct.eq("closed")]
                # Drop invalids (same as uploader)
                imported_df = imported_df[imported_df["Code"].astype(str).str.strip().ne("")]
                imported_df = imported_df[imported_df["New Rate"].notna()]
                imported_df = imported_df[imported_df["Effective Date"].notna()]
                if not imported_df.empty:
                    print(f"🧊 Stashing future rates for {len(imported_df)} imported row(s) in JeraSoft...")
                    future_stash_result = stash_future_rates_df_to_jerasoft(imported_df, int(table_id), dry_run=dry_run)
                    print(f"🧊 Future-stash result: {future_stash_result}")
                    result = {**result, "future_stashed": future_stash_result}
            except Exception as e:
                print(f"⚠️ Failed to apply future-stash behavior: {e}")

        # 2) Apply stashed rows via API
        stash_result: Dict[str, Any] = {"status": "skipped", "reason": "no_stashed_rows"}
        if stashed_df is not None and not stashed_df.empty:
            enable_stash_api = os.getenv("JERASOFT_ENABLE_STASH_API", "1").strip() not in ("0", "false", "False", "no", "NO")
            if enable_stash_api:
                print(f"🧊 Applying {len(stashed_df)} Stashed rows in JeraSoft...")
                stash_result = stash_rates_df_to_jerasoft(stashed_df, int(table_id), dry_run=dry_run)
                print(f"🧊 Stash result: {stash_result}")
                result = {**result, "stashed": stash_result}
            else:
                print(f"🧊 Skipping stashed API updates (JERASOFT_ENABLE_STASH_API=0).")
                stash_result = {"status": "skipped", "reason": "stash_api_disabled", "rows": int(len(stashed_df))}
                result = {**result, "stashed": stash_result}
        
        # Overall success rules:
        # - Normal case: bulk import succeeded.
        # - Stash-only case: bulk import returns "skipped/no_accepted_rates" BUT stashing succeeded.
        bulk_ok = result.get("status") == "success"
        stash_ok = True
        if stashed_df is not None and not stashed_df.empty:
            stash_ok = (
                isinstance(stash_result, dict)
                and stash_result.get("status") == "success"
                and int(stash_result.get("not_found", 0)) == 0
                and not stash_result.get("errors")
            )
        only_stash_ok = (
            (not bulk_ok)
            and result.get("status") == "skipped"
            and result.get("reason") == "no_accepted_rates"
            and (stashed_df is not None and not stashed_df.empty)
            and stash_ok
        )

        if bulk_ok or only_stash_ok:
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