# ───────────────────────── Push comparison results to JeraSoft ─────────────────────────
"""
JeraSoft Rate Upload Engine

PURPOSE:
This file handles the actual upload of rate changes to JeraSoft billing system.
It processes comparison files containing rate updates and pushes them to JeraSoft tables.

WHAT IT DOES:
• Reads Excel comparison files with rate changes (New Rate, Old Rate, Status, etc.)
• Filters rows by status (only uploads "Accepted" rates by default)  
• Connects to JeraSoft API and Web interface for bulk uploads
• Supports both individual rate updates and bulk file imports
• Handles authentication, file validation, and upload progress tracking
• Returns detailed results including success/failure counts and error logs

WHEN TO USE:
• Called by rate_upload_service.py for approved bulk uploads
• Processes files from database records marked with import_to_jera=TRUE
• Can be used standalone for direct rate file uploads to JeraSoft

KEY FUNCTIONS:
• push_rates_to_jerasoft() - Main upload function
• validate_and_filter_data() - Filters rates by status before upload
• upload_to_jerasoft_bulk() - Handles bulk file upload to JeraSoft web interface
"""

from typing import Optional, Dict, List, Tuple
import re
import pandas as pd
import os
import json
import requests
import tempfile
from pathlib import Path

# Environment variables for JeraSoft configuration
J_API_URL = os.getenv("JERASOFT_API_URL", "http://billing.voipsystem.org:3080") 
J_WEB_URL = os.getenv("JERASOFT_WEB_URL", "https://billing.voipsystem.org:443")
J_API_KEY = os.getenv("JERA_SOFT_API_KEY")  # CoreAPI token
J_WEB_LOGIN = os.getenv("JERASOFT_WEB_LOGIN")  # Web interface login
J_WEB_PASSWORD = os.getenv("JERASOFT_WEB_PASSWORD")  # Web interface password

# Shared session with retries (optional but recommended)
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
_session_push = requests.Session()
_session_push.mount("http://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=0.4, status_forcelist=[502,503,504])))
_session_push.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=0.4, status_forcelist=[502,503,504])))

def _rpc_call(method: str, params: Dict, api_url: Optional[str] = None) -> Dict:
    """Low-level JSON-RPC helper."""
    api_url = api_url or J_API_URL
    payload = {"jsonrpc": "2.0", "id": 1, "method": method, "params": params}
    r = _session_push.post(api_url, headers={"Content-Type":"application/json","Accept":"application/json"}, json=payload, timeout=300)
    r.raise_for_status()
    data = r.json()
    if "error" in data and data["error"]:
        # JeraSoft error comes here
        raise RuntimeError(f"JeraSoft RPC error in {method}: {json.dumps(data['error'])}")
    return data.get("result")

_bi_re = re.compile(r"^\s*(\d+)\s*/\s*(\d+)\s*$")
def _parse_billing_increment(bi: Optional[str]) -> Tuple[Optional[int], Optional[int]]:
    """'60/60' -> (60,60); returns (None,None) if blank/invalid."""
    if not isinstance(bi, str):
        return None, None
    m = _bi_re.match(bi.strip())
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))

def _fmt_date(d: object) -> Optional[str]:
    """Accepts datetime/str; returns 'YYYY-MM-DD' or None."""
    if d is None or (isinstance(d, float) and pd.isna(d)):
        return None
    if isinstance(d, str):
        d = d.strip()
        if not d:
            return None
        try:
            return str(pd.to_datetime(d, errors="coerce").date())
        except Exception:
            return None
    try:
        return str(pd.to_datetime(d, errors="coerce").date())
    except Exception:
        return None

def _search_rate_id(table_id: int, code: str, effective_from: str, api_url: Optional[str]=None) -> Optional[int]:
    """
    Find an existing rate row id by table + code.
    Note: JeraSoft rates.search doesn't support effective_from filter.
    Returns None if not found.
    """
    params = {
        "AUTH": J_API_KEY,
        "rate_tables_id": table_id,
        "code": code,
        "status": "active",  # JeraSoft only accepts 'active' or 'stashed'
        "limit": 10,  # Get multiple results to filter by effective_from manually
        "offset": 0,
    }
    res = _rpc_call("rates.search", params, api_url=api_url)
    if isinstance(res, list) and res:
        # Filter results manually by effective_from since API doesn't support it
        for rate in res:
            if rate.get("effective_from") == effective_from:
                rid = rate.get("id")
                return int(rid) if rid is not None else None
        
        # If no exact match found, return the first result for now
        # This allows updates to work even if effective_from doesn't match exactly
        rid = res[0].get("id")
        return int(rid) if rid is not None else None
    return None

def _create_rate(table_id: int, code: str, value: float, eff: str,
                 min_vol: Optional[int], pay_int: Optional[int],
                 api_url: Optional[str]=None) -> int:
    """Create a rate row and return its id."""
    params = {
        "AUTH": J_API_KEY,
        "rate_tables_id": table_id,
        "code": str(code),
        "effective_from": eff,
        "value": float(value),
        "status": "active",
        "time_profiles_id": "1",  # Default time profile (required by JeraSoft)
    }
    if min_vol is not None: params["min_volume"] = int(min_vol)
    if pay_int is not None: params["pay_interval"] = int(pay_int)
    res = _rpc_call("rates.create", params, api_url=api_url)
    
    # Handle both dict and list responses from JeraSoft
    if isinstance(res, dict):
        rid = res.get("id")
    elif isinstance(res, list) and res:
        rid = res[0].get("id")
    else:
        rid = None
        
    if rid is None:
        raise RuntimeError(f"rates.create returned no id: {res}")
    return int(rid)

def _update_rate(rate_id: int, value: Optional[float]=None,
                 min_vol: Optional[int]=None, pay_int: Optional[int]=None,
                 status: Optional[str]=None, api_url: Optional[str]=None) -> None:
    """Update a rate row by id (only provided fields are changed)."""
    params = {"AUTH": J_API_KEY, "id": int(rate_id)}
    if value is not None:    params["value"] = float(value)
    if min_vol is not None:  params["min_volume"] = int(min_vol)
    if pay_int is not None:  params["pay_interval"] = int(pay_int)
    if status is not None:   params["status"] = status
    _ = _rpc_call("rates.update", params, api_url=api_url)

# ──────────────────────── BULK UPLOAD WORKFLOW ────────────────────────
def get_jwt_token(web_url: Optional[str] = None, login: Optional[str] = None, 
                  password: Optional[str] = None) -> str:
    """Get JWT token for file upload authentication."""
    web_url = web_url or J_WEB_URL
    login = login or J_WEB_LOGIN  
    password = password or J_WEB_PASSWORD
    
    if not all([web_url, login, password]):
        raise ValueError("Missing web credentials: JERASOFT_WEB_URL, JERASOFT_WEB_LOGIN, JERASOFT_WEB_PASSWORD")
    
    params = {
        '__api': 1,
        'auth[login]': login,
        'auth[password]': password
    }
    
    response = requests.post(
        f'{web_url}/admin/_auth/jwt',
        params=params,
        verify=False,
        timeout=30
    )
    
    if response.status_code != 200:
        raise RuntimeError(f"Cannot get JWT token. Status: {response.status_code}, Response: {response.text}")
    
    return response.json()['token']

def upload_file_to_jerasoft(file_path: str, web_url: Optional[str] = None) -> str:
    """Upload file to JeraSoft and return files_id for bulk import."""
    web_url = web_url or J_WEB_URL
    
    print(f"🔐 Getting JWT token...")
    token = get_jwt_token(web_url)
    
    print(f"📤 Uploading file: {file_path}")
    with open(file_path, 'rb') as file:
        file_bytes = file.read()
    
    response = requests.post(
        f'{web_url}/upload',
        data=file_bytes,
        headers={
            "Authorization": f"Bearer {token}",
            'X-File-Name': os.path.basename(file_path)
        },
        verify=False,
        timeout=120
    )
    
    if response.status_code != 200:
        raise RuntimeError(f"File upload failed. Status: {response.status_code}, Response: {response.text}")
    
    files_id = response.json()['id']
    print(f"✅ File uploaded successfully. files_id: {files_id}")
    return files_id

def upload_rates_to_jerasoft(df: pd.DataFrame, table_id: int,
                            *, dry_run: bool = False,
                            force_bulk: Optional[bool] = None,
                            bulk_min_rows: int = 50) -> Dict:
    """
    Main function to upload rates to JeraSoft.
    Automatically chooses between bulk and individual upload based on configuration.
    
    Args:
        df: DataFrame with rate comparison data
        table_id: JeraSoft table ID 
        dry_run: If True, don't actually upload
        force_bulk: If True, always use bulk. If False, always use individual. If None, auto-decide
        bulk_min_rows: Minimum rows to trigger bulk upload (if not forced)
        
    Returns:
        Dict with upload results
    """
    # Check environment configuration
    env_force_bulk = os.getenv("JERASOFT_FORCE_BULK_UPLOAD", "false").lower() in ("true", "1", "yes")
    env_bulk_min = int(os.getenv("JERASOFT_BULK_UPLOAD_MIN_ROWS", str(bulk_min_rows)))
    
    # Determine upload method
    use_bulk = force_bulk if force_bulk is not None else (env_force_bulk or len(df) >= env_bulk_min)
    
    print(f"📊 Upload Method Decision:")
    print(f"   • Rows to upload: {len(df)}")
    print(f"   • Force bulk (env): {env_force_bulk}")
    print(f"   • Force bulk (param): {force_bulk}")
    print(f"   • Bulk min rows: {env_bulk_min}")
    print(f"   • 🎯 Selected method: {'BULK' if use_bulk else 'INDIVIDUAL'}")
    
    if use_bulk:
        return bulk_import_rates(df, table_id, dry_run=dry_run)
    else:
        return push_comparison_to_jerasoft(df, table_id, dry_run=dry_run)

def bulk_import_rates(df: pd.DataFrame, table_id: int, 
                     import_templates_id: Optional[int] = None,
                     temp_file_path: Optional[str] = None,
                     dry_run: bool = False) -> Dict:
    """
    Bulk import rates using JeraSoft's official workflow:
    1. Create CSV file from DataFrame
    2. Upload file to get files_id  
    3. Call rates.imports.prepare
    4. Call rates.imports.enqueue
    """
    if not J_API_KEY:
        raise ValueError("Missing JERA_SOFT_API_KEY")
    
    # Default import settings (can be overridden with import_templates_id)
    default_settings = {
        "agreements_tolerance": "",
        "az_codes": None,
        "az_interval_type": "days", 
        "az_interval_value": None,
        "az_mode": "",
        "billing_increment_check": False,
        "billing_increment_format": ["grace_volume", "pay_interval", "min_volume"],
        "blocked_keywords": ["block"],
        "closed_keywords": ["close", "delete", "terminate", "deactivate", "remove"],
        "code_deck_mode": "",
        "code_decks_id": "1",
        "datetime_format": "mdy",
        "error_mode": "skip_rows",
        "rate_deviation_tolerance": "",
        "sheets": [{
            "code_rules": [{
                "effective_from_interval_type": "days",
                "effective_from_interval_value": "",
                "end_date_interval_type": "days", 
                "end_date_interval_value": "",
                "grace_volume": "0",
                "min_volume": "1",
                "notes": "",
                "num_length_max": "",
                "num_length_min": "",
                "pay_interval": "1", 
                "pay_setup": "0",
                "status": "active",
                "tag": "@",
                "time_profiles_id": "1"
            }],
            "columns": ["code_name", "code", "value", "effective_from"],
            "skip_rows_bottom": "0",
            "skip_rows_top": "0", 
            "type": "rates"
        }],
        "skip_dash": False,
        "split_src_code_name": "0",
        "src_code_decks_id": "",
        "total_changes_threshold": "",
        "unchanged_mode": "skip_rows",
        "warning_mode": "save_rows"
    }
    
    # Step 1: Prepare DataFrame and create CSV
    print(f"📊 Preparing {len(df)} rows for bulk import...")
    
    # Map DataFrame columns to JeraSoft format
    jera_df = df.copy()
    if 'Code' in jera_df.columns:
        jera_df['code'] = jera_df['Code']
        jera_df['code_name'] = jera_df['Code']  # Use code as name if no separate name column
    if 'New Rate' in jera_df.columns:
        jera_df['value'] = jera_df['New Rate']
    if 'Effective Date' in jera_df.columns:
        jera_df['effective_from'] = pd.to_datetime(jera_df['Effective Date']).dt.strftime('%Y-%m-%d')
    
    # Select only required columns for JeraSoft
    required_cols = ['code_name', 'code', 'value', 'effective_from']
    missing_cols = [col for col in required_cols if col not in jera_df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns after mapping: {missing_cols}")
    
    export_df = jera_df[required_cols]
    
    if dry_run:
        print("🧪 DRY RUN - Preview of data to be uploaded:")
        print(export_df.head(10))
        return {"status": "dry_run", "rows_prepared": len(export_df)}
    
    # Step 2: Create temporary CSV file
    if temp_file_path:
        csv_path = temp_file_path
    else:
        temp_dir = tempfile.gettempdir()
        csv_path = os.path.join(temp_dir, f"jera_bulk_import_{table_id}.csv")
    
    print(f"💾 Saving CSV to: {csv_path}")
    export_df.to_csv(csv_path, index=False)
    
    try:
        # Step 3: Upload file to JeraSoft
        files_id = upload_file_to_jerasoft(csv_path)
        
        # Step 4: Prepare import
        print(f"🔧 Calling rates.imports.prepare...")
        prepare_result = _rpc_call("rates.imports.prepare", {
            "files_id": files_id,
            "rate_tables_id": table_id,
            "AUTH": J_API_KEY
        })
        
        import_queue_id = prepare_result['id']
        print(f"✅ Import prepared. Queue ID: {import_queue_id}")
        
        # Step 5: Enqueue import
        print(f"🚀 Calling rates.imports.enqueue...")
        enqueue_params = {
            "import_queue_id": import_queue_id,
            "AUTH": J_API_KEY
        }
        
        # Use either template or settings
        if import_templates_id:
            enqueue_params["import_templates_id"] = import_templates_id
        else:
            enqueue_params["settings"] = default_settings
        
        enqueue_result = _rpc_call("rates.imports.enqueue", enqueue_params)
        
        print(f"✅ Bulk import enqueued successfully!")
        return {
            "status": "success",
            "files_id": files_id,
            "import_queue_id": import_queue_id,
            "enqueue_result": enqueue_result,
            "rows_uploaded": len(export_df),
            "csv_file": csv_path
        }
        
    finally:
        # Cleanup temporary file if we created it
        if not temp_file_path and os.path.exists(csv_path):
            try:
                os.remove(csv_path)
                print(f"🗑️ Cleaned up temporary file: {csv_path}")
            except Exception as e:
                print(f"⚠️ Could not remove temporary file {csv_path}: {e}")

# ──────────────────────── INDIVIDUAL UPLOAD (ORIGINAL METHOD) ────────────────────────
def push_comparison_to_jerasoft(df: pd.DataFrame,
                                table_id: int,
                                *,
                                api_url: Optional[str]=None,
                                api_key: Optional[str]=None,
                                accepted_statuses: Tuple[str, ...]=("Accepted",),
                                create_if_missing: bool=True,
                                update_if_exists: bool=True,
                                dry_run: bool=False) -> Dict[str,int]:
    """
    Push comparison results into a JeraSoft rate table (upsert).
    Expects columns:
      'Code', 'New Rate', 'Effective Date', 'New Billing Increment' (optional), 'Status' (optional)

    Only rows with Status ∈ accepted_statuses are uploaded.

    Returns a summary dict.
    """
    if api_key:
        # allow overriding process-wide key
        global J_API_KEY
        J_API_KEY = api_key

    if not J_API_KEY:
        raise ValueError("Missing JERA_SOFT_API_KEY (env) or pass api_key=)")

    req_cols = ["Code","New Rate","Effective Date"]
    for c in req_cols:
        if c not in df.columns:
            raise ValueError(f"Missing required column in DataFrame: {c}")

    summary = {"processed": 0, "created": 0, "updated": 0, "skipped_status": 0, "skipped_invalid": 0, "errors": 0}

    for idx, row in df.iterrows():
        summary["processed"] += 1

        # Filter by status (optional)
        if "Status" in df.columns and accepted_statuses:
            st = str(row.get("Status") or "").strip()
            if st not in accepted_statuses:
                summary["skipped_status"] += 1
                continue

        code = str(row.get("Code") or "").strip()
        rate = row.get("New Rate")
        eff  = _fmt_date(row.get("Effective Date"))
        bi   = row.get("New Billing Increment")
        mv, pi = _parse_billing_increment(bi)

        if not code or eff is None:
            summary["skipped_invalid"] += 1
            continue
        try:
            rate = float(rate)
        except Exception:
            summary["skipped_invalid"] += 1
            continue

        if dry_run:
            print(f"[DRY-RUN] upsert code={code} eff={eff} rate={rate} min_vol={mv} pay_int={pi}")
            continue

        try:
            rid = _search_rate_id(table_id, code, eff, api_url=api_url)

            if rid is None:
                if not create_if_missing:
                    summary["skipped_invalid"] += 1
                    continue
                _ = _create_rate(table_id, code, rate, eff, mv, pi, api_url=api_url)
                summary["created"] += 1
            else:
                if update_if_exists:
                    _update_rate(rid, value=rate, min_vol=mv, pay_int=pi, api_url=api_url)
                    summary["updated"] += 1
                else:
                    # exists but we chose not to update
                    pass

        except Exception as e:
            summary["errors"] += 1
            print(f"[push-error] row #{idx} code={code} eff={eff}: {e}")

    return summary




def bulk_upload_comparison_to_jerasoft(comparison_file_path: str, table_id: int, 
                                      accepted_statuses: Tuple[str, ...] = ("Accepted",),
                                      dry_run: bool = False) -> Dict:
    """
    ALWAYS use bulk upload for comparison files (no individual rate uploads).
    This function reads a comparison result file and uploads it via JeraSoft bulk import.
    
    Args:
        comparison_file_path: Path to Excel/CSV comparison result file
        table_id: JeraSoft rate table ID
        accepted_statuses: Only upload rates with these statuses
        dry_run: If True, show preview without uploading
    
    Returns:
        Result dict with upload status and details
    """
    print(f"🚀 Starting BULK upload from {comparison_file_path} to table {table_id}")
    
    # Read comparison file
    if comparison_file_path.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(comparison_file_path)
    elif comparison_file_path.endswith('.csv'):
        df = pd.read_csv(comparison_file_path)
    else:
        raise ValueError(f"Unsupported file format: {comparison_file_path}")
    
    print(f"📊 Loaded {len(df)} rows from comparison file")
    
    # Filter by status if specified
    if accepted_statuses and 'Status' in df.columns:
        original_count = len(df)
        df = df[df['Status'].isin(accepted_statuses)]
        print(f"📋 Filtered to {len(df)} rows with accepted statuses: {accepted_statuses}")
        if len(df) == 0:
            return {"status": "skipped", "reason": "no_accepted_rates", "original_rows": original_count}
    
    # Check required columns
    required_cols = ['Code', 'New Rate', 'Effective Date']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")
    
    # Use the existing bulk_import_rates function
    try:
        result = bulk_import_rates(df, table_id, dry_run=dry_run)
        
        # Add source file info to result
        result["source_file"] = comparison_file_path
        result["filtered_rows"] = len(df)
        result["upload_method"] = "bulk_file_upload"
        
        print(f"✅ Bulk upload completed! Status: {result.get('status')}")
        return result
        
    except Exception as e:
        print(f"❌ Bulk upload failed: {e}")
        return {
            "status": "error",
            "error": str(e),
            "source_file": comparison_file_path,
            "upload_method": "bulk_file_upload"
        }

# ──────────────────────── EXAMPLE USAGE ────────────────────────
def example_usage():
    """Example of how to use both upload methods."""
    
    # Example comparison results DataFrame
    sample_df = pd.DataFrame({
        'Code': ['1', '44', '91', '380', '7'],
        'New Rate': [0.05, 0.03, 0.08, 0.12, 0.15], 
        'Effective Date': ['2024-01-01', '2024-01-01', '2024-01-01', '2024-01-01', '2024-01-01'],
        'New Billing Increment': ['60/60', '1/1', '60/60', '30/30', '60/1'],
        'Status': ['Accepted', 'Accepted', 'Rejected', 'Accepted', 'Accepted']
    })
    
    table_id = 4330  # Replace with your actual JeraSoft rate table ID
    
    print("🔄 JeraSoft Rate Upload Options")
    print("=" * 50)
    
    # Method 1: Individual rate upload (for small datasets or precise control)
    print("\n📋 Method 1: Individual Rate Upload")
    print("   ✅ Best for: < 100 rates, precise error handling")
    print("   ⚠️  Slower for large datasets")
    
    individual_result = push_comparison_to_jerasoft(
        sample_df, 
        table_id=table_id,
        accepted_statuses=("Accepted",),
        dry_run=True  # Set False for actual upload
    )
    print(f"   Result: {individual_result}")
    
    # Method 2: Bulk import (for large datasets)  
    print("\n🚀 Method 2: Bulk Import")
    print("   ✅ Best for: > 100 rates, faster processing")
    print("   ⚠️  Less granular error handling")
    
    try:
        bulk_result = bulk_import_rates(
            sample_df,
            table_id=table_id,
            dry_run=True  # Set False for actual upload
        )
        print(f"   Result: {bulk_result}")
        
    except ValueError as e:
        print(f"   ❌ Configuration error: {e}")
        print("   💡 Make sure to set these environment variables:")
        print("      - JERASOFT_WEB_URL (e.g., https://billing.voipsystem.org:443)")
        print("      - JERASOFT_WEB_LOGIN (web interface username)")  
        print("      - JERASOFT_WEB_PASSWORD (web interface password)")
        print("      - JERA_SOFT_API_KEY (CoreAPI token)")

# Example usage:
if __name__ == "__main__":
    print("🧪 JeraSoft Rate Upload Demo")
    print("=" * 60)
    
    # Check configuration
    config_ok = True
    required_vars = {
        "JERA_SOFT_API_KEY": J_API_KEY,
        "JERASOFT_WEB_LOGIN": J_WEB_LOGIN, 
        "JERASOFT_WEB_PASSWORD": J_WEB_PASSWORD
    }
    
    for var_name, var_value in required_vars.items():
        if not var_value:
            print(f"❌ Missing: {var_name}")
            config_ok = False
        else:
            print(f"✅ {var_name}: {'*' * (len(var_value)-5) + var_value[-5:]}")
    
    if config_ok:
        print("\n🎯 Configuration looks good! Running examples...")
        example_usage()
    else:
        print("\n⚠️  Please set missing environment variables before running.")
        print("    See .env file or set them in your system environment.")