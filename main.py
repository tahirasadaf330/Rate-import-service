""" 
Step 1: run the email verification script to fetch all the new valid files which will be stored in the attachments folder
Step 2: then run the jerasoft script to fetch all the relevant tables for comparision
Step 3: run the preprocess script on all the files 
Step 4: run the ratesheet comparision script to generate the comparision report
Step 5: run the database script to push the comparision results to the database 

"""
"""
while saving the files from email also create a meta data file for that directory that should include subject, sender, date, time
path to the directory.

then read that meta data file and create all the comparision files using jerasoft and save in the same directory.

then preprocess all the files 
"""
from multithreading import cleaned_out_path, find_jerasoft_file, vendor_files,as_of_from_metadata, read_comparison_table, df_to_detail_dicts,compute_upload_stats, parse_received_at,load_metadata, save_metadata
from multithreading import run_pipeline_mt
from jerasoft import export_rates_by_query
from ratesheet_comparision_engine import read_table, compare, write_excel
from email_verification import verify_fetch_emails
import os
import json
from pathlib import Path
from preprocess_data import load_clean_rates
from typing import Iterable, Tuple, Optional, Dict, Any, List, Mapping
from database import insert_rate_upload, bulk_insert_rate_upload_details, push_failed_emails_json_to_db, fetch_approved_unprocessed_paths_map, insert_rejected_email_row, insert_or_update_ingest_file, mark_ingest_processed,upsert_processing_status,mark_processing_stage
import pandas as pd
from datetime import date, datetime, timezone

import re


FAILED_EMAILS_PATH = Path(__file__).with_name("failed_emails.json")


#_____________ Email Verification Script_____________

# after = "2025-09-29"              # only include emails on/after this date (YYYY-MM-DD) or None     "2025-08-29"
after = datetime.now().strftime("%Y-%m-%d")
before = None       # only include emails on/before this date (YYYY-MM-DD) or None
unread_only = False    
ATTEMPTS = 2

#________________ Manual Date Verification _____________________

MAX_PREVIEW_ROWS = 200
EXCEL_EXTS = ('.xlsx', '.xlsm', '.xls')

def _parse_iso_utc_dt(s: Optional[str]) -> Optional[datetime]:
    if not s or not isinstance(s, str):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None

def _df_preview_records(df: pd.DataFrame, limit: int = MAX_PREVIEW_ROWS) -> list[dict]:
    """Rename columns to col1..colN and return first N rows as list of dicts (strings, blanks for NaN)."""
    if df is None or df.empty:
        return []
    df = df.head(limit).reset_index(drop=True)
    # ensure all values are strings; keep blanks for NaN
    try:
        df = df.astype("string")
    except Exception:
        df = df.astype(str)
    ncols = int(df.shape[1])
    df.columns = [f"col{i+1}" for i in range(ncols)]
    # fillna("") for string dtype keeps <NA> out of JSON
    try:
        df = df.fillna("")
    except Exception:
        pass
    return df.to_dict(orient="records")

def _first_attachment_path(meta: dict) -> Optional[Path]:
    """Return absolute Path to the first attachment (join with directory if needed)."""
    attachments = meta.get("attachments") or []
    if not isinstance(attachments, list) or len(attachments) == 0:
        return None

    first_item = str(attachments[0]).strip()
    if not first_item:
        return None

    p = Path(first_item)
    if p.is_absolute():
        return p

    # directory contains the folder path (Linux server)
    base_dir = str(meta.get("directory") or "").strip()
    if not base_dir:
        return None
    return Path(base_dir) / first_item

#############################################
# Helpers for verifying the date format


def _read_excel_native(path: str, sheet=0) -> pd.DataFrame:
    """
    Read an Excel sheet WITHOUT dtype=str so real Excel date cells
    stay as datetime/date/Timestamp. Tries calamine first, then openpyxl.
    """
    last_err = None
    for eng in ("calamine", "openpyxl"):
        try:
            return pd.read_excel(path, sheet_name=sheet, header=None, engine=eng)
        except Exception as e:
            last_err = e
            continue
    raise last_err or RuntimeError("Failed reading excel with calamine/openpyxl")

def _has_native_datetimes(df: pd.DataFrame, min_hits: int = 5) -> bool:
    """
    True if the grid appears to contain native datetime/date cells
    (either a datetime64 column or at least `min_hits` datetime-like objects).
    """
    # fast path: any datetime64 dtype column
    if any(pd.api.types.is_datetime64_any_dtype(t) for t in df.dtypes):
        return True

    # slower scan: mixed-type cells
    hits = 0
    for v in df.to_numpy().ravel():
        if isinstance(v, (datetime, date, pd.Timestamp)):
            hits += 1
            if hits >= min_hits:
                return True
    return False

###############################################################

MAX_PREVIEW_ROWS = 200
EXCEL_EXTS = ('.xlsx', '.xlsm', '.xls')

def ingest_files_for_manual_date(attachments_root: str | Path = "attachments") -> tuple[int, int, int]:
    """
    Walk attachments/* folders, and for each folder whose metadata.json has
    no 'date_verification_ingestion' (or it's False):
      - pick the first attachment
      - build preview_cache (first 200 rows, stringified)
      - INSERT/UPSERT into ingest_files
      - mark metadata['date_verification_ingestion'] = True

    NEW:
      - For Excel files, if the preview read shows native datetime/date cells,
        we auto-approve and set:
          DB:  status=True, date_format="YYYY-MM-DD",
               approved_at=now(UTC), is_format_auto_detected=True
          metadata: date_verification_ingestion_status=True,
                    date_format_identified="YYYY-MM-DD"
    Returns: (scanned_folders, inserted_rows, skipped_folders)
    """
    root = Path(attachments_root).expanduser().resolve()
    if not root.exists():
        print(f"[INGEST] attachments root not found: {root}")
        return (0, 0, 0)

    scanned = 0
    inserted = 0
    skipped = 0

    for folder in sorted(root.iterdir()):
        if not folder.is_dir():
            continue
        scanned += 1

        meta_path = folder / "metadata.json"
        if not meta_path.exists():
            print(f"[INGEST][SKIP] {folder.name}: no metadata.json")
            skipped += 1
            continue

        try:
            with meta_path.open("r", encoding="utf-8") as f:
                meta = json.load(f) or {}
        except Exception as e:
            print(f"[INGEST][SKIP] {folder.name}: failed reading metadata.json: {e}")
            skipped += 1
            continue

        # Skip if already ingested for manual date verification
        if bool(meta.get("date_verification_ingestion")):
            continue

        # Step 2: find first attachment
        fpath = _first_attachment_path(meta)
        if not fpath:
            print(f"[INGEST][SKIP] {folder.name}: metadata.attachments missing/empty")
            skipped += 1
            continue

        # Build the preview_cache
        preview_cache: list[dict] = []
        error_message: Optional[str] = None
        autodetected: bool = False  # only possible for Excel

        try:
            ext = fpath.suffix.lower()

            if ext == ".csv":
                # CSV: raw grid, all strings; manual format later
                df = pd.read_csv(str(fpath), header=None, nrows=MAX_PREVIEW_ROWS, dtype=str, on_bad_lines="skip")
                preview_cache = _df_preview_records(df, MAX_PREVIEW_ROWS)

            elif ext in EXCEL_EXTS:
                # Read natively (no dtype=str) to detect Excel-native date cells
                df_native = _read_excel_native(str(fpath), sheet=0)
                autodetected = _has_native_datetimes(df_native)
                # Stringified preview for UI
                preview_cache = _df_preview_records(df_native, MAX_PREVIEW_ROWS)

            else:
                # best-effort text read as CSV
                try:
                    df = pd.read_csv(str(fpath), header=None, nrows=MAX_PREVIEW_ROWS, dtype=str, engine="python")
                    preview_cache = _df_preview_records(df, MAX_PREVIEW_ROWS)
                except Exception as e2:
                    raise RuntimeError(f"Unsupported file type {ext} and CSV fallback failed: {e2}") from e2

        except Exception as e:
            error_message = f"preview build failed: {e}"
            preview_cache = []  # still ingest a row with the error
            autodetected = False

        # Collect DB fields
        email_address = (meta.get("sender") or "").strip() or None
        subject = (meta.get("subject") or "").strip() or None
        received_at = _parse_iso_utc_dt(meta.get("receivedDateTime_raw"))
        processed_at = _parse_iso_utc_dt(meta.get("processed_at_utc"))
        file_path = str(fpath)

        # NEW: optional DB flags when auto-detected for Excel
        upsert_kwargs = {}
        if ext in EXCEL_EXTS and autodetected:
            upsert_kwargs.update({
                "status": "approved",
                "date_format": "YYYY-MM-DD",
                "approved_at": datetime.now(timezone.utc),
                "is_format_auto_detected": True,
            })

        try:
            _id = insert_or_update_ingest_file(
                email_address=email_address,
                subject=subject,
                received_at=received_at,
                processed_at=processed_at,
                file_path=file_path,
                preview_cache=preview_cache,
                error_message=error_message,
                **upsert_kwargs,  # only applies if autodetected True
            )
            inserted += 1
            print(f"[INGEST][OK] id={_id} -> {folder.name} :: {fpath.name}")

            # Always mark the ingestion pass as done
            meta["date_verification_ingestion"] = True

            # If we auto-detected native dates for Excel, mark folder approved too
                        # If we auto-detected native dates for Excel, mark folder approved too
            if ext in EXCEL_EXTS and autodetected:
                meta["date_verification_ingestion_status"] = True
                meta["date_format_identified"] = "YYYY-MM-DD"
                try:
                    mark_processing_stage(directory_name=folder.name, stage="date_format_fetched")
                except Exception as e:
                    print(f"[STATUS][WARN] failed to mark date_format_fetched for {folder.name}: {e}")

            with meta_path.open("w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)


        except Exception as db_e:
            print(f"[INGEST][ERR] DB insert failed for {folder.name}: {db_e}")
            # do NOT set the flag so we can retry later

    print(f"[INGEST] summary: scanned={scanned}, inserted={inserted}, skipped={skipped}")
    return (scanned, inserted, skipped)



def mark_date_verification_ingestion(path_to_format: Mapping[str, Optional[str]]) -> Tuple[int, int, int]:
    """
    For each file path like:
      /.../attachments/<dir>/<filename.xlsx>
    find its parent directory <dir>, open <dir>/metadata.json, and set:
      date_verification_ingestion_status = True
      date_format_identified = <value from path_to_format[file_path]>

    - Input: dict mapping file_path -> date_format
    - Does NOT modify the input.
    - Deduplicates parent directories (first occurrence wins).
    - Returns (dirs_seen, updated, missing_meta).
    """
    # Collect unique parent directories in order, tracking chosen date_format per dir
    parents = []
    seen = set()
    dir_fmt = {}
    for p, fmt in path_to_format.items():
        try:
            d = Path(p).parent.resolve()
        except Exception:
            continue
        if d not in seen:
            seen.add(d)
            parents.append(d)
            dir_fmt[d] = fmt  # remember the first date_format seen for this dir

    dirs_seen = len(parents)
    updated = 0
    missing_meta = 0

    for d in parents:
        meta_path = d / "metadata.json"
        if not meta_path.exists():
            print(f"[INGEST][SKIP] No metadata.json in {d}")
            missing_meta += 1
            continue

        try:
            with meta_path.open("r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception as e:
            print(f"[INGEST][WARN] Failed to read {meta_path}: {e}")
            missing_meta += 1
            continue

        # Keep original behavior: if already true, do not modify file
        if meta.get("date_verification_ingestion_status") is True:
            print(f"[INGEST] Already marked true: {d}")
            continue

        meta["date_verification_ingestion_status"] = True
        meta["date_format_identified"] = dir_fmt.get(d)
        try:
            mark_processing_stage(directory_name=d.name, stage="date_format_fetched")
        except Exception as e:
            print(f"[STATUS][WARN] failed to mark date_format_fetched for {d.name}: {e}")

        # Atomic-ish write
        tmp = meta_path.with_suffix(meta_path.suffix + ".tmp")

        with tmp.open("w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        os.replace(tmp, meta_path)

        updated += 1
        print(f"[INGEST] Marked date_verification_ingestion_status=true (date_format_identified={dir_fmt.get(d)!r}) → {meta_path}")

    print(f"[INGEST] summary: dirs_seen={dirs_seen}, updated={updated}, missing_meta={missing_meta}")
    return dirs_seen, updated, missing_meta


#_______________ Preprocess Script _____________

ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}



def iter_preprocessed_dirs_(attachments_root: Path):
    """
    Yield directories under attachments_root whose metadata.json has "jerasoft_preprocessed": true.
    """
    for child in sorted(attachments_root.iterdir()):
        if not child.is_dir():
            continue
        meta = child / "metadata.json"
        if not meta.exists():
            continue
        try:
            with meta.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            print('  ✖ Failed to load metadata')
            continue

        # NEW: require manual date verification approval
        if not bool(data.get("date_verification_ingestion_status")):
            print(f"[SKIP] Awaiting date verification approval: {child}")
            continue

        if bool(data.get("jerasoft_preprocessed")) is True:
            results = data.get("preprocessed_results", {})
            # if not results or any(v is False for v in results.values()):
            if not results:
                yield child


def files_to_clean(folder: Path):
    """
    Return all cleanable files in folder (CSV/XLS/XLSX/XLSM), excluding metadata.json.
    """
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.name.lower() == "metadata.json":
            continue
        if f.name.startswith("~$") or f.name.startswith("."):
             continue
        if f.suffix.lower() in ALLOWED_EXTS:
            yield f


#________________ Ratesheet Comparision Script _____________

ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}

def iter_preprocessed_dirs(attachments_root: Path) -> Iterable[Path]:
    """
    Yield folders that finished JeraSoft and still need comparison:
    - no comparision_result yet, or
    - comparision_result exists but at least one vendor flag is not True.
    """
    for child in sorted(attachments_root.iterdir()):
        if not child.is_dir():
            continue
        meta_path = child / "metadata.json"
        if not meta_path.exists():
            continue
        try:
            with meta_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue

        # NEW: require manual date verification approval
        if not bool(data.get("date_verification_ingestion_status")):
            print(f"[SKIP] Awaiting date verification approval: {child}")
            continue

        # if data.get("jerasoft_preprocessed") is True:
        #     comp = data.get("comparision_result") or {}
        #     per_vendor = {k: v for k, v in comp.items() if k != "result"}
        #     # run if no result yet, or any vendor isn’t exactly True
        #     if not comp or any(v is not True for v in per_vendor.values()):
        #         yield child
        
        if data.get("jerasoft_preprocessed") is True:
            comp = data.get("comparision_result") or {}
            # run only if 'result' key is not present
            if "result" not in comp:
                yield child







def _read_metadata(folder: Path) -> dict:
    with (folder / "metadata.json").open("r", encoding="utf-8") as f:
        return json.load(f)

def _write_metadata(folder: Path, data: dict) -> None:
    with (folder / "metadata.json").open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# ________________ Database Script _____________

ATTACHMENTS_ROOT = "attachments"
_ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}



# ------------ metadata helpers ------------

def _parse_iso_utc_safe(s: Optional[str]) -> Optional[datetime]:
    if not s or not isinstance(s, str):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None

def _atomic_write_json(path: Path, data: dict) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)




def comparison_result_ok(meta: Dict[str, Any]) -> bool:
    """
    True iff metadata has a result 'ok' under either:
      meta['comparison_result']['result']  or  meta['comparision_result']['result']
    """
    d = meta.get("comparison_result") or meta.get("comparision_result")
    if not isinstance(d, dict):
        return False
    return str(d.get("result", "")).strip().lower() == "ok"


def mark_results_pushed(folder: Path, filename: str, status: Any) -> None:
    meta = load_metadata(folder) or {}
    rp = meta.get("results_pushed")
    if not isinstance(rp, dict):
        rp = {}
    rp[filename] = status
    meta["results_pushed"] = rp
    save_metadata(folder, meta)

# ------------ file discovery ------------

def find_result_files(folder: Path) -> List[Path]:
    """
    Return files whose stem **ends with** '_comparision_result' (case-insensitive)
    and have an allowed extension.
    """
    out: List[Path] = []
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in _ALLOWED_EXTS:
            continue
        if f.stem.lower().endswith("_comparision_result"):
            out.append(f)
    return out


def finalize_processed_flags(paths_map: Dict[str, Optional[str]]) -> Tuple[int, int, int]:
    """
    Given {file_path: date_format or None}, look up each file's parent folder.
    If that folder's metadata.json has results_pushed with all values strictly True,
    mark that specific file_path row as is_processed = TRUE in the DB.

    Returns: (dirs_scanned, eligible_dirs, rows_marked)
    """
    # Deduplicate by directory, but keep a mapping back to at least one file_path per dir
    dir_to_any_fp: Dict[Path, str] = {}
    for fp in paths_map.keys():
        try:
            d = Path(fp).parent.resolve()
        except Exception:
            continue
        # keep the first seen file_path for this dir
        dir_to_any_fp.setdefault(d, fp)

    dirs_scanned = 0
    eligible_dirs = 0
    to_mark: list[str] = []

    for d, fp_for_dir in dir_to_any_fp.items():
        dirs_scanned += 1
        meta_path = d / "metadata.json"
        if not meta_path.exists():
            print(f"[INGEST][SKIP] No metadata.json in {d}")
            continue

        try:
            with meta_path.open("r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception as e:
            print(f"[INGEST][WARN] Failed to read {meta_path}: {e}")
            continue

        rp = meta.get("results_pushed")
        if not isinstance(rp, dict) or not rp:
            print(f"[INGEST][SKIP] No results_pushed or empty in {meta_path}")
            continue

        # All entries must be strictly boolean True
        all_true = all(v is True for v in rp.values())
        if not all_true:
            print(f"[INGEST][WAIT] Not all results pushed in {meta_path} -> {rp}")
            continue

        # This directory is eligible -> mark its chosen file_path
        eligible_dirs += 1
        to_mark.append(fp_for_dir)

    rows_marked = 0
    if to_mark:
        try:
            rows_marked = mark_ingest_processed(to_mark, processed=True)
            print(f"[INGEST] Marked is_processed=TRUE for {rows_marked} rows")
        except Exception as e:
            print(f"[INGEST][ERR] Failed to update is_processed flags: {e}")

    print(f"[INGEST] finalize_processed_flags summary: dirs_scanned={dirs_scanned}, "
          f"eligible_dirs={eligible_dirs}, rows_marked={rows_marked}")
    return dirs_scanned, eligible_dirs, rows_marked
#____________________________________________________
from datetime import date, datetime

def _folder_is_today_or_newer(meta: dict) -> bool:
    """
    Return True if the email's received date in metadata is today or newer.
    Falls back to date_utc if receivedDateTime_raw is missing.
    If both are missing/unparseable, we allow (return True).
    """
    raw = (meta.get("receivedDateTime_raw") or "").strip()
    if raw:
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00")).date() >= date.today()
        except Exception:
            pass
    ds = (meta.get("date_utc") or "").strip()
    if len(ds) == 10:
        try:
            return datetime.fromisoformat(ds).date() >= date.today()
        except Exception:
            pass
    # If we can't tell, allow it (so we don't accidentally skip new data)
    return True


def seed_processing_status_rows(attachments_root: str | Path = "attachments") -> tuple[int, int, int]:
    """
    Walk attachments/*, read metadata.json, and ensure a processing_statuses row exists.
    Returns (folders_scanned, rows_upserted, missing_or_invalid_meta).
    """
    root = Path(attachments_root).expanduser().resolve()
    if not root.exists():
        print(f"[STATUS] attachments root not found: {root}")
        return (0, 0, 0)

    scanned = 0
    upserts = 0
    bad = 0

    for d in sorted(root.iterdir()):
        if not d.is_dir():
            continue
        scanned += 1
        meta_path = d / "metadata.json"
        if not meta_path.exists():
            bad += 1
            continue

        try:
            with meta_path.open("r", encoding="utf-8") as f:
                meta = json.load(f) or {}

            # ⬇️ filter out folders older than today
            if not _folder_is_today_or_newer(meta):
                print(f"[STATUS] skip {d.name}: older than today")
                continue

        except Exception as e:
            print(f"[STATUS][WARN] failed reading {meta_path}: {e}")
            bad += 1
            continue

        # derive fields safely
        try:
            internet_message_id = str(
                meta.get("internet_message_id")
                or meta.get("internetMessageId")
                or meta.get("message_id")
                or ""
            ).strip() or None

            directory_name = (Path(meta.get("directory") or d).name)

            sender_email = (meta.get("sender") or None)
            email_subject = (meta.get("subject") or None)

            # reuse your helper for ISO strings; fallback to None on error
            try:
                email_received_at = _parse_iso_utc_dt(meta.get("receivedDateTime_raw"))
            except Exception:
                email_received_at = None

            if not internet_message_id:
                # it's OK if you don't have it yet; we can still create the row keyed by message id later.
                # For now, skip if there is no message-id to keep your UNIQUE(internet_message_id) constraint happy.
                print(f"[STATUS] skip {d.name}: no internet_message_id in metadata yet")
                continue

            _id = upsert_processing_status(
                internet_message_id=internet_message_id,
                directory_name=directory_name,
                sender_email=sender_email,
                email_subject=email_subject,
                email_received_at=email_received_at,
            )
            upserts += 1
            print(f"[STATUS] ensured processing_statuses id={_id} for {d.name}")

        except Exception as e:
            print(f"[STATUS][ERR] upsert failed for {d.name}: {e}")
            bad += 1

    print(f"[STATUS] seed summary: scanned={scanned}, upserted={upserts}, bad={bad}")
    return scanned, upserts, bad


#______________________________________________________________________________


if __name__ == "__main__":
    # scrap all the valid emails
    verify_fetch_emails(after, before, unread_only)
    seed_processing_status_rows("attachments")

    #________________________________________________
    # run the ingestion pass before any further processing
    ingest_files_for_manual_date("attachments")

    valid_paths = fetch_approved_unprocessed_paths_map()

    mark_date_verification_ingestion(valid_paths)
    #________________________________________________


    # # fetch jerasoft rates
    # process_all_directories()

    # # preprocessing the files
    # clean_preprocessed_folders("attachments")

    # # running comparision engine on all the files
    # compare_preprocessed_folders("attachments", notice_days=7, rate_tol=0.0001)  #check the difference upto 4 decimal places.

    # # pushing all the relevant details to the data base
    # push_rejections_from_metadata("attachments")

    # push_all_ok_results(ATTACHMENTS_ROOT)

    push_failed_emails_json_to_db("failed_emails.json")  
    
    run_pipeline_mt("attachments", max_workers=3)


    #___________________________________________________________
    # Setting the is_processed status of the processed files to true in the db
    # Try to mark is_processed for those whose folders have fully-pushed results
    finalize_processed_flags(valid_paths)


