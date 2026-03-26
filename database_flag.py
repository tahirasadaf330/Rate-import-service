from typing import Iterable, Tuple, Optional, Dict, Any, List, Mapping
from pathlib import Path
import json
from datetime import date, datetime, timezone, timedelta
from database import (
    mark_ingest_processed,
    upsert_processing_status,
    get_processing_status,
    get_conn,
    REPROCESS_QUEUED_STATUS,
)
from date_verification import _parse_iso_utc_dt



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

def _folder_is_today_or_newer(meta: dict) -> bool:
    """
    Return True if the email's received date in metadata is today or newer.
    Falls back to date_utc if receivedDateTime_raw is missing.
    If both are missing/unparseable, we allow (return True).
    """
    cutoff_date = date.today() - timedelta(days=2)
    raw = (meta.get("receivedDateTime_raw") or "").strip()
    if raw:
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00")).date() >= cutoff_date
        except Exception:
            pass
    ds = (meta.get("date_utc") or "").strip()
    if len(ds) == 10:
        try:
            return datetime.fromisoformat(ds).date() >= cutoff_date
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
                # Preserve the current DB flag on existing rows. Passing False here
                # would overwrite a user-triggered reprocess request during seeding.
                is_reprocessing_enabled=None,
            )
            upserts += 1
            print(f"[STATUS] ensured processing_statuses id={_id} for {d.name}")

        except Exception as e:
            print(f"[STATUS][ERR] upsert failed for {d.name}: {e}")
            bad += 1

    print(f"[STATUS] seed summary: scanned={scanned}, upserted={upserts}, bad={bad}")
    return scanned, upserts, bad


def reset_processing_flags(directory_name: str, attachments_root: str = "attachments") -> bool:
    """
    Reset all processing status flags to FALSE to start fresh processing.
    Preserves jera_fetched only if it should be kept based on metadata.
    
    Args:
        directory_name: Name of the directory to reset processing flags for
        attachments_root: Root directory for attachments (default: "attachments")
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Get current status to check jera_fetched
        current_status = get_processing_status(directory_name=directory_name)
        if not current_status:
            print(f"  No processing status found for {directory_name}")
            return False
        
        # Check if we should preserve jera_fetched from metadata
        directory_path = Path(attachments_root) / directory_name
        metadata_path = directory_path / "metadata.json"
        preserve_jera = False
        
        if metadata_path.exists():
            try:
                with metadata_path.open("r", encoding="utf-8") as f:
                    metadata = json.load(f)
                preserve_jera = metadata.get("jera_fetched", False)
            except Exception:
                pass
        
        # Reset all flags to start fresh processing
        sql = """
            UPDATE processing_statuses 
            SET is_date_format_fetched = FALSE,
                is_jera_fetched = %s,
                is_file_cleaned = FALSE,
                is_rate_compared = FALSE, 
                is_rate_uploaded = FALSE,
                status = %s,
                updated_at = NOW()
            WHERE directory_name = %s
        """
        
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql, (preserve_jera, REPROCESS_QUEUED_STATUS, directory_name))
            affected = cur.rowcount
            conn.commit()
        
        if affected > 0:
            jera_status = "preserved" if preserve_jera else "reset"
            print(f"  Reset processing flags (jera_fetched {jera_status}) for {directory_name}")
            return True
        else:
            print(f"  No rows updated for {directory_name}")
            return False
            
    except Exception as e:
        print(f"  Error resetting processing flags for {directory_name}: {e}")
        return False
