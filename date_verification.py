from database import (
    mark_processing_stage,
    insert_or_update_ingest_file,
    fetch_authorized_sender_date_format,
    upsert_authorized_sender_date_format,
)
from datetime import date, datetime, timezone
import pandas as pd
from pathlib import Path
import json
import os
from typing import Iterable, Tuple, Optional, Dict, Any, List, Mapping
MAX_PREVIEW_ROWS = 200
EXCEL_EXTS = ('.xlsx', '.xlsm', '.xls')
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

def _read_excel_native(path: str, sheet=0) -> pd.DataFrame:
    """
    Read an Excel sheet WITHOUT dtype=str so real Excel date cells
    stay as datetime/date/Timestamp. Prefer openpyxl first on Windows to avoid
    rare calamine panics, then fall back to calamine.
    """
    last_err = None
    for eng in ("openpyxl", "calamine"):
        try:
            return pd.read_excel(path, sheet_name=sheet, header=None, engine=eng)
        except Exception as e:
            last_err = e
            continue
        except BaseException as e:  # catch non-Exception panics from native libs
            last_err = e
            continue
    raise last_err or RuntimeError("Failed reading excel with openpyxl/calamine")

def _parse_iso_utc_dt(s: Optional[str]) -> Optional[datetime]:
    if not s or not isinstance(s, str):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None
    
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
        
        # DB-first: if sender already has a saved date_format, we will still upsert ingest_files
        # but mark it approved with that format.
        try:
            sender_email = (meta.get("sender") or "").strip() or None
        except Exception:
            sender_email = None
        sender_fmt = None
        try:
            sender_fmt = fetch_authorized_sender_date_format(sender_email)
        except Exception:
            sender_fmt = None

        # Build the preview_cache
        preview_cache: list[dict] = []
        error_message: Optional[str] = None
        autodetected: bool = False  # only possible for Excel and only used when DB format is missing

        try:
            ext = fpath.suffix.lower()

            if ext == ".csv":
                # CSV: raw grid, all strings; manual format later
                df = pd.read_csv(str(fpath), header=None, nrows=MAX_PREVIEW_ROWS, dtype=str, on_bad_lines="skip")
                preview_cache = _df_preview_records(df, MAX_PREVIEW_ROWS)

            elif ext in EXCEL_EXTS:
                # Read natively (no dtype=str) to detect Excel-native date cells
                df_native = _read_excel_native(str(fpath), sheet=0)
                # Only attempt autodetect when we don't already have a sender-level format
                autodetected = (not bool(sender_fmt)) and _has_native_datetimes(df_native)
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

        # Optional DB flags when:
        # - sender_fmt exists (DB-first) -> treat as approved using that format
        # - autodetected Excel native dates -> approve as YYYY-MM-DD
        upsert_kwargs = {}
        if sender_fmt:
            upsert_kwargs.update({
                "status": "approved",
                "date_format": sender_fmt,
                "approved_at": datetime.now(timezone.utc),
                "is_format_auto_detected": False,
            })
        elif ext in EXCEL_EXTS and autodetected:
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
            if sender_fmt:
                meta["date_verification_ingestion_status"] = True
                meta["date_format_identified"] = sender_fmt
                try:
                    mark_processing_stage(directory_name=folder.name, stage="date_format_fetched")
                except Exception as e:
                    print(f"[STATUS][WARN] failed to mark date_format_fetched for {folder.name}: {e}")
            elif ext in EXCEL_EXTS and autodetected:
                meta["date_verification_ingestion_status"] = True
                meta["date_format_identified"] = "YYYY-MM-DD"
                # Autodetected date format should be persisted per-sender so next emails can reuse it.
                try:
                    upsert_authorized_sender_date_format(
                        email=email_address,
                        date_format="YYYY-MM-DD",
                    )
                except Exception:
                    pass
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

        chosen_fmt = dir_fmt.get(d)
        meta["date_verification_ingestion_status"] = True
        meta["date_format_identified"] = chosen_fmt
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
