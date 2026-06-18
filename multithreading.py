from __future__ import annotations
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime, timezone
import pandas as pd
import traceback  # >>> ADDED

# ---- import your existing functions ----
from jerasoft import export_rates_by_query
from ratesheet_comparision_engine import read_table, compare, write_excel, summarize_changes
from preprocess_data import load_clean_rates
from database import (
    insert_rate_upload,
    bulk_insert_rate_upload_details,
    mark_processing_stage,
    fetch_active_vendor_header_mapping,
    fetch_vendor_date_format_by_sender_email,
    fetch_vendor_context_by_sender_email,
    vendor_has_pending_jera_upload_today,
    set_processing_status_text,
    WAITING_PREVIOUS_VENDOR_PENDING,
    WAITING_JERA_TABLE_FOR_SUBJECT,
    get_processing_status,
    get_or_create_invalid_subject,
    insert_invalid_subject_detail,
    find_invalid_subject_detail,
    find_approved_jera_table_for_subject,
)
from jerasoft import export_rates_by_query, get_table_id_by_name, fetch_active_current_future_rates, save_rates_to_excel

# Reuse constants from your code
ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}
EXPECTED_COLS = ["Code", "Old Rate", "New Rate", "Effective Date", "Status", "Change Type", "Notes"]

def _derive_prefix(meta: Dict[str, Any]) -> Optional[str]:
    """
    Derive a normalized prefix string from metadata.

    Order:
      1) meta["prefix"]
      2) meta["force_jerasoft_table_name"]
      3) meta["subject"]

    Normalization:
      - "none" -> "NONE"
      - digits preserved as string
      - returns None if not found
    """
    def _norm(val: object) -> Optional[str]:
        if val is None:
            return None
        s = str(val).strip()
        if not s:
            return None
        return "NONE" if s.casefold() == "none" else s

    direct = _norm(meta.get("prefix"))
    if direct:
        return direct

    try:
        for src in (meta.get("force_jerasoft_table_name"), meta.get("subject")):
            s = str(src or "")
            m = re.search(r"\bprefix\b\s*[:\s-]*\s*(none|\d+)", s, flags=re.IGNORECASE)
            if m:
                return "NONE" if m.group(1).strip().lower() == "none" else m.group(1).strip()
    except Exception:
        return None

    return None

def _derive_trunk(meta: Dict[str, Any]) -> Optional[str]:
    """
    Derive trunk for vendor/day waiting gate.

    Rules:
      1) valid subject -> trunk from subject parser
      2) invalid subject -> derive from force_jerasoft_table_name
    """
    try:
        from email_verification import validate_subject
    except Exception:
        validate_subject = None  # type: ignore[assignment]

    subject = str(meta.get("subject") or "").strip()
    if subject and validate_subject:
        try:
            parsed = validate_subject(subject)
        except Exception:
            parsed = None
        trunk = str((parsed or {}).get("trunk") or "").strip()
        if trunk:
            return trunk

    force_table = str(meta.get("force_jerasoft_table_name") or "").strip()
    if not force_table:
        return None

    if validate_subject:
        try:
            parsed = validate_subject(force_table)
        except Exception:
            parsed = None
        trunk = str((parsed or {}).get("trunk") or "").strip()
        if trunk:
            return trunk

    # Keyword fallback for invalid subjects/table names used by operations.
    # Pick the first keyword by position in the force-table string.
    keyword_patterns = [
        ("CC", r"\bcall\s+center\b"),
        ("CC", r"\bcc\b"),
        ("STD", r"\bstandard\b"),
        ("STD", r"\bgold\b"),
        ("STD", r"\bwholesale\b"),
        ("STD", r"\bstd\b"),
        ("PRM", r"\bprm\b"),
        ("PRM", r"\bprs\b"),
        ("PRM", r"\bpremium\b"),
        ("PRM", r"\bsilver\b"),
        ("ORTP", r"\bortp\b"),
        ("TDM", r"\btdm\b"),
        ("DID", r"\bdid\b"),
        ("SPECIAL", r"\bspecial\b"),
        ("ATX", r"\batx\b"),
    ]
    best: Optional[tuple[int, str]] = None
    for label, pat in keyword_patterns:
        m = re.search(pat, force_table, flags=re.IGNORECASE)
        if not m:
            continue
        pos = m.start()
        if best is None or pos < best[0]:
            best = (pos, label)
    if best:
        return best[1]

    m = re.search(r"\b([A-Za-z][\w\-]*)\s+trunk\b", force_table, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    return None

def _mark_subject_invalid_pending_jera(
    folder: Path,
    meta: Dict[str, Any],
    error_message: str,
) -> None:
    """
    Route a valid-but-ambiguous subject (JERA lookup returned 0 or >=2 tables)
    into invalid_subjects / invalid_subject_details so an admin can assign a
    JeraSoft table manually. Also flips processing_statuses to a waiting text
    and flags the folder's metadata so it short-circuits on future runs.
    """
    try:
        from email_verification import _strip_date_time_tokens_for_invalid_subject
    except Exception:
        _strip_date_time_tokens_for_invalid_subject = None  # type: ignore[assignment]

    sender = str(meta.get("sender") or "").strip()
    subject = str(meta.get("subject") or "").strip()
    subject_key = subject
    if _strip_date_time_tokens_for_invalid_subject:
        try:
            subject_key = _strip_date_time_tokens_for_invalid_subject(subject)
        except Exception:
            subject_key = subject

    try:
        rcvd_dt = None
        raw = meta.get("receivedDateTime_raw")
        if isinstance(raw, str) and raw.strip():
            try:
                rcvd_dt = datetime.fromisoformat(raw.strip().replace("Z", "+00:00"))
            except Exception:
                rcvd_dt = None

        parent_id = get_or_create_invalid_subject(
            email=sender or "",
            received_at=rcvd_dt,
            processed_at=datetime.now(timezone.utc),
            status="pending",
        )
        found = find_invalid_subject_detail(parent_id, subject_key)
        if not found:
            insert_invalid_subject_detail(parent_id, subject_key, None)
    except Exception as e:
        print(f"[{folder.name}] warn: failed to record invalid_subject for JERA ambiguity: {e}")

    meta["jera_fetched"] = False
    meta["keyword_error"] = error_message
    meta["waiting_for_jera_table"] = True
    save_metadata(folder, meta)

    try:
        set_processing_status_text(
            directory_name=folder.name,
            status_text=WAITING_JERA_TABLE_FOR_SUBJECT,
        )
    except Exception as e:
        print(f"[{folder.name}] warn: failed to set waiting status: {e}")


def load_metadata(folder: Path) -> Optional[Dict[str, Any]]:
    meta = folder / "metadata.json"
    if not meta.exists():
        return None
    try:
        with meta.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def save_metadata(folder: Path, data: Dict[str, Any]) -> None:
    # atomic write to avoid corrupting metadata.json
    path = folder / "metadata.json"
    import tempfile, os
    with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8", dir=folder) as tmp:
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.flush(); os.fsync(tmp.fileno())
        tmpname = tmp.name
    os.replace(tmpname, path)

# def cleaned_out_path(p: Path) -> Path:
#     """Return <stem>_cleaned<suffix> in the same folder."""
#     p = Path(p)
#     return p.with_name(f"{p.stem}_cleaned{p.suffix}")

def cleaned_out_path(p: Path) -> Path:
    """Return <stem>_cleaned.xlsx in the same folder (force xlsx)."""
    p = Path(p)
    base = p.with_suffix(".xlsx")
    return base.with_name(f"{base.stem}_cleaned.xlsx")



def find_jerasoft_file(folder: Path) -> Optional[Path]:
    """Prefer jerasoft_comparison_all.xlsx, else first *_jerasoft_comparison.xlsx."""
    prime = folder / "jerasoft_comparison_all_cleaned.xlsx"
    if prime.exists():
        return prime
    candidates = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xls")
        and p.name.lower().endswith("_jerasoft_comparison_cleaned.xlsx")
    )
    return candidates[0] if candidates else None

def vendor_files(folder: Path) -> list[Path]:
    """
    Return vendor files to compare, preferring *_cleaned.* when both exist.
    Excludes metadata.json and any JeraSoft comparison outputs (raw or cleaned).
    """
    def is_jerasoft(p: Path) -> bool:
        n = p.name.lower()
        return (
            n == "jerasoft_comparison_all.xlsx"
            or n == "jerasoft_comparison_all_cleaned.xlsx"
            or n.endswith("_jerasoft_comparison.xlsx")
            or n.endswith("_jerasoft_comparison_cleaned.xlsx")
        )

    # collect candidates
    candidates: list[Path] = []
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.name.lower() == "metadata.json":
            continue
        if f.suffix.lower() not in ALLOWED_EXTS:
            continue
        if is_jerasoft(f):
            continue
        # if not f.stem.lower().endswith("_cleaned"):
        #     continue  # ❌ skip non-cleaned vendor files
        candidates.append(f)

    # prefer *_cleaned over raw twin
    by_base: dict[str, dict[str, Path]] = {}
    for f in candidates:
        stem = f.stem
        is_cleaned = stem.endswith("_cleaned")
        base_stem = stem[:-8] if is_cleaned else stem  # strip "_cleaned"
        key = f"{base_stem}{f.suffix.lower()}"         # base name + ext

        entry = by_base.setdefault(key, {})
        if is_cleaned:
            entry["cleaned"] = f
        else:
            entry["raw"] = f

    chosen: list[Path] = []
    for key, pair in by_base.items():
        chosen.append(pair.get("cleaned") or pair.get("raw"))

    return sorted(chosen)

def as_of_from_metadata(folder: Path) -> str:
    """Use metadata.date_utc if available, else today (UTC, YYYY-MM-DD)."""
    meta = folder / "metadata.json"
    if meta.exists():
        try:
            with meta.open("r", encoding="utf-8") as f:
                data = json.load(f)
            d = (data.get("date_utc") or "").strip()
            if len(d) == 10 and d[4] == "-" and d[7] == "-":
                return d
        except Exception:
            pass
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def read_comparison_table(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}")

    # robust column rename
    rename_map: Dict[str, str] = {}
    cols_norm = {c: " ".join(str(c).strip().split()).lower() for c in df.columns}
    for c, n in cols_norm.items():
        if n == "code": rename_map[c] = "Code"
        elif n == "old rate": rename_map[c] = "Old Rate"
        elif n == "new rate": rename_map[c] = "New Rate"
        elif n == "effective date": rename_map[c] = "Effective Date"
        elif n == "status": rename_map[c] = "Status"
        elif n == "change type": rename_map[c] = "Change Type"
        elif n == "notes": rename_map[c] = "Notes"
        elif n == "old billing increment": rename_map[c] = "Old Billing Increment"
        elif n == "new billing increment": rename_map[c] = "New Billing Increment"
        elif n == "code name": rename_map[c] = "Dst Code Name"
        elif n == "dst code name": rename_map[c] = "Dst Code Name"
    df = df.rename(columns=rename_map)

    # ensure required columns exist
    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name}: missing expected columns {missing}. Found: {list(df.columns)}")

    # Clean types but DO NOT drop optional columns
    df["Code"] = df["Code"].astype(str).str.strip()
    df = df[df["Code"].ne("")]  # drop empty codes
    df["Old Rate"] = pd.to_numeric(df["Old Rate"], errors="coerce")
    df["New Rate"] = pd.to_numeric(df["New Rate"], errors="coerce")
    df["Effective Date"] = pd.to_datetime(df["Effective Date"], errors="coerce", utc=True)
    df.dropna(how="all", inplace=True)
    return df


def build_full_new_comparison_from_vendor(df_vendor: pd.DataFrame) -> pd.DataFrame:
    """Construct a synthetic comparison result when there is *no* Jera baseline.

    Each vendor row becomes a "New" row with Status="Accepted" and no Old Rate.
    This lets the rest of the pipeline (DB push, Jera upload) work unchanged,
    because it still sees a normal comparison-result style table.
    """
    rows: list[dict[str, Any]] = []
    for _, r in df_vendor.iterrows():
        code = str(r.get("Dst Code") or "").strip()
        if not code:
            continue

        name = r.get("Dst Code Name")
        bi_new = r.get("Billing Increment")

        rows.append(
            {
                "Code": code,
                "Dst Code Name": None if pd.isna(name) or str(name).strip() == "" else str(name).strip(),
                "Old Rate": None,
                "New Rate": r.get("Rate"),
                "Old Billing Increment": None,
                "New Billing Increment": None if pd.isna(bi_new) else (str(bi_new).strip() or None),
                "Effective Date": r.get("Effective Date"),
                "Status": "Accepted",
                "Change Type": "New",
                "Notes": "imported as new (no baseline ratesheet)",
            }
        )

    if not rows:
        return pd.DataFrame(
            columns=[
                "Code",
                "Dst Code Name",
                "Old Rate",
                "New Rate",
                "Old Billing Increment",
                "New Billing Increment",
                "Effective Date",
                "Status",
                "Change Type",
                "Notes",
            ]
        )

    return pd.DataFrame(rows)

def df_to_detail_dicts(df: pd.DataFrame, received_at: Optional[datetime] = None) -> List[Dict[str, Any]]:
    details: List[Dict[str, Any]] = []

    has_old_bi    = "Old Billing Increment" in df.columns
    has_new_bi    = "New Billing Increment" in df.columns
    has_code_name = "Dst Code Name" in df.columns  # <-- align with rename

    for _, r in df.iterrows():
        eff = r["Effective Date"]
        eff_py = None if pd.isna(eff) else eff.to_pydatetime()  # tz-aware UTC

        change_type = None if pd.isna(r["Change Type"]) else str(r["Change Type"]).strip()
        # Ensure Closed rows always have an effective_date for UI/DB:
        # if missing in the comparison sheet, use the email received_at (DATE ONLY).
        # Store as UTC midnight so DB/UI show 'YYYY-MM-DD'.
        if eff_py is None and received_at is not None and isinstance(change_type, str) and change_type.strip().lower() == "closed":
            try:
                ra_utc = received_at.astimezone(timezone.utc)
                eff_py = datetime(ra_utc.year, ra_utc.month, ra_utc.day, tzinfo=timezone.utc)
            except Exception:
                eff_py = received_at

        item: Dict[str, Any] = {
            "dst_code": None if pd.isna(r["Code"]) else str(r["Code"]).strip(),
            "rate_existing": None if pd.isna(r["Old Rate"]) else float(r["Old Rate"]),
            "rate_new": None if pd.isna(r["New Rate"]) else float(r["New Rate"]),
            "effective_date": eff_py,
            "change_type": change_type,
            "status": None if pd.isna(r["Status"]) else str(r["Status"]).strip(),
            "notes": None if pd.isna(r["Notes"]) else str(r["Notes"]).strip(),
        }

        if has_old_bi:
            v = r.get("Old Billing Increment")
            item["old_billing_increment"] = None if pd.isna(v) else str(v).strip()
        if has_new_bi:
            v = r.get("New Billing Increment")
            item["new_billing_increment"] = None if pd.isna(v) else str(v).strip()
        if has_code_name:
            v = r.get("Dst Code Name")
            item["code_name"] = None if pd.isna(v) else str(v).strip()

        details.append(item)

    return details

BOUND = r"(?:(?<=^)|(?<=,))\s*{label}\s*(?:(?=,)|(?=$))"  # comma-boundary regex

def compute_upload_stats(dfs: List[pd.DataFrame]) -> Dict[str, int]:
    if not dfs:
        return {
            "total_rows": 0,
            "new": 0, "increase": 0, "decrease": 0, "unchanged": 0, "closed": 0,
            "stashed": 0,
            "backdated_increase": 0, "backdated_decrease": 0,
            "billing_increment_changes": 0,
        }

    df = pd.concat(dfs, ignore_index=True)

    # Membership by Change Type (supports multi-label like "Billing ... Changes,Backdated Increase")
    is_new      = _has_ct(df, "New")
    is_closed   = _has_ct(df, "Closed")
    is_unchanged= _has_ct(df, "Unchanged")
    is_stashed  = _has_ct(df, "Stashed")

    is_back_inc = _has_ct(df, "Backdated Increase")
    is_back_dec = _has_ct(df, "Backdated Decrease")

    # Normal inc/dec exclude backdated so totals don’t double-count
    is_inc = _has_ct(df, "Increase") & ~is_back_inc
    is_dec = _has_ct(df, "Decrease") & ~is_back_dec

    # Billing increment changes: count only rows whose Change Type includes
    # the "Billing Increments Changes" label. This avoids treating Closed
    # rows (which often have Old BI set and New BI as NaN) as BI changes.
    bic = int(_has_ct(df, "Billing Increments Changes").sum())

    return {
        "total_rows": int(len(df)),
        "new":               int(is_new.sum()),
        "increase":          int(is_inc.sum()),
        "decrease":          int(is_dec.sum()),
        "unchanged":         int(is_unchanged.sum()),
        "closed":            int(is_closed.sum()),
        "stashed":           int(is_stashed.sum()),
        "backdated_increase":int(is_back_inc.sum()),
        "backdated_decrease":int(is_back_dec.sum()),
        "billing_increment_changes": bic,
    }

def parse_received_at(meta: Dict[str, Any]) -> Optional[datetime]:
    raw = meta.get("receivedDateTime_raw")
    if isinstance(raw, str) and raw.strip():
        try:
            s = raw.strip().replace("Z", "+00:00")
            return datetime.fromisoformat(s).astimezone(timezone.utc)
        except Exception:
            pass
    date_s = meta.get("date_utc")
    time_s = meta.get("time_utc")
    if isinstance(date_s, str) and date_s.strip():
        try:
            ts = f"{date_s.strip()}T{(time_s or '00:00:00').strip()}"
            dt = datetime.fromisoformat(ts)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            else:
                dt = dt.astimezone(timezone.utc)
            return dt
        except Exception:
            return None
    return None

# def load_metadata(folder: Path) -> Optional[Dict[str, Any]]:
#     meta = folder / "metadata.json"
#     if not meta.exists():
#         return None
#     try:
#         with meta.open("r", encoding="utf-8") as f:
#             return json.load(f)
#     except Exception:
#         return None
    
#     def save_metadata(folder: Path, data: Dict[str, Any]) -> None:
#     # atomic write to avoid corrupting metadata.json
#      path = folder / "metadata.json"
#     import tempfile, os
#     with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8", dir=folder) as tmp:
#         json.dump(data, tmp, ensure_ascii=False, indent=2)
#         tmp.flush(); os.fsync(tmp.fileno())
#         tmpname = tmp.name
#     os.replace(tmpname, path)
    
def _has_ct(df: pd.DataFrame, label: str) -> pd.Series:
    pat = re.compile(BOUND.format(label=re.escape(label)), flags=re.IGNORECASE)
    ct = df.get("Change Type", pd.Series([], dtype="object")).astype(str)
    return ct.str.contains(pat, na=False)


# ----------------- per-folder pipeline worker -----------------

def process_one_folder(folder: Path) -> str:
    """
    Run the full pipeline for a single attachments/<folder>.
    Returns a short log string.
    """
    meta_path = folder / "metadata.json"
    meta = load_metadata(folder)
    if not meta:
        return f"[{folder.name}] skip: no/invalid metadata.json"

    # DB guard: skip failed dirs unless reprocessing is enabled; skip successful dirs.
    # This prevents auto-retrying directories that already failed once.
    try:
        from database import get_processing_status

        db_status = get_processing_status(directory_name=folder.name)
        if db_status:
            status = str(db_status.get("status") or "").strip().lower()
            is_reprocessing_enabled = bool(db_status.get("is_reprocessing_enabled"))

            if status.startswith("failed") and not is_reprocessing_enabled:
                return f"[{folder.name}] skip: marked as failed (reprocessing disabled)"
            if status == "success":
                return f"[{folder.name}] already pushed"
    except Exception:
        # If DB check fails, keep going (don't block pipeline on DB issues).
        pass
    # DB is source-of-truth for date format, based on vendor:
    # sender email -> vendor_contacts -> vendors.date_format
    vendor_fmt = None
    is_partial_vendor = False
    vendor_ctx: Dict[str, Any] = {"vendor_id": None, "vendor_date_format": None, "is_partial": False, "is_mapped": False}
    active_vendor_header_mapping: Optional[Dict[str, Any]] = None
    vendor_mapping_error: Optional[str] = None
    try:
        vendor_ctx = fetch_vendor_context_by_sender_email(meta.get("sender"))
        vendor_fmt = vendor_ctx.get("vendor_date_format") or None
        is_partial_vendor = bool(vendor_ctx.get("is_partial"))
    except Exception:
        vendor_ctx = {"vendor_id": None, "vendor_date_format": None, "is_partial": False, "is_mapped": False}
        vendor_fmt = None
        is_partial_vendor = False

    if bool(vendor_ctx.get("is_mapped")):
        vendor_id = vendor_ctx.get("vendor_id")
        if vendor_id is None:
            vendor_mapping_error = "Vendor is marked mapped but vendor_id could not be resolved"
        else:
            try:
                active_vendor_header_mapping = fetch_active_vendor_header_mapping(vendor_id)
            except Exception as exc:
                vendor_mapping_error = f"Failed to fetch active vendor header mapping: {exc}"
            if active_vendor_header_mapping is None and vendor_mapping_error is None:
                vendor_mapping_error = f"No active vendor header mapping found for vendor_id {vendor_id}"

    if vendor_fmt:
        # Override metadata (even if it was auto-detected as YYYY-MM-DD)
        if (meta.get("date_format_identified") or "").strip() != vendor_fmt:
            meta["date_format_identified"] = vendor_fmt
        if not bool(meta.get("date_verification_ingestion_status")):
            meta["date_verification_ingestion_status"] = True
        if vendor_ctx.get("vendor_id") is not None:
            meta["vendor_id"] = vendor_ctx.get("vendor_id")
        save_metadata(folder, meta)
        try:
            mark_processing_stage(directory_name=folder.name, stage="date_format_fetched")
        except Exception:
            pass
    else:
        # No DB format: fall back to existing ingest/manual approval gate
        if not bool(meta.get("date_verification_ingestion_status")):
            return f"[{folder.name}] skip: waiting for date verification approval"

    # -------- Vendor/day gating: don't process a new sheet if a previous JeraSoft upload
    # for this vendor is still pending today. Persist this as a DB-visible "waiting" status.
    try:
        sender = str(meta.get("sender") or "").strip() or None
        trunk = _derive_trunk(meta)
        internet_message_id = str(
            meta.get("internet_message_id")
            or meta.get("internetMessageId")
            or meta.get("message_id")
            or ""
        ).strip() or None

        if vendor_has_pending_jera_upload_today(
            sender_email=sender,
            trunk=trunk,
            exclude_internet_message_id=internet_message_id,
        ):
            try:
                set_processing_status_text(
                    directory_name=folder.name,
                    status_text=WAITING_PREVIOUS_VENDOR_PENDING,
                )
            except Exception as e:
                # Status update failure should not crash the pipeline; still skip processing.
                print(f"[{folder.name}] warn: failed to set waiting status in DB: {e}")
            return f"[{folder.name}] skip: {WAITING_PREVIOUS_VENDOR_PENDING}"
    except Exception:
        # If the gating check fails (e.g., DB down), keep processing rather than blocking.
        pass

    # -------- 1) JeraSoft export (if needed) --------

    if not bool(meta.get("jerasoft_preprocessed")):
        company     = (meta.get("company") or "").strip()
        subject     = (meta.get("subject") or "").strip()
        prefix      = _derive_prefix(meta)
        trunk       = _derive_trunk(meta)
        dir_path    = meta.get("directory")
        attachments = meta.get("attachments", [])
        force_table = (meta.get("force_jerasoft_table_name") or "").strip()

        # Proactive check: before running the JERA lookup, see if this exact subject
        # already has an admin-approved table in invalid_subject_details. This covers
        # both (a) folders we previously flagged as waiting, and (b) brand-new emails
        # whose subject has been approved once before by the admin — no extra tick.
        if not force_table and subject:
            try:
                from email_verification import _strip_date_time_tokens_for_invalid_subject
                sender = str(meta.get("sender") or "").strip()
                subject_key = _strip_date_time_tokens_for_invalid_subject(subject)
                approved = find_approved_jera_table_for_subject(sender, subject_key)
                if approved:
                    force_table = approved.strip()
                    meta["force_jerasoft_table_name"] = force_table
                    meta["waiting_for_jera_table"] = False
                    save_metadata(folder, meta)
                    print(f"[{folder.name}] reused approved JERA table: {force_table}")
            except Exception as e:
                print(f"[{folder.name}] warn: approved-table lookup failed: {e}")

        # If we previously flagged this folder as waiting and no approval was found
        # in the proactive check above, keep waiting.
        if not force_table and bool(meta.get("waiting_for_jera_table")):
            try:
                set_processing_status_text(
                    directory_name=folder.name,
                    status_text=WAITING_JERA_TABLE_FOR_SUBJECT,
                )
            except Exception:
                pass
            return f"[{folder.name}] skip: {WAITING_JERA_TABLE_FOR_SUBJECT}"

        if not dir_path or not attachments:
            return f"[{folder.name}] skip: missing directory/attachments info"

        # choose output filename
        if len(attachments) == 1:
            base_name = Path(attachments[0]).stem
            output_file = f"{base_name}_jerasoft_comparison.xlsx"
        else:
            output_file = "jerasoft_comparison_all.xlsx"
        output_path = str(Path(dir_path) / output_file)

        try:
            if force_table:
                # Override path: resolve table by exact name and export directly
                tid = get_table_id_by_name(force_table)
                if not tid:
                    raise RuntimeError(f"JeraSoft table not found: {force_table}")

                js_df    = fetch_active_current_future_rates(table_id=int(tid))
                saved_to = save_rates_to_excel(js_df, output_path)

                rows_js = int(js_df.shape[0]) if hasattr(js_df, "shape") else 0
                meta["best_table_name"] = force_table
                meta["table_id"] = int(tid)  # Store table_id for JeraSoft upload

            else:
                # Normal path: build a safe non-empty target_query (fallback to subject if needed)
                bad_kw = {"", "a-z", "rates", "rate", "pricing", "update", "standard", "retail"}
                target_query = company if company.lower() not in bad_kw and len(company) >= 3 else subject
                if not target_query:
                    target_query = subject

                info = export_rates_by_query(
                    target_query=target_query,
                    output_path=output_path,
                    subject=subject,
                    prefix_code=prefix,
                    trunk_code=trunk,
                )
                if isinstance(info, str):
                    # export_rates_by_query returns a string on error.
                    # Route to invalid_subjects so an admin can assign a table manually.
                    _mark_subject_invalid_pending_jera(folder, meta, info)
                    return f"[{folder.name}] waiting for manual JERA table: {info}"

                # read back the file for row count
                rows_js = 0
                try:
                    ext = Path(output_path).suffix.lower()
                    df_js = pd.read_excel(output_path) if ext in (".xlsx", ".xls") else pd.read_csv(output_path)
                    rows_js = int(df_js.shape[0])
                except Exception:
                    pass

                meta["best_table_name"] = info.get("best_table_name")
                meta["table_id"] = info.get("table_id")  # Store table_id for JeraSoft upload

            # common metadata after either path
            meta["human_eval_details_jerasoft"] = {"file": Path(output_path).name, "rows": rows_js}
            meta["need_human_eval_jerasoft"] = bool(meta.get("need_human_eval_jerasoft")) or (rows_js < 100)
            meta["jerasoft_preprocessed"] = True
            meta["jera_fetched"] = True  # Flag for JeraSoft fetched in metadata
            save_metadata(folder, meta)

            try:
                mark_processing_stage(directory_name=folder.name, stage="jera_fetched")
            except Exception as e:
                print(f"[{folder.name}] stage warn (jera_fetched): {e}")

        except Exception as e:
            meta["keyword_error"] = str(e)
            meta["jera_fetched"] = False  # Mark JeraSoft fetch as failed in metadata
            save_metadata(folder, meta)
            
            # Update processing_statuses to mark jera_fetched as failed
            try:
                mark_processing_stage(
                    directory_name=folder.name,
                    stage="jera_fetched",
                    final_status=False,
                    error_message=str(e),
                )
            except Exception as db_error:
                print(f"[{folder.name}] stage warn (jera_fetched failed): {db_error}")
            
            return f"[{folder.name}] ✖ export failed: {e}"

    # -------- 2) Cleaning (if needed) --------
    # we consider "needed" if metadata has no 'preprocessed_results' or it's empty
    meta = load_metadata(folder) or {}
    pre_map: dict = meta.get("preprocessed_results", {}) or {}
    if not pre_map:
        # DB-first: if available, prefer vendor-level format over metadata for parsing vendor files
        fmt_db = vendor_fmt
        if not fmt_db:
            try:
                fmt_db = fetch_vendor_date_format_by_sender_email(meta.get("sender"))
            except Exception:
                fmt_db = None
        any_files = False
        for file_path in sorted(p for p in folder.iterdir() if p.is_file()):
            if file_path.name.lower() == "metadata.json":
                continue
            if file_path.suffix.lower() not in ALLOWED_EXTS:
                continue
            any_files = True

            in_path = str(file_path)
            out_path = str(cleaned_out_path(file_path))

            date_fmt = (fmt_db or (meta.get("date_format_identified") or "").strip() or None)
            fname = file_path.name.lower()
            # force ISO for jerasoft outputs
            date_fmt_to_use = "YYYY-MM-DD" if ("jerasoft_comparison_all" in fname or "jerasoft_comparison" in fname) else date_fmt
            vendor_header_mapping_to_use = None
            if "jerasoft" not in fname and bool(vendor_ctx.get("is_mapped")):
                if vendor_mapping_error:
                    raise ValueError(vendor_mapping_error)
                vendor_header_mapping_to_use = active_vendor_header_mapping

            try:
                cleaned_df = load_clean_rates(
                    in_path,
                    out_path,
                    0,
                    date_format_email=date_fmt_to_use,
                    vendor_header_mapping=vendor_header_mapping_to_use,
                )
                raw_name = Path(in_path).name
                clean_name = Path(out_path).name
                pre_map[raw_name] = True
                pre_map[clean_name] = True
                try:
                    err_map = meta.get("preprocess_errors")
                    if isinstance(err_map, dict):
                        err_map.pop(raw_name, None)
                        err_map.pop(clean_name, None)
                        if err_map:
                            meta["preprocess_errors"] = err_map
                        else:
                            meta.pop("preprocess_errors", None)
                except Exception:
                    pass

                # small-output hint
                try:
                    row_count = int(getattr(cleaned_df, "shape", [0])[0])
                except Exception:
                    row_count = 0
                meta.setdefault("human_eval_details_pre", {})[raw_name] = {"rows": row_count}
                meta["need_human_eval_pre"] = bool(meta.get("need_human_eval_pre")) or (row_count < 100)
            except Exception as e:
                pre_map[file_path.name] = False
                print(f"[{folder.name}] clean fail {file_path.name}: {e}")

                # Track detailed preprocessing error per file in metadata so we can
                # later see which columns failed to map / other reasons for failure.
                # This keeps preprocessed_results structure (filename -> bool) intact
                # while adding a parallel map with error strings.
                try:
                    err_map = meta.get("preprocess_errors") or {}
                    # Store a simple string; the message from _canonicalize_headers
                    # already includes "Missing required columns: [...]" when mapping fails.
                    err_map[str(file_path.name)] = str(e)
                    meta["preprocess_errors"] = err_map
                except Exception:
                    # Never let metadata error break the main cleaning loop.
                    pass

        if any_files:
            meta["preprocessed_results"] = pre_map
            save_metadata(folder, meta)

            # --- Check if any JeraSoft or Vendor file is False ---
            # Split flags into JeraSoft and Vendor based on filename
            jera_flags = [v for name, v in pre_map.items() if "jerasoft" in name.lower()]
            vendor_flags = [v for name, v in pre_map.items() if "jerasoft" not in name.lower()]

            # Flag is false if any file in either group is False
            jera_failed = any(not v for v in jera_flags)
            vendor_failed = any(not v for v in vendor_flags)

            # Final status depends on if there is any failure in either group
            final_ok = not (jera_failed or vendor_failed)

            meta["final_ok"] = final_ok
            save_metadata(folder, meta)

            # If cleaning failed, build a short reason string from metadata
            # (e.g. first entry from preprocess_errors) and push it to
            # processing_statuses.status via mark_processing_stage.
            error_text = None
            if not final_ok:
                err_map = meta.get("preprocess_errors") or {}
                if isinstance(err_map, dict) and err_map:
                    # Only show the reason, not the specific file name, so
                    # statuses look like "failed: Header not found ..." or
                    # "failed: Missing required columns ...".
                    _, first_err = next(iter(err_map.items()))
                    error_text = str(first_err)
                else:
                    error_text = "file_cleaned failed (see metadata.json for details)"

            try:
                mark_processing_stage(
                    directory_name=folder.name,
                    stage="file_cleaned",
                    final_status=final_ok,
                    error_message=error_text,
                )
            except Exception as e:
                print(f"[{folder.name}] stage warn (file_cleaned): {e}")

            # Optional: helpful debug logs
            if not final_ok:
                print(
                    f"[{folder.name}] file_cleaned: FALSE — "
                    f"JeraSoft failed={jera_failed} (flags={jera_flags}), "
                    f"Vendor failed={vendor_failed} (flags={vendor_flags})"
                )
            else:
                print(
                    f"[{folder.name}] file_cleaned: TRUE — "
                    f"JeraSoft ok={not jera_failed}, Vendor ok={not vendor_failed}"
                )


    # -------- 3) Comparison (if needed) --------
    meta = load_metadata(folder) or {}
    if "comparision_result" not in (meta.keys()):
        # Track whether we've already generated a full-new comparison so we
        # can skip the baseline compare block below.
        skip_baseline = False

        # Hard stop: if preprocessing failed, do NOT proceed to comparison.
        # This prevents compare/DB push when cleaned outputs are incomplete or invalid.
        if meta.get("final_ok") is False:
            # Special case: JeraSoft export succeeded but has 0 rows.
            # In this case, treat the vendor file as a full NEW import
            # instead of failing the whole pipeline.
            jera_info = meta.get("human_eval_details_jerasoft") or {}
            rows_js = jera_info.get("rows")
            zero_row_js = isinstance(rows_js, int) and rows_js == 0 and bool(meta.get("jera_fetched"))

            if zero_row_js:
                pre_map_raw = meta.get("preprocessed_results") or {}

                def _norm_local(s: str) -> str:
                    return str(s).strip().casefold()

                pre_map_ci = {_norm_local(k): v for k, v in pre_map_raw.items()}

                vfiles = vendor_files(folder)
                if not vfiles:
                    meta["comparision_result"] = {"result": "no vendor files to import as new"}
                    save_metadata(folder, meta)
                    # No comparison results means nothing for DB push; stop here.
                    return f"[{folder.name}] no vendor files for full-new import (Jera rows=0)"

                comp_result: Dict[str, bool] = {}
                writes = 0

                as_of_date = as_of_from_metadata(folder)

                for v in vfiles:
                    vname = v.name
                    vkey = _norm_local(vname)
                    raw_v = vkey.replace("_cleaned.xlsx", ".xlsx") if vkey.endswith("_cleaned.xlsx") else vkey
                    vendor_ok = bool(pre_map_ci.get(vkey) or pre_map_ci.get(raw_v))
                    if not vendor_ok:
                        print(f"[{folder.name}] full-new import skip {vname}: not preprocessed")
                        comp_result[vname] = False
                        continue

                    try:
                        # Read the cleaned vendor sheet and synthesize a comparison-style
                        # table where every row is New/Accepted.
                        df_vendor = read_table(str(v), None)
                        result_df = build_full_new_comparison_from_vendor(df_vendor)

                        out_path = folder / f"{v.stem}_comparision_result.xlsx"
                        write_excel(result_df, str(out_path))
                        print(f"[{folder.name}] wrote full-new result to {out_path} (Jera rows=0)")

                        # Optional attachment stats, consistent with normal path
                        stats = summarize_changes(result_df)

                        comp_result[vname] = True
                        writes += 1

                        meta.setdefault("attachment_stats", {})
                        meta["attachment_stats"][v.name] = {
                            **stats,
                            "source_attachment": v.name,
                            "result_file": out_path.name,
                            "generated_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                        }

                        # Also store comparison file path for this vendor so
                        # rate_upload_service can recover the original vendor
                        # attachment name and build a friendly CSV filename
                        # for JeraSoft Import History (like manual uploads).
                        try:
                            jerasoft_table_id = meta.get("table_id") or meta.get("best_table_id")
                            meta.setdefault("jerasoft_upload", {})
                            meta["jerasoft_upload"][v.name] = {
                                "table_id": jerasoft_table_id,
                                "comparison_file": str(out_path.absolute()),
                                "created_at": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                                "status": "ready_for_manual_upload",
                            }
                        except Exception:
                            # Never block processing if metadata update fails.
                            pass
                        save_metadata(folder, meta)

                    except Exception as e:
                        comp_result[vname] = False
                        print(f"[{folder.name}] full-new import fail {vname}: {e}")

                if comp_result:
                    success_any = any(comp_result.values())
                    meta["comparision_result"] = {"result": "ok" if success_any else "no comparisons succeeded", **comp_result}
                else:
                    meta["comparision_result"] = {"result": "no eligible vendor files"}

                meta["processed_at_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
                save_metadata(folder, meta)

                try:
                    mark_processing_stage(directory_name=folder.name, stage="rate_compared")
                except Exception as e:
                    print(f"[{folder.name}] stage warn (rate_compared, full-new): {e}")

                # We already generated a full-new comparison result; skip
                # the baseline comparison logic below and go straight to
                # the DB push section.
                skip_baseline = True

            else:
                # Default behaviour: if preprocessing failed and it's *not* the
                # zero-row JeraSoft case, skip comparison as before.
                meta["comparision_result"] = {"result": "comparison skipped: preprocessing failed (final_ok=false)"}
                save_metadata(folder, meta)
                return f"[{folder.name}] skip compare: preprocessing failed"

        # If we already handled the zero-row JeraSoft case and built a
        # full-new comparison, skip the baseline comparison entirely and
        # fall through to the DB push section.
        if skip_baseline:
            pass
        else:
            # pre_map = meta.get("preprocessed_results", {}) or {}

            pre_map_raw = meta.get("preprocessed_results") or {}
            pre_map = {k.lower(): v for k, v in pre_map_raw.items()}

            left_path = find_jerasoft_file(folder)

            if left_path:
                print(f"[{folder.name}] DEBUG: baseline file chosen = {left_path.name!r}")
                print(f"[{folder.name}] DEBUG: pre_map keys ({len(pre_map)}): {[k for k in pre_map.keys()]}")
                for k in pre_map.keys():
                    print(f"[{folder.name}] DEBUG: compare key={k!r} == left? {k == left_path.name}  "
                        f"len(key)={len(k)} len(left)={len(left_path.name)}")

                print(f"[{folder.name}] DEBUG: baseline file chosen = {left_path}")
                print(f"[{folder.name}] DEBUG: baseline file chosen = {left_path.name}")
         
                print(f"[{folder.name}] DEBUG: pre_map keys = {list(pre_map.keys())[:50]}")
                print(f"[{folder.name}] DEBUG: exact match? {left_path.name in pre_map}")
                print(f"[{folder.name}] DEBUG: pre_map[left] = {pre_map.get(left_path.name)}")
                print(f"[{folder.name}] DEBUG: lower match? {left_path.name.lower() in {k.lower(): v for k,v in pre_map.items()}}")
            else:
                print(f"[{folder.name}] DEBUG: baseline file chosen = None")

            if not left_path:
                meta["comparision_result"] = {"result": "comparison skipped: no baseline file found"}
                save_metadata(folder, meta)
                return f"[{folder.name}] skip compare: no baseline"

            # baseline_ok = bool(pre_map.get(left_path.name))
            def norm(s: str) -> str: return s.strip().casefold()
            pre_map_ci = {norm(k): v for k, v in (meta.get("preprocessed_results") or {}).items()}
            lp = left_path.name
            lp_n = norm(lp)
            raw_n = norm(lp.replace("_cleaned.xlsx", ".xlsx")) if lp_n.endswith("_cleaned.xlsx") else lp_n
            baseline_ok = bool(pre_map_ci.get(lp_n) or pre_map_ci.get(raw_n))




            if not baseline_ok:
                meta["comparision_result"] = {"result": "comparison skipped: comparison file failed preprocessing"}
                save_metadata(folder, meta)
                return f"[{folder.name}] skip compare: baseline not preprocessed"

            vfiles = vendor_files(folder)
            if not vfiles:
                meta["comparision_result"] = {"result": "no vendor files to compare"}
                save_metadata(folder, meta)
                return f"[{folder.name}] no vendor files"

            as_of_date = as_of_from_metadata(folder)
            try:
                left_df = read_table(str(left_path), None)
            except Exception as e:
                meta["comparision_result"] = {"result": f"comparison skipped: failed to read baseline ({e})"}
                save_metadata(folder, meta)
                return f"[{folder.name}] skip compare: read baseline fail"

            comp_result: Dict[str, bool] = {}
            writes = 0
        # for v in vfiles:
        #     vname = v.name
        #     if not pre_map.get(vname):
        #         comp_result[vname] = False
        #         continue
        #     try:
        #         right_df = read_table(str(v), None)
        #         result, stats = compare(left_df, right_df, as_of_date, 7, 0.0001)
        #         out_path = folder / f"{v.stem}_comparision_result.xlsx"
        #         write_excel(result, str(out_path))
        #         print(f"[{folder.name}] wrote result to {out_path}")

        #         writes += 1
        #         comp_result[vname] = True

        #         meta.setdefault("attachment_stats", {})
        #         meta["attachment_stats"][v.name] = {
        #             **stats,
        #             "source_attachment": v.name,
        #             "result_file": out_path.name,
        #             "generated_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        #         }
        #         save_metadata(folder, meta)

        
        
        def _norm(s: str) -> str:
            return s.strip().casefold()

        # reuse the case-insensitive map you already build above:
        # pre_map_ci = { _norm(k): v }  (it already exists in your code)

        if not skip_baseline:
            for v in vfiles:
                vname = v.name
                vkey  = _norm(vname)
                raw_v = vkey.replace("_cleaned.xlsx", ".xlsx") if vkey.endswith("_cleaned.xlsx") else vkey
                vendor_ok = bool(pre_map_ci.get(vkey) or pre_map_ci.get(raw_v))
                if not vendor_ok:
                    print(f"[{folder.name}] compare skip {vname}: not preprocessed (have keys={list(pre_map_ci.keys())})")
                    comp_result[vname] = False
                    continue

                try:
                    right_df = read_table(str(v), None)
                    # exact match mode (no tolerance)
                    # show progress for large comparisons so it doesn't look "stuck"
                    progress_every = int(os.getenv("COMPARE_PROGRESS_EVERY", "0") or "0")
                    result, stats = compare(left_df, right_df, as_of_date, 6, 0.0, is_partial=is_partial_vendor, progress_every=progress_every)
                    out_path = folder / f"{v.stem}_comparision_result.xlsx"
                    write_excel(result, str(out_path))
                    print(f"[{folder.name}] wrote result to {out_path}")

                    writes += 1
                    comp_result[vname] = True

                    # Get table_id from metadata for JeraSoft upload
                    jerasoft_table_id = meta.get("table_id") or meta.get("best_table_id")
                    
                    # Mark for JeraSoft BULK upload in database (always bulk, never individual)
                    jera_upload_enabled = os.getenv("JERASOFT_DB_CONTROL", "true").lower() in ("true", "1", "yes")
                    min_rates = int(os.getenv("JERASOFT_MIN_RATES_FOR_AUTO_UPLOAD", "1"))
                    
                    if jera_upload_enabled and jerasoft_table_id and len(result) >= min_rates:
                        try:
                            print(f"[{folder.name}] JeraSoft table {jerasoft_table_id} detected - comparison file ready for MANUAL bulk upload")
                            print(f"[{folder.name}] ℹ️  To upload: Set is_rate_approved_by_admin = TRUE in database for this record")
                            
                            # Store comparison file path in metadata for future manual upload
                            # No automatic marking - manual control required
                            subject = meta.get("subject", f"Rate Update - {folder.name}")
                            sender_email = meta.get("sender", "unknown@example.com")
                            
                            # Store file info for manual upload control
                            print(f"[{folder.name}] 📋 Comparison file ready: {out_path.name}")
                            print(f"[{folder.name}] 🎯 Target table: {jerasoft_table_id}")
                            print(f"[{folder.name}] ⚡ Manual control: Set is_rate_approved_by_admin=TRUE in database to trigger bulk upload")
                            
                            # Store upload info in metadata (for reference only)
                            meta.setdefault("jerasoft_upload", {})
                            meta["jerasoft_upload"][v.name] = {
                                "table_id": jerasoft_table_id,
                                "comparison_file": str(out_path.absolute()),
                                "subject": subject,
                                "sender_email": sender_email,
                                "created_at": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                                "status": "ready_for_manual_upload"
                            }
                            
                        except Exception as upload_error:
                            print(f"[{folder.name}] ❌ Failed to mark for JeraSoft upload: {upload_error}")
                            meta.setdefault("jerasoft_upload", {})
                            meta["jerasoft_upload"][v.name] = {
                                "table_id": jerasoft_table_id,
                                "error": str(upload_error),
                                "uploaded_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                                "status": "failed"
                            }
                    else:
                        if not jerasoft_table_id:
                            print(f"[{folder.name}] ⚠️ No table_id in metadata, skipping JeraSoft upload")
                        elif len(result) == 0:
                            print(f"[{folder.name}] ⚠️ No comparison results, skipping JeraSoft upload")

                    meta.setdefault("attachment_stats", {})
                    meta["attachment_stats"][v.name] = {
                        **stats,
                        "source_attachment": v.name,
                        "result_file": out_path.name,
                        "generated_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                    }
                    save_metadata(folder, meta)



                except Exception as e:
                    comp_result[vname] = False
                    print(f"\n\n\n\n\n[{folder.name}] compare fail {vname}: {e}\n\n\n\n")

        if not skip_baseline:
            if comp_result:
                success_any = any(comp_result.values())
                print(f"\n\nDEBUG: Putting comparision result in the metadata file \n\n")
                meta["comparision_result"] = {"result": "ok" if success_any else "no comparisons succeeded", **comp_result}
            else:
                meta["comparision_result"] = {"result": "no eligible vendor files"}
            meta["processed_at_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
            save_metadata(folder, meta)

            try:
                mark_processing_stage(directory_name=folder.name, stage="rate_compared")
            except Exception as e:
                print(f"[{folder.name}] stage warn (rate_compared): {e}")

    # -------- 4) DB push (per-folder) --------
    meta = load_metadata(folder) or {}
    comp = meta.get("comparision_result")
    if not (isinstance(comp, dict) and str(comp.get("result", "")).strip().lower() == "ok"):
        return f"[{folder.name}] skip DB push: result not ok"

    # Special handling for the zero-row JeraSoft case where we treated the
    # vendor sheet as a full NEW import. In that path, the cleaning step
    # may have marked file_cleaned as failed (because the Jera file has
    # no rows), which blocks later stages in processing_statuses. Once we
    # have a successful comparison result, we can backfill the stage flags
    # so processing_statuses reflects the true outcome for this message.
    jera_info = meta.get("human_eval_details_jerasoft") or {}
    rows_js = jera_info.get("rows")
    zero_row_js = isinstance(rows_js, int) and rows_js == 0 and bool(meta.get("jera_fetched"))

    if zero_row_js:
        try:
            # Ensure file_cleaned and rate_compared are marked as completed
            mark_processing_stage(directory_name=folder.name, stage="file_cleaned")
            mark_processing_stage(directory_name=folder.name, stage="rate_compared")

            # If comparison results were already pushed in a prior run,
            # also advance the final rate_uploaded stage so all flags are
            # TRUE for this special full-new import case.
            rp_existing = meta.get("results_pushed") or {}
            if any(v is True for v in rp_existing.values()):
                mark_processing_stage(directory_name=folder.name, stage="rate_uploaded", final_status=True)
        except Exception as e:
            print(f"[{folder.name}] stage warn (zero-row Jera backfill): {e}")

    # discover result files
    result_files = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xls", ".csv")
        and p.stem.lower().endswith("_comparision_result")
    )
    if not result_files:
        return f"[{folder.name}] no result files to push"

    rp = meta.get("results_pushed") or {}
    to_push = [f for f in result_files if rp.get(f.name) is not True]
    if not to_push:
        return f"[{folder.name}] already pushed"

    # read DFs (cache) and aggregate stats (only non-empty)
    dfs_to_push: List[pd.DataFrame] = []
    per_file_df: Dict[str, pd.DataFrame] = {}
    for f in to_push:
        try:
            df_tmp = read_comparison_table(f)
            per_file_df[f.name] = df_tmp
            if not df_tmp.empty:
                dfs_to_push.append(df_tmp)
        except Exception as e:
            print(f"[{folder.name}] ⚠ read for stats {f.name}: {e}")
            per_file_df[f.name] = pd.DataFrame()

    if not any((not d.empty) for d in per_file_df.values()):
        for f in to_push:
            rp[f.name] = "empty file no results to push to the database"
        meta["results_pushed"] = rp
        save_metadata(folder, meta)
        return f"[{folder.name}] all empty; no upload"


    stats_totals = compute_upload_stats(dfs_to_push)
    try:
        df_all = pd.concat(dfs_to_push, ignore_index=True)
        new_rate_ge_1 = pd.to_numeric(df_all.get("New Rate"), errors="coerce") >= 1.0
        ct = df_all.get("Change Type")
        if ct is None:
            not_unchanged = True
        else:
            ct_s = ct.astype(str)
            # Exclude rows where Change Type includes "Unchanged" (supports multi-label values like "X,Unchanged,Y")
            is_unchanged = ct_s.str.contains(r"(?:(?<=^)|(?<=,))\s*unchanged\s*(?:(?=,)|(?=$))", case=False, regex=True, na=False)
            not_unchanged = ~is_unchanged
        rates_gte_one_usd_count = int((new_rate_ge_1 & not_unchanged).sum())
    except Exception:
        rates_gte_one_usd_count = 0

    sender = str(meta.get("sender") or "").strip() or None
    subject = (meta.get("subject") or "").strip() or None

 
    received_at = parse_received_at(meta)
    processed_at = None
    try:
        rawp = meta.get("processed_at_utc")
        processed_at = datetime.fromisoformat((rawp or "").replace("Z", "+00:00")) if rawp else None
    except Exception:
        processed_at = None

    # Get table_id from metadata
    jera_table_id = meta.get("table_id") or meta.get("best_table_id")
    
    # Extract comparison file path from metadata (jerasoft_upload section)
    comparison_file_path = None
    jerasoft_upload = meta.get("jerasoft_upload", {})
    if jerasoft_upload:
        # Get the first available comparison file path
        for upload_info in jerasoft_upload.values():
            if isinstance(upload_info, dict) and "comparison_file" in upload_info:
                comparison_file_path = upload_info["comparison_file"]
                break
    
    # Internet message id to link rate_uploads back to processing_statuses
    internet_message_id = str(
        meta.get("internet_message_id")
        or meta.get("internetMessageId")
        or meta.get("message_id")
        or ""
    ).strip() or None

    processing_status_id = None
    try:
        ps = get_processing_status(directory_name=folder.name)
        processing_status_id = int(ps["id"]) if ps and ps.get("id") is not None else None
    except Exception:
        processing_status_id = None
    trunk_for_db = _derive_trunk(meta)

    upload_id = meta.get("rate_upload_id")
    if not upload_id:
        try:
            upload_id = insert_rate_upload(
                sender_email=sender,
                subject=subject,
                received_at=received_at,
                processed_at=processed_at,
                totals=stats_totals,
                rates_gte_one_usd_count=rates_gte_one_usd_count,
                jera_table_id=jera_table_id,
                comparison_file_path=comparison_file_path,
                internet_message_id=internet_message_id,
                processing_status_id=processing_status_id,
                trunk=trunk_for_db,
            )
            meta["rate_upload_id"] = int(upload_id)
            save_metadata(folder, meta)
        except Exception as e:
            for f in to_push:
                rp[f.name] = False
            meta["results_pushed"] = rp
            save_metadata(folder, meta)
            return f"[{folder.name}] ✖ create upload row: {e}"

    # push each result file
    pushed_any = False
    rp = meta.get("results_pushed") or {}
    for f in to_push:
        try:
            # >>> CHANGED: never use "or" with a DataFrame on the left
            df = per_file_df.get(f.name)  # may be None
            if df is None:                # >>> CHANGED
                df = read_comparison_table(f)  # >>> CHANGED

            # Safe emptiness check
            if df is None or getattr(df, "empty", True):  # >>> CHANGED
                rp[f.name] = "empty file no results to push to the database"
                print(f"[{folder.name}]    ⚠ empty comparison file; nothing to push")  # >>> ADDED
                continue

            try:  # >>> ADDED (optional logging)
                print(f"[{folder.name}]    info: {f.name} df shape = {tuple(df.shape)}")
            except Exception:
                pass

            details = df_to_detail_dicts(df, received_at=received_at)
            if not details:  # list truthiness is fine
                rp[f.name] = "no rows extracted to push"
                print(f"[{folder.name}]    ⚠ no detail rows extracted; nothing to push")  # >>> ADDED
                continue

            inserted = bulk_insert_rate_upload_details(int(upload_id), details)
            rp[f.name] = True
            pushed_any = pushed_any or (inserted > 0)
            print(f"[{folder.name}]    ✔ inserted {inserted} rows from {f.name}")  # >>> ADDED

        except Exception as e:
            rp[f.name] = False
            print(f"[{folder.name}] ✖ push {f.name}: {e}")
            print(traceback.format_exc())  # >>> ADDED

    meta["results_pushed"] = rp
    save_metadata(folder, meta)

    if pushed_any:
        try:
            mark_processing_stage(directory_name=folder.name, stage="rate_uploaded", final_status=True)
        except Exception as e:
            print(f"[{folder.name}] stage warn (rate_uploaded): {e}")

    return f"[{folder.name}] done (pushed={sum(1 for v in rp.values() if v is True)}/{len(rp)})"


# ----------------- pool runner -----------------

def run_pipeline_mt(attachments_base: str = "attachments", max_workers: int = 4) -> None:
    base = Path(attachments_base)
    if not base.exists():
        print(f"[PIPELINE] attachments base not found: {base}")
        return

    # choose folders to process (those with metadata.json)
    folders = []
    for d in base.iterdir():
        if not d.is_dir():
            continue
        if (d / "metadata.json").exists():
            folders.append(d)

    if not folders:
        print("[PIPELINE] no folders to process.")
        return

    # IMPORTANT: serialize processing per vendor+trunk (per sender email + trunk).
    # Our business rule blocks only vendor+trunk when a pending upload exists today,
    # so vendor A / trunk X can run in parallel with vendor A / trunk Y.
    queues: dict[tuple[str, str], list[Path]] = {}
    for d in folders:
        meta = load_metadata(d) or {}
        sender = str(meta.get("sender") or "").strip().casefold() or "__unknown_sender__"
        trunk = _derive_trunk(meta)
        trunk_key = (str(trunk).strip() if trunk is not None else "") or "__unknown_trunk__"
        key = (sender, trunk_key.casefold())
        queues.setdefault(key, []).append(d)

    for q in queues.values():
        q.sort(key=lambda p: p.name)

    print(f"[PIPELINE] starting multithread pipeline: {len(folders)} folder(s), workers={max_workers}")

    def _process_vendor_queue(vendor_key: tuple[str, str], vendor_folders: list[Path]) -> None:
        for d in vendor_folders:
            try:
                msg = process_one_folder(d)
            except Exception as e:
                msg = f"[{d.name}] ✖ pipeline error: {e}\n{traceback.format_exc()}"
            print(msg)

    with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="pipe") as ex:
        futs = [
            ex.submit(_process_vendor_queue, vendor_key, vendor_folders)
            for vendor_key, vendor_folders in queues.items()
        ]
        for fut in as_completed(futs):
            # surface any unexpected exceptions
            fut.result()

