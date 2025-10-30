import os
from dotenv import load_dotenv
import psycopg2
from valid_emails import VERIFIED_SENDERS
from typing import Iterable, Dict, Any, Optional, List, Mapping, Tuple
from datetime import datetime
from decimal import Decimal
import json
from psycopg2.extras import execute_values, Json 
from pathlib import Path
from typing import Optional


# Load environment variables
load_dotenv()

def get_conn():
    """Create a new PostgreSQL connection using env vars."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_DATABASE"),
        user=os.getenv("DB_USERNAME"),
        password=os.getenv("DB_PASSWORD"),
    )

def insert_authorized_senders(emails):
    """
    Insert a list of emails into authorized_senders.
    Uses ON CONFLICT DO NOTHING to avoid duplicate errors.
    """
    query = """
        INSERT INTO authorized_senders (email, status, created_at, updated_at)
        VALUES %s
        ON CONFLICT (email) DO NOTHING;
    """

    # Prepare rows: status = true, timestamps = now()
    rows = [(email, True, "NOW()", "NOW()") for email in emails]

    # psycopg2 cannot handle NOW() as string, so use SQL functions directly
    values_template = "(%s, %s, NOW(), NOW())"

    with get_conn() as conn, conn.cursor() as cur:
        execute_values(cur, query, [(email, True) for email in emails], template=values_template)
        conn.commit()
        print(f"Inserted {cur.rowcount} new emails into authorized_senders.")


# -----------------------------
# Rejected emails (row-per-item) support
# -----------------------------

def _parse_iso_utc(s: Optional[str]) -> Optional[datetime]:
    """
    Parse an ISO timestamp like '2025-09-28T12:34:56Z' into a tz-aware UTC datetime.
    Returns None if parsing fails or s is falsy.
    """
    if not s or not isinstance(s, str):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None

def insert_rejected_email(
    sender_email: Optional[str],
    subject: Optional[str],
    category: str,
    notes: Optional[str],
    received_at: Optional[datetime],
    processed_at: Optional[datetime],
) -> int:
    """
    Insert a single rejected email row and return its id.
    Table columns (managed by DB): id (PK), created_at, updated_at auto.
    """
    sql = """
        INSERT INTO rejected_emails
        (sender_email, subject, category, notes, received_at, processed_at, created_at, updated_at)
        VALUES (%s, %s, %s, %s, %s, %s, NOW(), NOW())
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (sender_email, subject, category, notes, received_at, processed_at),
        )
        new_id = cur.fetchone()[0]
        conn.commit()
        return new_id

# import here to avoid requiring psycopg2 unless this function is used
def insert_rejected_emails(rows: Iterable[Mapping[str, Any]]) -> List[int]:
    """
    Bulk-insert multiple rejected email rows and return list of new ids in the same order.

    `rows` should be an iterable of mappings with keys:
      - sender_email (Optional[str])
      - subject (Optional[str])
      - category (str)                # required
      - notes (Optional[str])
      - received_at (Optional[datetime])
      - processed_at (Optional[datetime])

    Returns: list of inserted ids (may be empty).
    """
    # collect values in the order expected by the DB
    values = []
    for r in rows:
        # ensure category exists (raise so caller notices bad input)
        category = r.get("category")
        if category is None:
            raise ValueError("Each row must include a 'category' value.")
        values.append((
            r.get("sender_email"),
            r.get("subject"),
            str(category),
            r.get("notes"),
            r.get("received_at"),
            r.get("processed_at"),
        ))

    if not values:
        return []

    # use psycopg2.extras.execute_values for efficient bulk insert + RETURNING
    try:
        from psycopg2.extras import execute_values
    except Exception as e:
        raise RuntimeError("psycopg2.extras.execute_values is required for bulk insert.") from e

    sql = """
        INSERT INTO rejected_emails
        (sender_email, subject, category, notes, received_at, processed_at, created_at, updated_at)
        VALUES %s
        RETURNING id;
    """

    with get_conn() as conn, conn.cursor() as cur:
        # execute_values will expand the VALUES %s placeholder into many tuples
        execute_values(cur, sql, values, template=None, page_size=100)
        ids = [row[0] for row in cur.fetchall()]
        conn.commit()
    return ids

def _atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    """Write JSON atomically to avoid corruption on crashes."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)



def insert_rejected_email_row(
    *,
    sender_email: Optional[str],
    subject: Optional[str],
    category: str,
    notes: Optional[str],
    received_at: Optional[datetime],
    processed_at: Optional[datetime],
) -> int:
    """
    Insert a single row into rejected_emails and return its id.

    Schema (expected):
      rejected_emails(
        id SERIAL PK,
        sender_email TEXT,
        subject TEXT,
        category TEXT,
        notes TEXT,
        received_at TIMESTAMPTZ,
        processed_at TIMESTAMPTZ,
      )
    """
    sql = """
        INSERT INTO rejected_emails
          (sender_email, subject, category, notes, received_at, processed_at)
        VALUES (%s, %s, %s, %s, %s, %s)
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (sender_email, subject, category, notes, received_at, processed_at),
        )
        new_id = cur.fetchone()[0]
        conn.commit()
        return new_id

def push_failed_emails_json_to_db(path: Optional[str | Path] = None) -> tuple[int, int, int]:
    """
    Read failed_emails.json and insert any entries not yet pushed (no 'already_pushed': true)
    into the rejected_emails table using a bulk insert. After a successful insert, mark the JSON
    entry with 'already_pushed': true and atomically rewrite the JSON file.

    Returns:
        (inserted_count, skipped_already_pushed, errors)
    """
    json_path = Path(path) if path else Path(__file__).with_name("failed_emails.json")
    if not json_path.exists():
        print(f"(info) {json_path} not found; nothing to push.")
        return (0, 0, 0)

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"(warn) could not read {json_path}: {e}")
        return (0, 0, 1)

    buckets = (data or {}).get("buckets") or {}
    if not isinstance(buckets, dict):
        print(f"(warn) malformed failed emails JSON: missing 'buckets' object")
        return (0, 0, 1)

    inserted = 0
    skipped = 0
    errors = 0
    print(f"DEBUG: Pushing failed emails from {json_path}, categories: {list(buckets.keys())}")
    # Prepare a list of rows to bulk-insert and keep references to the original entries
    rows_to_insert: list[dict] = []
    entry_refs: list[dict] = []

    for category, items in buckets.items():
        if not isinstance(items, list):
            continue

        for entry in items:
            # Skip if already pushed
            if bool(entry.get("already_pushed")):
                skipped += 1
                continue

            sender_email = (entry.get("sender") or "").strip() or None
            subject = (entry.get("subject") or "").strip() or None
            received_at = _parse_iso_utc(entry.get("receivedDateTime"))
            processed_at = _parse_iso_utc(entry.get("logged_at_utc"))

            # Build a compact notes string that does not restate the category itself.
            details = entry.get("details")
            if details is None:
                notes = ""
            else:
                try:
                    notes_full = json.dumps(details, ensure_ascii=False)
                except Exception:
                    notes_full = str(details)
                notes = notes_full if len(notes_full) <= 2000 else (notes_full[:1970] + "...(truncated)")

            rows_to_insert.append({
                "sender_email": sender_email,
                "subject": subject,
                "category": str(category),
                "notes": notes,
                "received_at": received_at,
                "processed_at": processed_at,
            })
            entry_refs.append(entry)

    # Nothing to insert
    if not rows_to_insert:
        print(f"(info) no new failed emails to push. skipped={skipped}")
        return (0, skipped, errors)

    # Attempt bulk insert
    try:
        ids = insert_rejected_emails(rows_to_insert)  # expects list of ids
        # Normalize returned ids to a list
        if isinstance(ids, int):
            ids = [ids]
        elif ids is None:
            ids = []

        # Mark the corresponding JSON entries as pushed
        for i in range(min(len(ids), len(entry_refs))):
            entry_refs[i]["already_pushed"] = True
        inserted = len(ids)

    except Exception as bulk_exc:
        print(f"(warn) bulk insert failed: {bulk_exc}. Falling back to single-row inserts.")
        # Fallback: try inserting row-by-row so partial progress is possible
        for i, row in enumerate(rows_to_insert):
            entry = entry_refs[i]
            try:
                # Prefer single-row API if available; otherwise call bulk API with single item
                try:
                    new_id = insert_rejected_email(
                        sender_email=row["sender_email"],
                        subject=row["subject"],
                        category=row["category"],
                        notes=row["notes"],
                        received_at=row["received_at"],
                        processed_at=row["processed_at"],
                    )
                except NameError:
                    # insert_rejected_email not defined, try bulk function for single row
                    res = insert_rejected_emails([row])
                    new_id = (res[0] if res else 0) if isinstance(res, list) else (res or 0)

                # If we got here without exception, mark pushed
                entry["already_pushed"] = True
                inserted += 1
            except Exception as single_exc:
                errors += 1
                print(f"(warn) failed to insert rejected email (category={row['category']}): {single_exc}")

    # Persist the updated JSON with the already_pushed flags
    try:
        _atomic_write_json(json_path, data)
    except Exception as e:
        # If this write fails, you've still inserted rows, but flags weren't saved.
        # Next run may try to re-insert. Consider adding a uniqueness constraint if needed.
        errors += 1
        print(f"(warn) failed to update {json_path} with 'already_pushed' flags: {e}")

    print(f"(ok) rejected_emails sync → inserted={inserted}, skipped={skipped}, errors={errors}")
    return (inserted, skipped, errors)

def insert_rate_upload(
    *,
    sender_email: Optional[str],
    subject: Optional[str],
    received_at: Optional[datetime] = None,
    processed_at: Optional[datetime] = None,
    totals: Optional[Dict[str, int]] = None,
) -> int:
    """
    Insert one row into rate_uploads with summary counters.

    rate_uploads columns covered:
      subject, sender_email, received_at, processed_at,
      total_rows, new, increase, decrease, unchanged, closed,
      backdated_increase, backdated_decrease, billing_increment_changes
    """
    t = {
        "total_rows": 0,
        "new": 0,
        "increase": 0,
        "decrease": 0,
        "unchanged": 0,
        "closed": 0,
        "backdated_increase": 0,
        "backdated_decrease": 0,
        "billing_increment_changes": 0,
    }
    if totals:
        t.update({k: int(v) for k, v in totals.items() if k in t})

    sql = """
        INSERT INTO rate_uploads
        (subject, sender_email, received_at, processed_at,
         total_rows, "new", increase, decrease, unchanged, closed,
         backdated_increase, backdated_decrease, billing_increment_changes,
         created_at, updated_at)
        VALUES
        (%s, %s, COALESCE(%s, NOW()), %s,
         %s, %s, %s, %s, %s, %s,
         %s, %s, %s,
         NOW(), NOW())
        RETURNING id;
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (
                subject,
                sender_email,
                received_at,
                processed_at,
                t["total_rows"],
                t["new"],
                t["increase"],
                t["decrease"],
                t["unchanged"],
                t["closed"],
                t["backdated_increase"],
                t["backdated_decrease"],
                t["billing_increment_changes"],
            ),
        )
        new_id = cur.fetchone()[0]
        conn.commit()
        return new_id

def _chunked(seq: List[tuple], size: int):
    for i in range(0, len(seq), size):
        yield seq[i:i+size]

def bulk_insert_rate_upload_details(
    rate_upload_id: int,
    details: Iterable[Dict[str, Any]],
    batch_size: int = 1000,   # tune as needed
) -> int:
    """
    Bulk insert rows and return the TRUE total inserted count.
    Works even when execute_values paginates internally.
    """
    rows = [
        (
            rate_upload_id,
            d["dst_code"],
            d.get("rate_existing"),
            d.get("rate_new"),
            d.get("effective_date"),
            d.get("change_type"),
            d.get("status"),
            d.get("notes"),
            d.get("old_billing_increment"),   # New field
            d.get("new_billing_increment"),   # New field
            d.get("code_name"),               # New field
        )
        for d in details
    ]
    if not rows:
        return 0

    sql = """
        INSERT INTO rate_upload_details
        (rate_upload_id, dst_code, rate_existing, rate_new, effective_date,
         change_type, status, notes, old_billing_increment, new_billing_increment, code_name, created_at, updated_at)
        VALUES %s
        RETURNING 1
    """
    tpl = "(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW(),NOW())"

    total = 0
    with get_conn() as conn, conn.cursor() as cur:
        for chunk in _chunked(rows, batch_size):
            # Ensure a single statement per chunk by making page_size=len(chunk)
            try:
                returned = execute_values(
                    cur, sql, chunk, template=tpl, page_size=len(chunk), fetch=True
                )
                total += len(returned) if returned is not None else 0
            except TypeError:
                execute_values(cur, sql, chunk, template=tpl, page_size=len(chunk))
                total += cur.rowcount  # count for this chunk (single statement)
        conn.commit()
    return total
    
def fetch_authorized_sender_emails(active_only: bool = True) -> List[str]:
    """
    Return a list of email addresses from the authorized_senders table.

    Args:
        active_only: If True (default), only rows with status = TRUE are returned.
                     If False, all rows are returned regardless of status.

    Returns:
        List[str]: deduplicated, trimmed email addresses ordered alphabetically.
    """
    where = "WHERE status IS TRUE" if active_only else ""
    sql = f"""
        SELECT email
        FROM authorized_senders
        {where}
        ORDER BY email ASC;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    # Trim, drop empties, and dedupe while preserving sort from SQL
    seen = set()
    emails: List[str] = []
    for (email,) in rows:
        if not email:
            continue
        e = email.strip()
        if e and e not in seen:
            seen.add(e)
            emails.append(e)

    return emails

# ____________ Ingesting file for date format review _________________


def insert_or_update_ingest_file(
    *,
    email_address: Optional[str],
    subject: Optional[str],
    received_at: Optional[datetime],
    processed_at: Optional[datetime],
    file_path: str,
    preview_cache: Optional[Dict[str, Any]] = None,
    error_message: Optional[str] = None,
    # new optional fields
    status: Optional[str] = None,
    date_format: Optional[str] = None,
    approved_at: Optional[datetime] = None,
    is_format_auto_detected: Optional[bool] = None,
) -> int:
    """
    Upsert a single row into ingest_files keyed by unique(file_path).

    Columns written:
      email_address, subject, received_at, processed_at, file_path,
      preview_cache, error_message,
      status, date_format, approved_at, is_format_auto_detected
    """
    if not file_path:
        raise ValueError("file_path is required")

    # ---- Defaults so CSVs don't violate NOT NULL ----
    # status = bool(status) if status is not None else False

    status = status if status not in (None, "") else "pending"

    is_format_auto_detected = bool(is_format_auto_detected) if is_format_auto_detected is not None else False
    date_format = (date_format or None)
    # approved_at may remain None

    sql = """
        INSERT INTO ingest_files (
            email_address,
            subject,
            received_at,
            processed_at,
            file_path,
            preview_cache,
            error_message,
            status,
            date_format,
            approved_at,
            is_format_auto_detected,
            created_at,
            updated_at
        )
        VALUES (
            %(email_address)s,
            %(subject)s,
            %(received_at)s,
            %(processed_at)s,
            %(file_path)s,
            %(preview_cache)s,
            %(error_message)s,
            %(status)s,
            %(date_format)s,
            %(approved_at)s,
            %(is_format_auto_detected)s,
            NOW(),
            NOW()
        )
        ON CONFLICT (file_path) DO UPDATE SET
            email_address           = EXCLUDED.email_address,
            subject                 = EXCLUDED.subject,
            received_at             = EXCLUDED.received_at,
            processed_at            = EXCLUDED.processed_at,
            preview_cache           = EXCLUDED.preview_cache,
            error_message           = EXCLUDED.error_message,
            status                  = EXCLUDED.status,
            date_format             = EXCLUDED.date_format,
            approved_at             = EXCLUDED.approved_at,
            is_format_auto_detected = EXCLUDED.is_format_auto_detected,
            updated_at              = NOW()
        RETURNING id;
    """

    params = {
        "email_address": email_address,
        "subject": subject,
        "received_at": received_at,
        "processed_at": processed_at,
        "file_path": file_path,
        "preview_cache": Json(preview_cache) if preview_cache is not None else None,
        "error_message": error_message,
        "status": status,
        "date_format": date_format,
        "approved_at": approved_at,
        "is_format_auto_detected": is_format_auto_detected,
    }

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, params)
        new_id = cur.fetchone()[0]
        conn.commit()
        return new_id



def bulk_upsert_ingest_files(
    rows: Iterable[Dict[str, Any]],
    page_size: int = 200,
) -> int:
    """
    Bulk upsert many files. Returns the number of rows affected.
    Each row may include the same keys as insert_or_update_ingest_file args.

    Requires a unique index on (file_path).
    """
    # Normalize and coerce preview_cache to Json for each row
    prepared: list[Tuple] = []
    for r in rows:
        fp = r.get("file_path")
        if not fp:
            # skip silent or raise; choose sanity
            raise ValueError("bulk_upsert_ingest_files: each row must include file_path")

        prepared.append((
            r.get("email_address"),
            r.get("subject"),
            r.get("received_at"),
            r.get("processed_at"),
            fp,
            Json(r.get("preview_cache")) if r.get("preview_cache") is not None else None,
            r.get("error_message"),
        ))

    if not prepared:
        return 0

    sql = """
        INSERT INTO ingest_files (
            email_address,
            subject,
            received_at,
            processed_at,
            file_path,
            preview_cache,
            error_message,
            created_at,
            updated_at
        )
        VALUES %s
        ON CONFLICT (file_path) DO UPDATE SET
            email_address = EXCLUDED.email_address,
            subject       = EXCLUDED.subject,
            received_at   = EXCLUDED.received_at,
            processed_at  = EXCLUDED.processed_at,
            preview_cache = EXCLUDED.preview_cache,
            error_message = EXCLUDED.error_message,
            updated_at    = NOW()
    """
    tpl = "(%s,%s,%s,%s,%s,%s,%s,NOW(),NOW())"

    with get_conn() as conn, conn.cursor() as cur:
        execute_values(cur, sql, prepared, template=tpl, page_size=page_size)
        affected = cur.rowcount  # number of rows the last statement claims it touched
        conn.commit()
        # Rowcount is fine as a lower-bound; ON CONFLICT may not reflect all changes precisely.
        return affected

def fetch_approved_unprocessed_paths_map(limit: Optional[int] = None) -> Dict[str, Optional[str]]:
    """
    Return a dict mapping file_path -> date_format from ingest_files where:
      - status ILIKE 'approved' (case-insensitive)
      - is_processed = FALSE (NULL treated as FALSE)
      - file_path is present
    Ordered by received_at (oldest first), then id.
    If duplicate file_paths exist, the earliest (by ordering) wins.
    """
    base_sql = """
        SELECT file_path, date_format
        FROM ingest_files
        WHERE LOWER(COALESCE(status, '')) = 'approved'
          AND COALESCE(is_processed, FALSE) = FALSE
          AND file_path IS NOT NULL
          AND file_path <> ''
        ORDER BY received_at NULLS LAST, id ASC
    """
    sql = base_sql + (" LIMIT %s" if limit is not None else "")

    with get_conn() as conn, conn.cursor() as cur:
        if limit is not None:
            cur.execute(sql, (limit,))
        else:
            cur.execute(sql)
        rows = cur.fetchall()

    out: Dict[str, Optional[str]] = {}
    for file_path, date_format in rows:
        if file_path not in out:  # keep the earliest one if duplicates
            out[file_path] = date_format
    return out

def mark_ingest_processed(file_paths: Iterable[str], processed: bool = True) -> int:
    """
    Bulk-mark ingest_files rows as processed/unprocessed by exact file_path match.

    Returns number of rows updated.
    """
    paths = [p for p in set(file_paths) if p]
    if not paths:
        return 0

    sql = """
        UPDATE ingest_files
           SET is_processed = %s,
               updated_at = NOW()
         WHERE file_path = ANY(%s)
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (processed, paths))
        updated = cur.rowcount
        conn.commit()
    return updated
from typing import Literal

ProcessingStage = Literal[
    "date_format_fetched",
    "jera_fetched",
    "file_cleaned",
    "rate_compared",
    "rate_uploaded",
]
# --- stage gating helpers ---
STAGE_ORDER: list[ProcessingStage] = [
    "date_format_fetched",
    "jera_fetched",
    "file_cleaned",
    "rate_compared",
    "rate_uploaded",
]

COL_MAP = {
    "date_format_fetched": "is_date_format_fetched",
    "jera_fetched":        "is_jera_fetched",
    "file_cleaned":        "is_file_cleaned",
    "rate_compared":       "is_rate_compared",
    "rate_uploaded":       "is_rate_uploaded",
}

def _prereq_cols_for(stage: ProcessingStage) -> list[str]:
    idx = STAGE_ORDER.index(stage)
    return [COL_MAP[s] for s in STAGE_ORDER[:idx]]

def upsert_processing_status(
    *,
    internet_message_id: str,
    directory_name: str,
    sender_email: Optional[str] = None,
    email_subject: Optional[str] = None,
    email_received_at: Optional[datetime] = None,
) -> int:
    """
    Create or update a processing_statuses row keyed by internet_message_id.
    Also enforces directory_name uniqueness. Safe to call many times.
    Returns the row id.
    """
    if not internet_message_id or not directory_name:
        raise ValueError("internet_message_id and directory_name are required")

    sql = """
    INSERT INTO processing_statuses (
        internet_message_id, directory_name, sender_email, email_subject, email_received_at,
        created_at, updated_at
    ) VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
    ON CONFLICT (internet_message_id) DO UPDATE SET
        directory_name    = EXCLUDED.directory_name,
        sender_email      = COALESCE(EXCLUDED.sender_email, processing_statuses.sender_email),
        email_subject     = COALESCE(EXCLUDED.email_subject, processing_statuses.email_subject),
        email_received_at = COALESCE(EXCLUDED.email_received_at, processing_statuses.email_received_at),
        updated_at        = NOW()
    RETURNING id;
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (internet_message_id, directory_name, sender_email, email_subject, email_received_at),
        )
        rid = cur.fetchone()[0]
        conn.commit()
        return rid


def ensure_row_by_directory(
    *,
    directory_name: str,
    internet_message_id: Optional[str] = None,
    sender_email: Optional[str] = None,
    email_subject: Optional[str] = None,
    email_received_at: Optional[datetime] = None,
) -> int:
    """
    Idempotent ensure by directory_name (handy when you don't yet know the message-id).
    If the row exists (by directory_name), optionally fills missing fields once.
    If it doesn't exist, requires internet_message_id to insert.
    Returns id.
    """
    if not directory_name:
        raise ValueError("directory_name is required")

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT id FROM processing_statuses WHERE directory_name = %s",
            (directory_name,),
        )
        row = cur.fetchone()
        if row:
            rid = row[0]
            # backfill missing fields once
            cur.execute(
                """
                UPDATE processing_statuses
                   SET internet_message_id = COALESCE(%s, internet_message_id),
                       sender_email        = COALESCE(%s, sender_email),
                       email_subject       = COALESCE(%s, email_subject),
                       email_received_at   = COALESCE(%s, email_received_at),
                       updated_at          = NOW()
                 WHERE id = %s
                """,
                (internet_message_id, sender_email, email_subject, email_received_at, rid),
            )
            conn.commit()
            return rid

        if not internet_message_id:
            raise ValueError("internet_message_id is required for first insert")
        # insert new
        cur.execute(
            """
            INSERT INTO processing_statuses
              (internet_message_id, directory_name, sender_email, email_subject, email_received_at,
               created_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
            RETURNING id
            """,
            (internet_message_id, directory_name, sender_email, email_subject, email_received_at),
        )
        rid = cur.fetchone()[0]
        conn.commit()
        return rid


def mark_processing_stage(
    *,
    directory_name: Optional[str] = None,
    internet_message_id: Optional[str] = None,
    stage: ProcessingStage,
    final_status: Optional[bool] = None,
) -> int:
    """
    Status rules:
      - Any stage set to FALSE  -> status='fail'   (no prereq gating)
      - Any stage set to TRUE (but not rate_uploaded) -> status='processing'
      - rate_uploaded set to TRUE -> status='success'
    """
    if not directory_name and not internet_message_id:
        raise ValueError("provide directory_name or internet_message_id")

    col = COL_MAP[stage]
    is_final_stage = (stage == "rate_uploaded")

    # Key WHERE
    if directory_name:
        where_key_sql, where_key_args = "directory_name = %s", (directory_name,)
    else:
        where_key_sql, where_key_args = "internet_message_id = %s", (internet_message_id,)

    set_bits = ["updated_at = NOW()"]

    # ---- Immediate failure (no gating) ----
    if final_status is False:
        set_bits += [f"{col} = FALSE", "status = 'fail'"]
        sql = f"UPDATE processing_statuses SET {', '.join(set_bits)} WHERE {where_key_sql}"
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql, where_key_args)
            affected = cur.rowcount
            conn.commit()
            return affected

    # ---- Advance (requires prereqs TRUE) ----
    prereq_cols = _prereq_cols_for(stage)
    prereq_sql = " AND ".join(f"{c} = TRUE" for c in prereq_cols) if prereq_cols else ""
    where_sql = where_key_sql + (f" AND {prereq_sql}" if prereq_sql else "")

    set_bits.append(f"{col} = TRUE")

    if is_final_stage:
        # Only when the final stage is marked TRUE do we mark success
        set_bits.append("status = 'success'")
    else:
        # While progressing through earlier stages, show 'processing'
        # but don't overwrite a terminal state if it's already there.
        set_bits.append(
            "status = CASE WHEN status IN ('fail','success') THEN status ELSE 'processing' END"
        )

    sql = f"UPDATE processing_statuses SET {', '.join(set_bits)} WHERE {where_sql}"
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, where_key_args)
        affected = cur.rowcount
        conn.commit()
        return affected




def get_processing_status(
    *, directory_name: Optional[str] = None, internet_message_id: Optional[str] = None
) -> Optional[dict]:
    """
    Fetch a processing_statuses row for inspection.
    """
    if not directory_name and not internet_message_id:
        raise ValueError("provide directory_name or internet_message_id")

    where = ("directory_name = %s", (directory_name,)) if directory_name else ("internet_message_id = %s", (internet_message_id,))
    sql = f"""
    SELECT id, internet_message_id, directory_name, sender_email, email_subject, email_received_at,
           is_date_format_fetched, is_jera_fetched, is_file_cleaned, is_rate_compared, is_rate_uploaded, status,
           created_at, updated_at
      FROM processing_statuses
     WHERE {where[0]}
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, where[1])
        row = cur.fetchone()
        if not row:
            return None
        keys = [
            "id","internet_message_id","directory_name","sender_email","email_subject","email_received_at",
            "is_date_format_fetched","is_jera_fetched","is_file_cleaned","is_rate_compared","is_rate_uploaded","status",
            "created_at","updated_at",
        ]
        return dict(zip(keys, row))
