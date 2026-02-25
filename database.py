import os
from dotenv import load_dotenv
import psycopg2
from typing import Iterable, Dict, Any, Optional, List, Mapping, Tuple
from datetime import datetime
from decimal import Decimal
import json
from psycopg2.extras import execute_values, Json 
from pathlib import Path
from typing import Optional
from typing import Optional, Tuple
from typing import Sequence


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
    internet_message_id: Optional[str] = None,
) -> int:
    """
    Insert a single rejected email row and return its id.
    Table columns (managed by DB): id (PK), created_at, updated_at auto.
    Also supports optional internet_message_id for traceability.
    """
    from valid_emails import get_verified_senders
    # Only insert if sender is authorized
    if sender_email and sender_email not in get_verified_senders():
        print(f"Skipping unauthorized sender: {sender_email}")
        return -1  # or handle as needed
    sql = """
        INSERT INTO rejected_emails
        (sender_email, subject, category, notes, received_at, processed_at, internet_message_id, created_at, updated_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, NOW(), NOW())
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (sender_email, subject, category, notes, received_at, processed_at, internet_message_id),
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
            - internet_message_id (Optional[str])

    Returns: list of inserted ids (may be empty).
    """
    # collect values in the order expected by the DB
    values = []
    from valid_emails import get_verified_senders
    for r in rows:
        # ensure category exists (raise so caller notices bad input)
        category = r.get("category")
        sender_email = r.get("sender_email")
        if category is None:
            raise ValueError("Each row must include a 'category' value.")
        # Only insert if sender is authorized
        if sender_email and sender_email not in get_verified_senders():
            print(f"Skipping unauthorized sender: {sender_email}")
            continue
        values.append((
            sender_email,
            r.get("subject"),
            str(category),
            r.get("notes"),
            r.get("received_at"),
            r.get("processed_at"),
            r.get("internet_message_id"),
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
        (sender_email, subject, category, notes, received_at, processed_at, internet_message_id, created_at, updated_at)
        VALUES %s
        RETURNING id;
    """

    with get_conn() as conn, conn.cursor() as cur:
        # execute_values will expand the VALUES %s placeholder into many tuples
        # Template to match the 9 columns (7 data + 2 timestamps)
        template = "(%s,%s,%s,%s,%s,%s,%s,NOW(),NOW())"
        execute_values(cur, sql, values, template=template, page_size=100)
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
    internet_message_id: Optional[str] = None,
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
                internet_message_id TEXT,
      )
    """
    sql = """
        INSERT INTO rejected_emails
                    (sender_email, subject, category, notes, received_at, processed_at, internet_message_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
                        sql,
                        (sender_email, subject, category, notes, received_at, processed_at, internet_message_id),
        )
        new_id = cur.fetchone()[0]
        conn.commit()
        return new_id


def upsert_rejected_email_from_processing_failure(
    *,
    stage: "ProcessingStage | str",
    status_text: str,
    internet_message_id: Optional[str],
    sender_email: Optional[str],
    email_subject: Optional[str],
    email_received_at: Optional[datetime],
) -> Optional[int]:
    """Update or insert a rejected_emails row for a failed processing_statuses entry.

    If a row already exists for this internet_message_id, update its
    category/notes/processed_at to reflect the latest failure reason.
    Otherwise insert a new rejected_emails row.
    """
    category = str(stage)
    notes = status_text
    processed_at = datetime.utcnow()

    # If we don't have a message id, just fall back to a simple insert.
    if not internet_message_id:
        try:
            return insert_rejected_email(
                sender_email=sender_email,
                subject=email_subject,
                category=category,
                notes=notes,
                received_at=email_received_at,
                processed_at=processed_at,
                internet_message_id=None,
            )
        except Exception as e:
            print(f"[rejected_emails] warn: failed to insert without internet_message_id: {e}")
            return None

    # First try to update any existing row(s) for this internet_message_id.
    try:
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(
                """
                UPDATE rejected_emails
                   SET category = %s,
                       notes = %s,
                       processed_at = %s,
                       updated_at = NOW()
                 WHERE internet_message_id = %s
                 RETURNING id;
                """,
                (category, notes, processed_at, internet_message_id),
            )
            row = cur.fetchone()
            if row:
                conn.commit()
                return row[0]
            conn.commit()
    except Exception as e:
        print(f"[rejected_emails] warn: failed to update by internet_message_id={internet_message_id}: {e}")

    # No existing row; insert a fresh one.
    try:
        return insert_rejected_email(
            sender_email=sender_email,
            subject=email_subject,
            category=category,
            notes=notes,
            received_at=email_received_at,
            processed_at=processed_at,
            internet_message_id=internet_message_id,
        )
    except Exception as e:
        print(f"[rejected_emails] warn: failed to insert after update miss for internet_message_id={internet_message_id}: {e}")
        return None

def push_failed_emails_json_to_db(path: Optional[str | Path] = None) -> tuple[int, int, int]:
    """
    Read failed_emails.json and insert any entries not yet pushed (no 'already_pushed': true)
    into the rejected_emails table.

    Entries are marked as already_pushed ONLY after a successful DB insert, so
    failures (or sender-authorization skips) can retry on later runs.

    Returns:
        (inserted_count, skipped_already_pushed, errors)
    """
    json_path = Path(path) if path else Path(__file__).with_name("failed_emails.json")
    if not json_path.exists():
        print(f"(info) {json_path} not found; nothing to push.")
        return (0, 0, 0)

    # Ensure rejected_emails has internet_message_id column (safe to call many times).
    try:
        with get_conn() as conn, conn.cursor() as cur:
            try:
                cur.execute("ALTER TABLE rejected_emails ADD COLUMN IF NOT EXISTS internet_message_id TEXT;")
                conn.commit()
            except Exception:
                conn.rollback()
    except Exception as e:
        print(f"(warn) could not ensure rejected_emails.internet_message_id column exists: {e}")

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

            # Internet message id, if available, for better traceability.
            raw_msg_id = entry.get("internetMessageId") or entry.get("internet_message_id")
            if isinstance(raw_msg_id, str):
                internet_message_id = raw_msg_id.strip() or None
            elif raw_msg_id is None:
                internet_message_id = None
            else:
                internet_message_id = str(raw_msg_id)

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
                "internet_message_id": internet_message_id,
            })
            entry_refs.append(entry)

    # Nothing to insert
    if not rows_to_insert:
        print(f"(info) no new failed emails to push. skipped={skipped}")
        return (0, skipped, errors)

    # ─────────────────────────────────────────────────────────
    # 1) Insert into DB (row-by-row)
    # ─────────────────────────────────────────────────────────
    # NOTE: mark already_pushed only after a successful DB insert.

    # ─────────────────────────────────────────────────────────
    # 2) Persist the updated JSON AFTER successful inserts
    #    If this fails, entries will retry on the next run.
    # ─────────────────────────────────────────────────────────
    marked = 0
    for row, entry in zip(rows_to_insert, entry_refs):
        try:
            new_id = insert_rejected_email(
                sender_email=row["sender_email"],
                subject=row["subject"],
                category=row["category"],
                notes=row["notes"],
                received_at=row["received_at"],
                processed_at=row["processed_at"],
                internet_message_id=row.get("internet_message_id"),
            )
            # insert_rejected_email returns -1 for unauthorized senders.
            if isinstance(new_id, int) and new_id > 0:
                inserted += 1
                entry["already_pushed"] = True
                marked += 1
        except Exception as e:
            errors += 1
            print(f"(warn) failed to insert rejected email (category={row.get('category')}): {e}")

    if marked:
        try:
            _atomic_write_json(json_path, data)
        except Exception as e:
            errors += 1
            print(f"(warn) failed to update {json_path} after DB insert: {e}")

    print(f"(ok) rejected_emails sync -> inserted={inserted}, skipped={skipped}, errors={errors}")
    return (inserted, skipped, errors)

    # ─────────────────────────────────────────────────────────
    # 3) Now safely insert into DB (at-most-once semantics)
    # ─────────────────────────────────────────────────────────
def insert_rate_upload(
    *,
    sender_email: Optional[str],
    subject: Optional[str],
    received_at: Optional[datetime] = None,
    processed_at: Optional[datetime] = None,
    totals: Optional[Dict[str, int]] = None,
    rates_gte_one_usd_count: int = 0,
    jera_table_id: Optional[int] = None,
    comparison_file_path: Optional[str] = None,
    internet_message_id: Optional[str] = None,
    processing_status_id: Optional[int] = None,
) -> int:
    """
    Insert one row into rate_uploads with summary counters.

        rate_uploads columns covered:
            subject, sender_email, received_at, processed_at, internet_message_id,
      total_rows, new, increase, decrease, unchanged, closed,
      backdated_increase, backdated_decrease, billing_increment_changes,
      comparison_file_path
    """
    t = {
        "total_rows": 0,
        "new": 0,
        "increase": 0,
        "decrease": 0,
        "unchanged": 0,
        "closed": 0,
        "stashed": 0,
        "backdated_increase": 0,
        "backdated_decrease": 0,
        "billing_increment_changes": 0,
    }
    if totals:
        t.update({k: int(v) for k, v in totals.items() if k in t})

    sql = """
        INSERT INTO rate_uploads
        (subject, sender_email, received_at, processed_at, internet_message_id, processing_status_id,
         total_rows, "new", increase, decrease, unchanged, closed, stashed,
         backdated_increase, backdated_decrease, billing_increment_changes,
         rates_gte_one_usd_count,
         jera_table_id, comparison_file_path, created_at, updated_at)
        VALUES
        (%s, %s, COALESCE(%s, NOW()), %s, %s, %s,
         %s, %s, %s, %s, %s, %s, %s,
         %s, %s, %s, %s, %s, %s,
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
                internet_message_id,
                processing_status_id,
                t["total_rows"],
                t["new"],
                t["increase"],
                t["decrease"],
                t["unchanged"],
                t["closed"],
                t["stashed"],
                t["backdated_increase"],
                t["backdated_decrease"],
                t["billing_increment_changes"],
                int(rates_gte_one_usd_count or 0),
                jera_table_id,
                comparison_file_path,
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


def fetch_rate_upload_details_for_upload(
    rate_upload_id: int,
    *,
    statuses: Optional[Sequence[str]] = ("Accepted",),
) -> List[Dict[str, Any]]:
    """
    Fetch rate_upload_details rows for a given upload, typically for pushing to JeraSoft.

    Notes:
    - This reads the CURRENT status stored in DB (what the UI/admin set), not the Excel file.
    - `statuses` is optional; pass None/empty to fetch all statuses.
    """
    if not rate_upload_id:
        return []

    where_status = ""
    params: list[Any] = [rate_upload_id]
    if statuses:
        # Case-sensitive match (only exact statuses like "Accepted" are included)
        where_status = "AND status = ANY(%s)"
        params.append([str(s).strip() for s in statuses if str(s).strip()])

    sql = f"""
        SELECT
            dst_code,
            code_name,
            rate_new,
            effective_date,
            new_billing_increment,
            status,
            change_type,
            notes
        FROM rate_upload_details
        WHERE rate_upload_id = %s
        {where_status}
        ORDER BY created_at ASC, dst_code ASC
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, tuple(params))
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]

    return [dict(zip(cols, r)) for r in rows]
    
def fetch_vendor_contact_emails(active_only: bool = True) -> List[str]:
    """
    Return a list of vendor contact emails from vendor_contacts (optionally only active ones),
    limited to contacts whose vendor is active as well.

    Tables (as per your DB):
      - vendor_contacts(email, vendor_id, status)
      - vendors(id, status)
    """
    where_bits = []
    if active_only:
        where_bits.append("COALESCE(vc.status, TRUE) = TRUE")
        where_bits.append("COALESCE(v.status, TRUE) = TRUE")
    where = ("WHERE " + " AND ".join(where_bits)) if where_bits else ""

    sql = f"""
        SELECT vc.email
        FROM vendor_contacts vc
        JOIN vendors v ON v.id = vc.vendor_id
        {where}
        ORDER BY vc.email ASC
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()

    seen = set()
    emails: List[str] = []
    for (email,) in rows:
        if not email:
            continue
        e = str(email).strip()
        if e and e not in seen:
            seen.add(e)
            emails.append(e)
    return emails


# ─────────────────────── Vendor date format lookup ───────────────────────
def fetch_vendor_context_by_sender_email(email: Optional[str]) -> Dict[str, Any]:
    """
    Resolve sender email -> vendor_id -> vendor date_format using:

      vendor_contacts.email -> vendor_contacts.vendor_id -> vendors.date_format

    Returns:
      {"vendor_id": Optional[int], "vendor_date_format": Optional[str]}
    """
    out: Dict[str, Any] = {"vendor_id": None, "vendor_date_format": None}
    if not email:
        return out
    e = str(email).strip()
    if not e:
        return out

    sql = """
        SELECT vc.vendor_id, v.date_format
        FROM vendor_contacts vc
        JOIN vendors v ON v.id = vc.vendor_id
        WHERE LOWER(vc.email) = LOWER(%s)
          AND COALESCE(vc.status, TRUE) = TRUE
          AND COALESCE(v.status, TRUE) = TRUE
        ORDER BY vc.id ASC
        LIMIT 1
    """
    try:
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(sql, (e,))
            row = cur.fetchone()
        if not row:
            return out
        vendor_id, fmt = row[0], row[1]
        out["vendor_id"] = int(vendor_id) if vendor_id is not None else None
        out["vendor_date_format"] = (str(fmt).strip() if fmt is not None else None) or None
        return out
    except Exception:
        return out


def fetch_vendor_date_format_by_sender_email(email: Optional[str]) -> Optional[str]:
    """Convenience wrapper: return vendor date_format (if any) for sender email."""
    ctx = fetch_vendor_context_by_sender_email(email)
    return ctx.get("vendor_date_format") or None


# ____________ Ingesting file for date format review _________________


def insert_or_update_ingest_file(
    *,
    email_address: Optional[str],
    subject: Optional[str],
    received_at: Optional[datetime],
    processed_at: Optional[datetime],
    file_path: str,
    internet_message_id: Optional[str] = None,
    processing_status_id: Optional[int] = None,
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
            internet_message_id,
            processing_status_id,
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
            internet_message_id,
            processing_status_id,
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
            %(internet_message_id)s,
            %(processing_status_id)s,
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
            internet_message_id     = EXCLUDED.internet_message_id,
            processing_status_id    = COALESCE(EXCLUDED.processing_status_id, ingest_files.processing_status_id),
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
        "internet_message_id": internet_message_id,
        "processing_status_id": processing_status_id,
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
            r.get("internet_message_id"),
            r.get("processing_status_id"),
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
            internet_message_id,
            processing_status_id,
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
            processing_status_id = COALESCE(EXCLUDED.processing_status_id, ingest_files.processing_status_id),
            preview_cache = EXCLUDED.preview_cache,
            error_message = EXCLUDED.error_message,
            updated_at    = NOW()
    """
    tpl = "(%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW(),NOW())"

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


def backfill_ingest_files_jera_table_from_metadata(limit: Optional[int] = None) -> int:
    """Create/populate ingest_files.jera_table from attachment metadata.

    - Ensures ingest_files has a TEXT column named jera_table.
        - For each row where jera_table is NULL/empty and file_path is set,
            and received_at is today or newer,
      reads the neighbouring metadata.json and copies the JeraSoft table name
      (best_table_name or force_jerasoft_table_name) into jera_table.

    Returns: number of rows updated.
    """
    # 1) Ensure the column exists (safe to call many times).
    with get_conn() as conn, conn.cursor() as cur:
        try:
            cur.execute("ALTER TABLE ingest_files ADD COLUMN IF NOT EXISTS jera_table TEXT;")
            conn.commit()
        except Exception:
            # For older PostgreSQL without IF NOT EXISTS, ignore errors if column exists.
            conn.rollback()

    # 2) Backfill jera_table from metadata.json next to file_path.
    updated = 0
    with get_conn() as conn, conn.cursor() as cur:
        base_sql = """
                        SELECT id, file_path
                            FROM ingest_files
                         WHERE (jera_table IS NULL OR jera_table = '')
                             AND file_path IS NOT NULL
                             AND file_path <> ''
                             AND received_at >= CURRENT_DATE
        """
        sql = base_sql + (" LIMIT %s" if limit is not None else "")

        if limit is not None:
            cur.execute(sql, (limit,))
        else:
            cur.execute(sql)

        rows = cur.fetchall()

        from pathlib import Path
        import json as _json

        for rid, file_path in rows:
            try:
                folder = Path(file_path).parent
                meta_path = folder / "metadata.json"
                if not meta_path.exists():
                    continue

                try:
                    with meta_path.open("r", encoding="utf-8") as f:
                        meta = _json.load(f) or {}
                except Exception:
                    continue

                table_name = meta.get("best_table_name") or meta.get("force_jerasoft_table_name")
                if not table_name:
                    continue

                cur.execute(
                    "UPDATE ingest_files SET jera_table = %s, updated_at = NOW() WHERE id = %s",
                    (str(table_name), rid),
                )
                if cur.rowcount > 0:
                    updated += cur.rowcount
            except Exception as e:
                print(f"⚠️ backfill_ingest_files_jera_table_from_metadata: id={rid} path={file_path} error={e}")

        conn.commit()

    if updated:
        print(f"🔄 Backfilled jera_table for {updated} ingest_files row(s) from metadata.json")
    else:
        print("ℹ️ No ingest_files rows needed jera_table backfill")

    return updated


def ensure_internet_message_id_links() -> None:
    """Ensure ingest_files and rate_uploads have internet_message_id columns and FKs.

    - Adds ingest_files.internet_message_id TEXT and a foreign key to
      processing_statuses(internet_message_id).
    - Adds rate_uploads.internet_message_id TEXT and a foreign key to
      processing_statuses(internet_message_id).

    Safe to call many times; errors from existing columns/constraints are ignored.
    """
    with get_conn() as conn, conn.cursor() as cur:
        # ingest_files.internet_message_id column
        try:
            cur.execute("ALTER TABLE ingest_files ADD COLUMN IF NOT EXISTS internet_message_id TEXT;")
            conn.commit()
        except Exception:
            conn.rollback()

        # ingest_files FK
        try:
            cur.execute(
                """
                ALTER TABLE ingest_files
                ADD CONSTRAINT ingest_files_internet_message_id_fkey
                FOREIGN KEY (internet_message_id)
                REFERENCES processing_statuses (internet_message_id)
                ON UPDATE CASCADE
                ON DELETE SET NULL;
                """
            )
            conn.commit()
        except Exception:
            conn.rollback()

        # rate_uploads.internet_message_id column
        try:
            cur.execute("ALTER TABLE rate_uploads ADD COLUMN IF NOT EXISTS internet_message_id TEXT;")
            conn.commit()
        except Exception:
            conn.rollback()

        # rate_uploads FK
        try:
            cur.execute(
                """
                ALTER TABLE rate_uploads
                ADD CONSTRAINT rate_uploads_internet_message_id_fkey
                FOREIGN KEY (internet_message_id)
                REFERENCES processing_statuses (internet_message_id)
                ON UPDATE CASCADE
                ON DELETE SET NULL;
                """
            )
            conn.commit()
        except Exception:
            conn.rollback()

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
    is_reprocessing_enabled: Optional[bool] = None,
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
        is_reprocessing_enabled, created_at, updated_at
    ) VALUES (%s, %s, %s, %s, %s, %s, NOW(), NOW())
    ON CONFLICT (internet_message_id) DO UPDATE SET
        directory_name           = EXCLUDED.directory_name,
        sender_email             = COALESCE(EXCLUDED.sender_email, processing_statuses.sender_email),
        email_subject            = COALESCE(EXCLUDED.email_subject, processing_statuses.email_subject),
        email_received_at        = COALESCE(EXCLUDED.email_received_at, processing_statuses.email_received_at),
        is_reprocessing_enabled  = COALESCE(EXCLUDED.is_reprocessing_enabled, processing_statuses.is_reprocessing_enabled),
        updated_at               = NOW()
    RETURNING id;
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            sql,
            (internet_message_id, directory_name, sender_email, email_subject, email_received_at, is_reprocessing_enabled),
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
    is_reprocessing_enabled: Optional[bool] = None,
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
                   SET internet_message_id      = COALESCE(%s, internet_message_id),
                       sender_email             = COALESCE(%s, sender_email),
                       email_subject            = COALESCE(%s, email_subject),
                       email_received_at        = COALESCE(%s, email_received_at),
                       is_reprocessing_enabled  = COALESCE(%s, is_reprocessing_enabled),
                       updated_at               = NOW()
                 WHERE id = %s
                """,
                (internet_message_id, sender_email, email_subject, email_received_at, is_reprocessing_enabled, rid),
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
               is_reprocessing_enabled, created_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, NOW(), NOW())
            RETURNING id
            """,
            (internet_message_id, directory_name, sender_email, email_subject, email_received_at, is_reprocessing_enabled),
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
    error_message: Optional[str] = None,
) -> int:
    """
    Status rules:
      - Any stage set to FALSE  -> status='failed'   (no prereq gating)
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
        # Build a human-readable status text but always keep a 'failed' prefix
        # so reprocessing queries can still match with status LIKE 'failed%'.
        status_text = "failed"
        if error_message:
            msg = str(error_message).strip()
            if msg:
                # Avoid unbounded row size; keep the most important part.
                if len(msg) > 900:
                    msg = msg[:900] + "...(truncated)"
                status_text = f"failed: {msg}"

        set_bits += [f"{col} = FALSE", "status = %s"]
        sql = f"UPDATE processing_statuses SET {', '.join(set_bits)} WHERE {where_key_sql}"
        # Parameter order: first status_text for SET, then WHERE key(s).
        args = (status_text,) + where_key_args
        with get_conn() as conn, conn.cursor() as cur:
            # First, update processing_statuses
            cur.execute(sql, args)
            affected = cur.rowcount

            # If at least one row was marked failed, also sync to rejected_emails
            if affected > 0:
                # Fetch minimal context for the failed row
                cur.execute(
                    f"SELECT internet_message_id, sender_email, email_subject, email_received_at FROM processing_statuses WHERE {where_key_sql}",
                    where_key_args,
                )
                row = cur.fetchone()
                if row:
                    internet_message_id, sender_email, email_subject, email_received_at = row
                    try:
                        upsert_rejected_email_from_processing_failure(
                            stage=stage,
                            status_text=status_text,
                            internet_message_id=internet_message_id,
                            sender_email=sender_email,
                            email_subject=email_subject,
                            email_received_at=email_received_at,
                        )
                    except Exception as re_exc:
                        # Don't break main flow if rejected_emails sync fails
                        print(f"[processing_statuses] warn: failed to sync rejected_emails for {where_key_sql}: {re_exc}")

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
            "status = CASE WHEN status IN ('failed','success') THEN status ELSE 'processing' END"
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
           is_reprocessing_enabled, created_at, updated_at
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
            "is_reprocessing_enabled","created_at","updated_at",
        ]
        return dict(zip(keys, row))

def update_reprocessing_enabled(
    *,
    directory_name: Optional[str] = None,
    internet_message_id: Optional[str] = None,
    is_reprocessing_enabled: bool,
) -> int:
    """
    Update the is_reprocessing_enabled flag for a specific processing_statuses row.
    
    Args:
        directory_name: Directory name to identify the row
        internet_message_id: Message ID to identify the row
        is_reprocessing_enabled: New value for the reprocessing flag
        
    Returns:
        Number of rows affected (should be 1 if successful, 0 if row not found)
    """
    if not directory_name and not internet_message_id:
        raise ValueError("provide directory_name or internet_message_id")

    if directory_name:
        where_sql = "directory_name = %s"
        where_args = (is_reprocessing_enabled, directory_name)
    else:
        where_sql = "internet_message_id = %s"
        where_args = (is_reprocessing_enabled, internet_message_id)

    sql = f"""
        UPDATE processing_statuses 
        SET is_reprocessing_enabled = %s, updated_at = NOW()
        WHERE {where_sql}
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, where_args)
        affected = cur.rowcount
        conn.commit()
        return affected

def get_reprocessing_enabled_directories(limit: Optional[int] = None) -> List[str]:
    """
    Get a list of directory names where reprocessing is enabled AND status is failed.
    Only shows reprocessing option for failed emails, not successful ones.
    
    Args:
        limit: Optional limit on number of results
        
    Returns:
        List of directory names with reprocessing enabled and failed status
    """
    base_sql = """
        SELECT directory_name
        FROM processing_statuses
                WHERE is_reprocessing_enabled = TRUE
                    AND directory_name IS NOT NULL
                    AND status LIKE 'failed%'
        ORDER BY updated_at DESC
    """
    
    sql = base_sql + (" LIMIT %s" if limit is not None else "")
    
    with get_conn() as conn, conn.cursor() as cur:
        if limit is not None:
            cur.execute(sql, (limit,))
        else:
            cur.execute(sql)
        rows = cur.fetchall()
    
    return [row[0] for row in rows if row[0]]

def get_failed_directories_for_reprocessing(limit: Optional[int] = None) -> List[Dict[str, Any]]:
    """
    Get failed processing status records that are eligible for reprocessing.
    Only returns failed emails (not successful ones).
    
    Args:
        limit: Optional limit on number of results
        
    Returns:
        List of dictionaries with processing status information for failed emails
    """
    base_sql = """
        SELECT id, internet_message_id, directory_name, sender_email, email_subject, 
               email_received_at, status, is_reprocessing_enabled,
               created_at, updated_at
                FROM processing_statuses
                WHERE status LIKE 'failed%'
                    AND directory_name IS NOT NULL
        ORDER BY updated_at DESC
    """
    
    sql = base_sql + (" LIMIT %s" if limit is not None else "")
    
    with get_conn() as conn, conn.cursor() as cur:
        if limit is not None:
            cur.execute(sql, (limit,))
        else:
            cur.execute(sql)
        rows = cur.fetchall()
    
    keys = [
        "id", "internet_message_id", "directory_name", "sender_email", "email_subject",
        "email_received_at", "status", "is_reprocessing_enabled", "created_at", "updated_at"
    ]
    
    return [dict(zip(keys, row)) for row in rows]

# ===== Invalid Subject helpers =====

def get_or_create_invalid_subject(email: str,
                                  received_at: Optional[datetime] = None,
                                  processed_at: Optional[datetime] = None,
                                  status: str = "pending") -> int:
    """
    Ensure there is exactly one invalid_subjects row for this email.
    Returns its id.
    """
    sel = "SELECT id FROM invalid_subjects WHERE email = %s"
    ins = """
        INSERT INTO invalid_subjects (email, received_at, processed_at, status, created_at, updated_at)
        VALUES (%s, %s, %s, %s, NOW(), NOW())
        ON CONFLICT (email) DO UPDATE SET
          received_at = COALESCE(EXCLUDED.received_at, invalid_subjects.received_at),
          processed_at = COALESCE(EXCLUDED.processed_at, invalid_subjects.processed_at),
          updated_at = NOW()
        RETURNING id
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sel, (email,))
        row = cur.fetchone()
        if row:
            return int(row[0])
        cur.execute(ins, (email, received_at, processed_at, status))
        rid = cur.fetchone()[0]
        conn.commit()
        return int(rid)

# def insert_invalid_subject_detail(invalid_subject_id: int,
#                                   subject: str,
#                                   jera_table: Optional[str] = None) -> int:
#     sql = """
#         INSERT INTO invalid_subject_details (invalid_subject_id, subject, jera_table, created_at)
#         VALUES (%s, %s, %s, NOW())
#         (invalid_subject_id, subject) DO UPDATE SET
#           jera_table = COALESCE(EXCLUDED.jera_table, invalid_subject_details.jera_table)
#         RETURNING id
#     """
#     with get_conn() as conn, conn.cursor() as cur:
#         cur.execute(sql, (invalid_subject_id, subject, jera_table))
#         rid = cur.fetchone()[0]
#         conn.commit()
#         return int(rid)

def insert_invalid_subject_detail(
    invalid_subject_id: int,
    subject: str,
    jera_table: Optional[str] = None
) -> int:
    sql = """
        INSERT INTO invalid_subject_details (invalid_subject_id, subject, jera_table, created_at)
        VALUES (%s, %s, %s, NOW())
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (invalid_subject_id, subject, jera_table))
        rid = cur.fetchone()[0]
        conn.commit()
        return int(rid)


def find_invalid_subject_detail(invalid_subject_id: int, subject: str) -> Optional[Tuple[int, Optional[str]]]:
    """
    Return (detail_id, jera_table) if present for this subject; else None.
    """
    sql = "SELECT id, jera_table FROM invalid_subject_details WHERE invalid_subject_id = %s AND subject = %s"
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (invalid_subject_id, subject))
        row = cur.fetchone()
        if row:
            return int(row[0]), (row[1] if row[1] is not None else None)
        return None


# ─────────────────────── JERASOFT UPLOAD CONTROL ───────────────────────

def set_jera_upload_flag(rate_upload_id: int, is_rate_approved_by_admin: bool, jera_table_id: Optional[int] = None) -> None:
    """
    Set the is_rate_approved_by_admin flag for a rate_upload record.
    This acts as the 'button' to control JeraSoft uploads.
    
    Args:
        rate_upload_id: ID from rate_uploads table
        is_rate_approved_by_admin: True to enable JeraSoft upload, False to disable
        jera_table_id: Optional JeraSoft table ID for this upload
    """
    # Set proper status for bulk uploads
    upload_status = 'pending_bulk' if is_rate_approved_by_admin and jera_table_id else None
    
    sql = """
        UPDATE rate_uploads 
        SET is_rate_approved_by_admin = %s,
            jera_table_id = COALESCE(%s, jera_table_id),
            jera_upload_status = COALESCE(%s, jera_upload_status),
            updated_at = NOW()
        WHERE id = %s
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (is_rate_approved_by_admin, jera_table_id, upload_status, rate_upload_id))
        conn.commit()
        print(f"🔄 Updated upload {rate_upload_id}: is_rate_approved_by_admin={is_rate_approved_by_admin}, table_id={jera_table_id}, status={upload_status}")
        
        if cur.rowcount > 0:
            action = "enabled" if is_rate_approved_by_admin else "disabled"
            print(f"✅ JeraSoft upload {action} for rate_upload_id {rate_upload_id}")
        else:
            print(f"⚠️ No rate_upload found with id {rate_upload_id}")

def get_pending_jera_uploads() -> List[Dict[str, Any]]:
    """
    Get all rate_uploads that are flagged for JeraSoft upload but not yet uploaded.
    
    Returns:
        List of dictionaries with rate_upload data ready for JeraSoft upload
    """
    sql = """
        SELECT 
            id, subject, sender_email, jera_table_id,
            processed_at, total_rows,
            created_at, updated_at
        FROM rate_uploads 
        WHERE is_rate_approved_by_admin = TRUE 
        AND jera_upload_status IN ('pending', 'failed')
        ORDER BY created_at ASC
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()
        
        columns = [desc[0] for desc in cur.description]
        return [dict(zip(columns, row)) for row in rows]

def update_jera_upload_status(rate_upload_id: int, status: str, 
                             result: Optional[Dict] = None) -> None:
    """
    Update the JeraSoft upload status and result for a rate_upload.
    
    Args:
        rate_upload_id: ID from rate_uploads table
        status: 'pending', 'uploading', 'success', 'failed'
        result: Optional result dictionary from JeraSoft upload
    """
    # Only store error messages as string when status is failed, otherwise null
    error_message = None
    if status == 'failed' and result:
        if isinstance(result, dict):
            # Extract error message from result dict
            error_message = result.get('error', str(result))
        else:
            error_message = str(result)
    
    sql = """
        UPDATE rate_uploads 
        SET jera_upload_status = %s,
            jera_upload_result = %s,
            jera_uploaded_at = CASE WHEN %s = 'success' THEN NOW() ELSE jera_uploaded_at END,
            updated_at = NOW()
        WHERE id = %s
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (status, error_message, status, rate_upload_id))
        conn.commit()
        
        if status == 'success':
            print(f"✅ JeraSoft upload SUCCESS for rate_upload_id {rate_upload_id}")
        elif status == 'failed':
            print(f"❌ JeraSoft upload FAILED for rate_upload_id {rate_upload_id}: {error_message}")
        else:
            print(f"📊 JeraSoft upload status updated to '{status}' for rate_upload_id {rate_upload_id}")

def get_jera_upload_history(limit: int = 50) -> List[Dict[str, Any]]:
    """
    Get history of JeraSoft uploads with their status and results.
    
    Args:
        limit: Maximum number of records to return
        
    Returns:
        List of upload history records
    """
    sql = """
        SELECT 
            id, subject, sender_email, jera_table_id,
            is_rate_approved_by_admin, jera_upload_status,
            jera_upload_result, jera_uploaded_at,
            total_rows, processed_at, created_at
        FROM rate_uploads 
        WHERE is_rate_approved_by_admin = TRUE
        ORDER BY created_at DESC
        LIMIT %s
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (limit,))
        rows = cur.fetchall()
        
        columns = [desc[0] for desc in cur.description]
        return [dict(zip(columns, row)) for row in rows]

def mark_rate_upload_for_jera(subject: str, sender_email: str, jera_table_id: int) -> Optional[int]:
    """
    Find and mark a rate_upload for JeraSoft upload based on subject and sender.
    
    Args:
        subject: Email subject to match
        sender_email: Sender email to match  
        jera_table_id: JeraSoft table ID for upload
        
    Returns:
        rate_upload_id if found and marked, None otherwise
    """
    # First find the rate_upload
    find_sql = """
        SELECT id FROM rate_uploads 
        WHERE subject = %s AND sender_email = %s 
        ORDER BY created_at DESC 
        LIMIT 1
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(find_sql, (subject, sender_email))
        row = cur.fetchone()
        
        if row:
            rate_upload_id = row[0]
            set_jera_upload_flag(rate_upload_id, True, jera_table_id)
            return rate_upload_id
        else:
            print(f"⚠️ No rate_upload found for subject '{subject}' from '{sender_email}'")
            return None

def mark_comparison_file_for_bulk_upload(comparison_file_path: str, 
                                       subject: str, 
                                       sender_email: str, 
                                       jera_table_id: int) -> Optional[int]:
    """
    Mark a comparison file for BULK upload to JeraSoft.
    This stores the file path for later bulk processing instead of individual rates.
    
    Args:
        comparison_file_path: Absolute path to comparison result file
        subject: Email subject to match
        sender_email: Sender email to match
        jera_table_id: JeraSoft table ID for upload
        
    Returns:
        rate_upload_id if marked successfully, None otherwise
    """
    # Find the rate_upload record
    find_sql = """
        SELECT id FROM rate_uploads 
        WHERE subject = %s AND sender_email = %s 
        ORDER BY created_at DESC 
        LIMIT 1
    """
    
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(find_sql, (subject, sender_email))
        row = cur.fetchone()
        
        if row:
            rate_upload_id = row[0]
            
            # Update with bulk upload info
            update_sql = """
                UPDATE rate_uploads SET
                    is_rate_approved_by_admin = TRUE,
                    jera_table_id = %s,
                    jera_upload_status = 'pending_bulk',
                    jera_upload_result = %s,
                    updated_at = NOW()
                WHERE id = %s
            """
            
            # Store file path and upload method in result JSON
            upload_info = {
                "upload_method": "bulk_file_upload",
                "comparison_file": comparison_file_path,
                "marked_at": datetime.now().isoformat(),
                "status": "queued_for_bulk_upload"
            }
            
            cur.execute(update_sql, (jera_table_id, Json(upload_info), rate_upload_id))
            conn.commit()
            
            print(f"✅ Marked rate_upload_id {rate_upload_id} for BULK upload")
            print(f"📁 File: {comparison_file_path}")
            print(f"🎯 Target table: {jera_table_id}")
            
            return rate_upload_id
        else:
            print(f"⚠️ No rate_upload found for subject '{subject}' from '{sender_email}'")
            return None

def auto_update_status_on_import_flag_change():
    """
    Auto-update jera_upload_status when is_rate_approved_by_admin is manually changed to TRUE.
    Also sets jera_table_id from attachment metadata if missing.
    This should be called periodically (e.g., by cron job every minute).
    """
    
    # First handle records that already have table_id
    sql_with_table = """
        UPDATE rate_uploads 
        SET jera_upload_status = 'pending_bulk',
            updated_at = NOW()
        WHERE is_rate_approved_by_admin = TRUE 
        AND jera_upload_status = 'pending'
        AND jera_table_id IS NOT NULL
    """
    
    # Then find records that need table_id from metadata
    sql_find_missing = """
        SELECT id, subject, sender_email 
        FROM rate_uploads 
        WHERE is_rate_approved_by_admin = TRUE 
        AND jera_upload_status = 'pending'
        AND jera_table_id IS NULL
    """
    
    try:
        with get_conn() as conn, conn.cursor() as cur:
            # First update records that already have table_id
            cur.execute(sql_with_table)
            rows_updated_with_table = cur.rowcount
            
            # Find records missing table_id  
            cur.execute(sql_find_missing)
            missing_table_records = cur.fetchall()
            
            rows_updated_missing = 0
            
            # For each record missing table_id, try to find it from metadata
            for record_id, subject, sender_email in missing_table_records:
                try:
                    # Find attachment folder by sender email pattern
                    import os
                    import json
                    from pathlib import Path
                    
                    sender_pattern = sender_email.replace("@", "_at_").replace(".", "_")
                    attachments_dir = Path("attachments")
                    
                    table_id_found = None
                    
                    if attachments_dir.exists():
                        for folder in attachments_dir.iterdir():
                            if folder.is_dir() and sender_pattern in folder.name:
                                metadata_file = folder / "metadata.json"
                                if metadata_file.exists():
                                    try:
                                        with open(metadata_file) as f:
                                            meta = json.load(f)
                                            table_id_found = meta.get("table_id")
                                            if table_id_found:
                                                break
                                    except Exception:
                                        continue
                    
                    if table_id_found:
                        # Update both table_id and status
                        update_sql = """
                            UPDATE rate_uploads 
                            SET jera_table_id = %s,
                                jera_upload_status = 'pending_bulk',
                                updated_at = NOW()
                            WHERE id = %s
                        """
                        cur.execute(update_sql, (table_id_found, record_id))
                        rows_updated_missing += 1
                        print(f"   🔧 ID {record_id}: Set table_id={table_id_found} and status=pending_bulk")
                    else:
                        print(f"   ⚠️ ID {record_id}: Could not find table_id in metadata")
                        
                except Exception as e:
                    print(f"   ❌ ID {record_id}: Error finding table_id - {e}")
            
            conn.commit()
            total_updated = rows_updated_with_table + rows_updated_missing
            
            if total_updated > 0:
                print(f"🔄 Auto-updated {total_updated} record(s) from 'pending' to 'pending_bulk'")
                print(f"   - With existing table_id: {rows_updated_with_table}")  
                print(f"   - Found table_id from metadata: {rows_updated_missing}")
            
            return total_updated
            
    except Exception as e:
        print(f"❌ Error auto-updating status: {e}")
        return 0


# ─────────────────────── Vendor/day pending-upload gating ───────────────────────

WAITING_PREVIOUS_VENDOR_PENDING = "waiting: previous JeraSoft upload for this vendor today is still pending"

def vendor_has_pending_jera_upload_today(
    *,
    sender_email: Optional[str],
    exclude_internet_message_id: Optional[str] = None,
) -> bool:
    """
    Return True if there exists ANOTHER rate_uploads row for the same sender_email
    dated today (by received_at if present, else created_at) with a pending JeraSoft upload.

    Pending statuses (per spec): 'pending' or 'pending_bulk'

    Notes:
    - We exclude the current message id (when provided) to avoid blocking on itself during retries.
    - We compare sender_email case-insensitively.
    """
    email = (sender_email or "").strip()
    if not email:
        return False

    # Exclude current item by internet_message_id (best key across retries).
    exclude_mid = (exclude_internet_message_id or "").strip() or None

    sql = """
        SELECT 1
          FROM rate_uploads ru
         WHERE LOWER(COALESCE(ru.sender_email, '')) = LOWER(%s)
           AND COALESCE(ru.received_at, ru.created_at)::date = CURRENT_DATE
           AND ru.jera_upload_status IN ('pending', 'pending_bulk')
           AND (%s IS NULL OR COALESCE(ru.internet_message_id, '') <> %s)
         LIMIT 1
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (email, exclude_mid, exclude_mid))
        row = cur.fetchone()
        conn.commit()
        return bool(row)


def set_processing_status_text(
    *,
    directory_name: Optional[str] = None,
    internet_message_id: Optional[str] = None,
    status_text: str,
    # Never overwrite success/failed by default (per spec: don't clobber terminal statuses).
    overwrite_terminal: bool = False,
) -> int:
    """
    Update processing_statuses.status to a specific text (e.g., waiting reason).

    By default this will NOT overwrite terminal statuses:
      - status = 'success'
      - status LIKE 'failed%'

    Returns number of rows affected.
    """
    if not directory_name and not internet_message_id:
        raise ValueError("provide directory_name or internet_message_id")

    status_text = (status_text or "").strip()
    if not status_text:
        raise ValueError("status_text is required")

    if directory_name:
        where_key_sql, where_key_args = "directory_name = %s", (directory_name,)
    else:
        where_key_sql, where_key_args = "internet_message_id = %s", (internet_message_id,)

    terminal_guard_sql = ""
    if not overwrite_terminal:
        # IMPORTANT (psycopg2): literal '%' must be escaped as '%%' when query has parameters,
        # otherwise psycopg2 will treat it like a placeholder and raise "tuple index out of range".
        terminal_guard_sql = " AND COALESCE(status, '') <> 'success' AND COALESCE(status, '') NOT LIKE 'failed%%'"

    sql = f"""
        UPDATE processing_statuses
           SET status = %s,
               updated_at = NOW()
         WHERE {where_key_sql}
               {terminal_guard_sql}
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (status_text,) + where_key_args)
        affected = cur.rowcount
        conn.commit()
        return affected