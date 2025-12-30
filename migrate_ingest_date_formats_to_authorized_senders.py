#!/usr/bin/env python3
"""
One-time migration: copy date formats from ingest_files -> authorized_senders.

Use-case:
  You already have per-file date formats saved in `ingest_files` (typically via UI),
  and you want to populate `authorized_senders.date_format` per sender email.

Defaults:
  - Only uses ingest_files rows with status='approved' (case-insensitive)
  - Picks the latest approved row per email (approved_at desc, then updated_at/id)
  - Does NOT overwrite existing authorized_senders.date_format unless --overwrite is passed

Example:
  python tools/migrate_ingest_date_formats_to_authorized_senders.py --dry-run
  python tools/migrate_ingest_date_formats_to_authorized_senders.py
  python tools/migrate_ingest_date_formats_to_authorized_senders.py --overwrite
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from psycopg2.extras import execute_values

# Ensure project root is importable when running `python tools/<script>.py`
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import get_conn


def _has_column(cur, table: str, column: str, schema: str = "public") -> bool:
    cur.execute(
        """
        SELECT 1
        FROM information_schema.columns
        WHERE table_schema = %s
          AND table_name = %s
          AND column_name = %s
        """,
        (schema, table, column),
    )
    return cur.fetchone() is not None


def fetch_latest_approved_formats(cur, *, status: str = "approved") -> List[Tuple[str, str]]:
    """
    Returns list of (email, date_format) from ingest_files.
    """
    cur.execute(
        """
        SELECT DISTINCT ON (LOWER(BTRIM(email_address)))
               LOWER(BTRIM(email_address)) AS email,
               BTRIM(date_format)          AS date_format
        FROM ingest_files
        WHERE email_address IS NOT NULL
          AND BTRIM(email_address) <> ''
          AND date_format IS NOT NULL
          AND BTRIM(date_format) <> ''
          AND LOWER(COALESCE(status, '')) = LOWER(%s)
        ORDER BY LOWER(BTRIM(email_address)),
                 approved_at DESC NULLS LAST,
                 updated_at  DESC NULLS LAST,
                 id          DESC
        """,
        (status,),
    )
    return [(r[0], r[1]) for r in cur.fetchall()]


def fetch_authorized_sender_formats(cur) -> Dict[str, Optional[str]]:
    """
    Returns dict lower(email) -> date_format (trimmed or None).
    """
    cur.execute(
        """
        SELECT LOWER(BTRIM(email)) AS email,
               CASE
                 WHEN date_format IS NULL THEN NULL
                 ELSE NULLIF(BTRIM(date_format), '')
               END AS date_format
        FROM authorized_senders
        """
    )
    out: Dict[str, Optional[str]] = {}
    for email, fmt in cur.fetchall():
        out[str(email)] = (str(fmt).strip() if fmt is not None else None) or None
    return out


def _is_missing_or_auto(fmt: Optional[str]) -> bool:
    if fmt is None:
        return True
    s = str(fmt).strip()
    return (not s) or (s.lower() == "auto")


def upsert_authorized_sender_formats(
    cur,
    rows: List[Tuple[str, str]],
    *,
    overwrite: bool,
) -> int:
    """
    Upsert date_format into authorized_senders for the given rows.
    Returns number of statement-affected rows as reported by psycopg2.
    """
    if not rows:
        return 0

    # Insert missing senders with status=true; update existing date_format.
    where_clause = "" if overwrite else "WHERE authorized_senders.date_format IS NULL OR BTRIM(authorized_senders.date_format) = '' OR LOWER(BTRIM(authorized_senders.date_format)) = 'auto'"
    sql = f"""
        INSERT INTO authorized_senders (email, status, date_format, created_at, updated_at)
        VALUES %s
        ON CONFLICT (email) DO UPDATE
            SET date_format = EXCLUDED.date_format,
                updated_at  = NOW()
            {where_clause}
    """
    values = [(email, True, date_format) for (email, date_format) in rows]
    execute_values(cur, sql, values, template="(%s, %s, %s, NOW(), NOW())")
    return int(cur.rowcount or 0)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="Print what would change; do not write.")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite existing authorized_senders.date_format.")
    ap.add_argument("--status", default="approved", help="Ingest status to use (default: approved).")
    args = ap.parse_args()

    with get_conn() as conn, conn.cursor() as cur:
        if not _has_column(cur, "authorized_senders", "date_format"):
            raise SystemExit("Missing column: public.authorized_senders.date_format (add it before running).")

        src = fetch_latest_approved_formats(cur, status=args.status)
        if not src:
            print("No approved ingest_files rows with date_format found. Nothing to migrate.")
            return 0

        existing = fetch_authorized_sender_formats(cur)
        to_apply: List[Tuple[str, str]] = []
        for email, fmt in src:
            cur_fmt = existing.get(email)
            if args.overwrite or _is_missing_or_auto(cur_fmt):
                to_apply.append((email, fmt))

        print(f"Found {len(src)} sender(s) with approved date_format in ingest_files.")
        print(f"Will apply {len(to_apply)} update(s) to authorized_senders (overwrite={args.overwrite}).")

        if args.dry_run:
            for email, fmt in to_apply[:50]:
                print(f"  - {email} -> {fmt}")
            if len(to_apply) > 50:
                print(f"  ... and {len(to_apply) - 50} more")
            return 0

        affected = upsert_authorized_sender_formats(cur, to_apply, overwrite=args.overwrite)
        conn.commit()
        print(f"Done. DB reported affected rows: {affected}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())


