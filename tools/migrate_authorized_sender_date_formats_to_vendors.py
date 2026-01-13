#!/usr/bin/env python3
"""
One-time migration: copy date formats from authorized_senders -> vendors via vendor_contacts.

Use-case:
  You imported vendors + vendor_contacts (company -> emails).
  You already have per-sender date formats in authorized_senders.date_format.
  Now you want to populate vendors.date_format per company based on those sender formats.

Mapping logic:
  authorized_senders.email  (sender)  --(case-insensitive match)--> vendor_contacts.email
  vendor_contacts.vendor_id -------------------------------> vendors.id

Defaults:
  - Only uses authorized_senders rows where date_format is not null/blank and not 'AUTO'
  - Picks ONE format per vendor:
      latest sender updated_at DESC (falls back to vendor_contact email ordering)
  - Does NOT overwrite existing vendors.date_format unless --overwrite is passed

Safety:
  - --dry-run prints planned updates + warnings (no DB writes)
  - Warns if a vendor has multiple distinct date_formats among its contacts

Examples:
  python tools/migrate_authorized_sender_date_formats_to_vendors.py --dry-run
  python tools/migrate_authorized_sender_date_formats_to_vendors.py --apply
  python tools/migrate_authorized_sender_date_formats_to_vendors.py --apply --overwrite
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from psycopg2.extras import execute_values

# Ensure project root is importable when running from tools/
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import get_conn  # noqa: E402


def _has_table(cur, table: str, schema: str = "public") -> bool:
    cur.execute(
        """
        SELECT 1
        FROM information_schema.tables
        WHERE table_schema = %s
          AND table_name = %s
        """,
        (schema, table),
    )
    return cur.fetchone() is not None


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


def _is_missing_or_auto(fmt: Optional[str]) -> bool:
    if fmt is None:
        return True
    s = str(fmt).strip()
    return (not s) or (s.lower() == "auto")


def fetch_candidate_rows(cur) -> List[Tuple[int, str, str, str, Optional[str]]]:
    """
    Returns rows:
      (vendor_id, company_name, contact_email, date_format, sender_updated_at_iso)
    """
    # NOTE: We assume authorized_senders has updated_at; if not, we still work (NULLs last).
    cur.execute(
        """
        SELECT
            v.id AS vendor_id,
            v.company_name,
            LOWER(BTRIM(vc.email)) AS contact_email,
            BTRIM(a.date_format) AS date_format,
            COALESCE(a.updated_at::text, NULL) AS sender_updated_at
        FROM vendor_contacts vc
        JOIN vendors v ON v.id = vc.vendor_id
        JOIN authorized_senders a ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        WHERE a.date_format IS NOT NULL
          AND BTRIM(a.date_format) <> ''
          AND LOWER(BTRIM(a.date_format)) <> 'auto'
        """
    )
    return [(int(r[0]), str(r[1]), str(r[2]), str(r[3]), (str(r[4]) if r[4] is not None else None)) for r in cur.fetchall()]


def choose_one_format_per_vendor(rows: List[Tuple[int, str, str, str, Optional[str]]]) -> Tuple[Dict[int, Tuple[str, str]], Dict[int, List[str]]]:
    """
    Build:
      chosen[vendor_id] = (date_format, chosen_email)
      conflicts[vendor_id] = sorted distinct date_formats (len>1)
    Strategy:
      - prefer row with newest sender_updated_at (lexicographic is fine for ISO text)
      - tie-break by email asc
    """
    by_vendor: Dict[int, List[Tuple[str, str, Optional[str]]]] = {}
    for vid, _company, email, fmt, updated_at in rows:
        by_vendor.setdefault(vid, []).append((fmt, email, updated_at))

    chosen: Dict[int, Tuple[str, str]] = {}
    conflicts: Dict[int, List[str]] = {}

    for vid, items in by_vendor.items():
        fmts = sorted({f for (f, _e, _u) in items})
        if len(fmts) > 1:
            conflicts[vid] = fmts

        # sort by updated_at desc (None last), then email asc
        def _key(t):
            fmt, email, updated_at = t
            # None should be last => use empty string and a flag
            return (updated_at is not None, updated_at or "", -len(email), email)

        # We want newest updated_at: sort by (has_updated_at, updated_at) descending.
        # We'll do it manually:
        items_sorted = sorted(
            items,
            key=lambda t: (
                0 if t[2] is None else 1,
                t[2] or "",
                t[1],
            ),
            reverse=True,
        )
        best_fmt, best_email, _best_u = items_sorted[0]
        chosen[vid] = (best_fmt, best_email)

    return chosen, conflicts


def fetch_existing_vendor_formats(cur) -> Dict[int, Optional[str]]:
    cur.execute(
        """
        SELECT id,
               CASE
                 WHEN date_format IS NULL THEN NULL
                 ELSE NULLIF(BTRIM(date_format), '')
               END AS date_format
        FROM vendors
        """
    )
    out: Dict[int, Optional[str]] = {}
    for vid, fmt in cur.fetchall():
        out[int(vid)] = (str(fmt).strip() if fmt is not None else None) or None
    return out


def apply_updates(cur, updates: List[Tuple[int, str]]) -> int:
    """
    updates: list of (vendor_id, date_format)
    """
    if not updates:
        return 0
    sql = """
        UPDATE vendors AS v
        SET date_format = x.date_format,
            updated_at = NOW()
        FROM (VALUES %s) AS x(vendor_id, date_format)
        WHERE v.id = x.vendor_id
    """
    execute_values(cur, sql, updates, template="(%s, %s)")
    return int(cur.rowcount or 0)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="Print what would change; do not write.")
    ap.add_argument("--apply", action="store_true", help="Write updates to DB.")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite vendors.date_format even if already set.")
    args = ap.parse_args()

    if args.dry_run and args.apply:
        print("Choose only one: --dry-run OR --apply")
        return 2
    if not args.dry_run and not args.apply:
        # default to dry-run (safe)
        args.dry_run = True

    with get_conn() as conn, conn.cursor() as cur:
        # Preconditions
        for t in ("authorized_senders", "vendors", "vendor_contacts"):
            if not _has_table(cur, t):
                raise SystemExit(f"Missing table: public.{t}")
        if not _has_column(cur, "authorized_senders", "date_format"):
            raise SystemExit("Missing column: public.authorized_senders.date_format")
        if not _has_column(cur, "vendors", "date_format"):
            raise SystemExit("Missing column: public.vendors.date_format")

        rows = fetch_candidate_rows(cur)
        if not rows:
            print("No vendor_contacts emails matched authorized_senders with a non-empty date_format. Nothing to migrate.")
            return 0

        chosen, conflicts = choose_one_format_per_vendor(rows)
        existing = fetch_existing_vendor_formats(cur)

        # Build updates list
        updates: List[Tuple[int, str]] = []
        for vid, (fmt, _email) in chosen.items():
            cur_fmt = existing.get(vid)
            if args.overwrite or _is_missing_or_auto(cur_fmt):
                updates.append((vid, fmt))

        print(f"Matched contacts with date_format: {len(rows)} row(s)")
        print(f"Vendors with at least one mapped date_format: {len(chosen)}")
        print(f"Will update vendors: {len(updates)} (overwrite={args.overwrite})")

        if conflicts:
            print(f"[WARN] {len(conflicts)} vendor(s) have multiple different date_formats among contacts (showing up to 10):")
            # show up to 10 vendors
            shown = 0
            # map vendor_id -> company_name
            vid_to_company: Dict[int, str] = {}
            for vid, company, _email, _fmt, _u in rows:
                vid_to_company.setdefault(vid, company)
            for vid, fmts in list(conflicts.items())[:10]:
                print(f"  - vendor_id={vid} company={vid_to_company.get(vid)!r} formats={fmts}")
                shown += 1

        if args.dry_run:
            # show sample planned updates
            # map vendor_id -> company_name
            vid_to_company: Dict[int, str] = {}
            for vid, company, _email, _fmt, _u in rows:
                vid_to_company.setdefault(vid, company)
            for vid, fmt in updates[:50]:
                print(f"  - {vid_to_company.get(vid)!r} (vendor_id={vid}) -> {fmt}")
            if len(updates) > 50:
                print(f"  ... and {len(updates) - 50} more")
            return 0

        affected = apply_updates(cur, updates)
        conn.commit()
        print(f"Done. DB reported affected rows: {affected}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())


