#!/usr/bin/env python3
"""
One-time migration:
  1) set vendors.auto_reject_email_enabled = TRUE
  2) set vendors.date_format based on authorized_senders.date_format matched via vendor_contacts.email

Safety:
  - Default mode is --dry-run (no DB writes)
  - --apply performs the UPDATE
  - Idempotent: only updates rows that would actually change

Examples:
  python tools/enable_vendor_auto_reject_email_enable.py --dry-run
  python tools/enable_vendor_auto_reject_email_enable.py --apply

Optional:
  python tools/enable_vendor_auto_reject_email_enable.py --apply --only-null
  python tools/enable_vendor_auto_reject_email_enable.py --apply --overwrite-date-format
  python tools/enable_vendor_auto_reject_email_enable.py --apply --all-contacts
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from psycopg2.extras import execute_values

# Ensure project root is importable when running from tools/
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import get_conn  # noqa: E402


_IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _validate_ident(name: str, kind: str) -> str:
    name = str(name or "").strip()
    if not name or not _IDENT_RE.match(name):
        raise SystemExit(f"Invalid {kind}: {name!r}. Use only letters/numbers/underscore.")
    return name


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


def _count(cur, schema: str, table: str, where_sql: Optional[str] = None) -> int:
    sql = f'SELECT COUNT(*) FROM "{schema}"."{table}"'
    if where_sql:
        sql += f" WHERE {where_sql}"
    cur.execute(sql)
    return int(cur.fetchone()[0])


def _sample_rows(cur, schema: str, table: str, column: str, where_sql: str) -> list[Tuple[int, str, Optional[bool]]]:
    # Best-effort: show id + company_name if present; else show id only.
    has_company = _has_column(cur, table, "company_name", schema=schema)
    if has_company:
        cur.execute(
            f"""
            SELECT id, company_name, {column}
            FROM "{schema}"."{table}"
            WHERE {where_sql}
            ORDER BY id ASC
            LIMIT 50
            """
        )
    else:
        cur.execute(
            f"""
            SELECT id, ''::text AS company_name, {column}
            FROM "{schema}"."{table}"
            WHERE {where_sql}
            ORDER BY id ASC
            LIMIT 50
            """
        )
    out = []
    for vid, company, val in cur.fetchall():
        out.append((int(vid), str(company or ""), (bool(val) if val is not None else None)))
    return out


def _is_missing_or_auto(fmt: Optional[str]) -> bool:
    if fmt is None:
        return True
    s = str(fmt).strip()
    return (not s) or (s.lower() == "auto")


def _fetch_sender_date_format_rows(
    cur,
    schema: str,
    vendors_table: str,
    *,
    contact_role: Optional[str],
) -> List[Tuple[int, str, str, str, Optional[str]]]:
    """
    Returns rows:
      (vendor_id, company_name, contact_email, date_format, sender_updated_at_iso)
    Source of truth is authorized_senders:
      authorized_senders.email -> vendor_contacts.email -> vendor_contacts.vendor_id -> vendors.id
    Filtering:
      - Only authorized_senders where date_format is non-empty and not 'AUTO'
      - If contact_role is provided: only vendor_contacts rows where role == contact_role (case-insensitive, trimmed)
    """
    # NOTE: We assume authorized_senders has updated_at; if not, we still work (NULLs last).
    where_role = ""
    params: list = []
    if contact_role:
        where_role = ' AND LOWER(BTRIM(vc.role)) = LOWER(BTRIM(%s))'
        params.append(contact_role)

    cur.execute(
        f"""
        SELECT
            v.id AS vendor_id,
            v.company_name,
            LOWER(BTRIM(vc.email)) AS contact_email,
            BTRIM(a.date_format) AS date_format,
            COALESCE(a.updated_at::text, NULL) AS sender_updated_at
        FROM "{schema}"."authorized_senders" a
        JOIN "{schema}"."vendor_contacts" vc
          ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        JOIN "{schema}"."{vendors_table}" v
          ON v.id = vc.vendor_id
        WHERE a.date_format IS NOT NULL
          AND BTRIM(a.date_format) <> ''
          AND LOWER(BTRIM(a.date_format)) <> 'auto'
          {where_role}
        """,
        tuple(params),
    )
    return [
        (int(r[0]), str(r[1]), str(r[2]), str(r[3]), (str(r[4]) if r[4] is not None else None))
        for r in cur.fetchall()
    ]


def _choose_one_format_per_vendor(
    rows: List[Tuple[int, str, str, str, Optional[str]]],
) -> Tuple[Dict[int, Tuple[str, str]], Dict[int, List[str]]]:
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


def _fetch_existing_vendor_date_formats(cur, schema: str, vendors_table: str) -> Dict[int, Optional[str]]:
    cur.execute(
        f"""
        SELECT id,
               CASE
                 WHEN date_format IS NULL THEN NULL
                 ELSE NULLIF(BTRIM(date_format), '')
               END AS date_format
        FROM "{schema}"."{vendors_table}"
        """
    )
    out: Dict[int, Optional[str]] = {}
    for vid, fmt in cur.fetchall():
        out[int(vid)] = (str(fmt).strip() if fmt is not None else None) or None
    return out


def _apply_vendor_date_format_updates(cur, schema: str, vendors_table: str, updates: List[Tuple[int, str]], *, has_updated_at: bool) -> int:
    """
    updates: list of (vendor_id, date_format)
    """
    if not updates:
        return 0
    if has_updated_at:
        sql = f"""
            UPDATE "{schema}"."{vendors_table}" AS v
            SET date_format = x.date_format,
                updated_at = NOW()
            FROM (VALUES %s) AS x(vendor_id, date_format)
            WHERE v.id = x.vendor_id
        """
    else:
        sql = f"""
            UPDATE "{schema}"."{vendors_table}" AS v
            SET date_format = x.date_format
            FROM (VALUES %s) AS x(vendor_id, date_format)
            WHERE v.id = x.vendor_id
        """
    execute_values(cur, sql, updates, template="(%s, %s)")
    return int(cur.rowcount or 0)


def _debug_date_format_migration(cur, schema: str, vendors_table: str, contact_role: Optional[str]) -> None:
    """
    Print diagnostics for why the date_format migration may produce zero rows.
    """
    print("[DATE-FORMAT][DEBUG] diagnostics:")
    if contact_role:
        print(f"[DATE-FORMAT][DEBUG] role filter enabled: role={contact_role!r}")
    else:
        print("[DATE-FORMAT][DEBUG] role filter disabled: matching ANY vendor_contacts role")

    # 1) What roles exist?
    cur.execute(
        f"""
        SELECT LOWER(BTRIM(role)) AS role_norm, COUNT(*) AS n
        FROM "{schema}"."vendor_contacts"
        GROUP BY LOWER(BTRIM(role))
        ORDER BY n DESC
        LIMIT 20
        """
    )
    roles = cur.fetchall()
    if roles:
        print("[DATE-FORMAT][DEBUG] top vendor_contacts.role values (normalized):")
        for r, n in roles:
            print(f"  - {r!r}: {int(n)}")
    else:
        print("[DATE-FORMAT][DEBUG] vendor_contacts has no rows.")

    # 2) How many contacts match the role (or total contacts if role filter disabled)?
    if contact_role:
        cur.execute(
            f"""
            SELECT COUNT(*)
            FROM "{schema}"."vendor_contacts"
            WHERE LOWER(BTRIM(role)) = LOWER(BTRIM(%s))
            """,
            (contact_role,),
        )
        n_role = int(cur.fetchone()[0])
        print(f"[DATE-FORMAT][DEBUG] contacts with role={contact_role!r}: {n_role}")
        role_where = "WHERE LOWER(BTRIM(vc.role)) = LOWER(BTRIM(%s))"
        role_params = (contact_role,)
    else:
        cur.execute(f'SELECT COUNT(*) FROM "{schema}"."vendor_contacts"')
        n_all = int(cur.fetchone()[0])
        print(f"[DATE-FORMAT][DEBUG] total vendor_contacts: {n_all}")
        role_where = ""
        role_params = tuple()

    # 3) How many vendor contact emails exist in authorized_senders?
    cur.execute(
        f"""
        SELECT COUNT(*)
        FROM "{schema}"."vendor_contacts" vc
        JOIN "{schema}"."authorized_senders" a
          ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        {role_where}
        """,
        role_params,
    )
    n_match_sender = int(cur.fetchone()[0])
    label = "role contacts" if contact_role else "contacts"
    print(f"[DATE-FORMAT][DEBUG] {label} whose email exists in authorized_senders: {n_match_sender}")

    # 4) How many of those have a usable date_format?
    role_and = f" AND {role_where.replace('WHERE', '', 1)}" if role_where else ""
    cur.execute(
        f"""
        SELECT COUNT(*)
        FROM "{schema}"."vendor_contacts" vc
        JOIN "{schema}"."authorized_senders" a
          ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        WHERE a.date_format IS NOT NULL
          AND BTRIM(a.date_format) <> ''
          AND LOWER(BTRIM(a.date_format)) <> 'auto'
          {role_and}
        """,
        role_params,
    )
    n_usable_fmt = int(cur.fetchone()[0])
    print(f"[DATE-FORMAT][DEBUG] matched senders with non-empty/non-AUTO date_format: {n_usable_fmt}")

    # 5) Sample a few vendor contact emails missing in authorized_senders
    cur.execute(
        f"""
        SELECT vc.vendor_id,
               COALESCE(v.company_name, '') AS company_name,
               LOWER(BTRIM(vc.email)) AS email_norm
        FROM "{schema}"."vendor_contacts" vc
        LEFT JOIN "{schema}"."authorized_senders" a
          ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        LEFT JOIN "{schema}"."{vendors_table}" v
          ON v.id = vc.vendor_id
        WHERE a.email IS NULL
          {role_and}
        ORDER BY vc.vendor_id ASC
        LIMIT 20
        """,
        role_params,
    )
    missing = cur.fetchall()
    if missing:
        print("[DATE-FORMAT][DEBUG] sample vendor_contacts emails NOT found in authorized_senders (up to 20):")
        for vid, company, email in missing:
            suffix = f" company={str(company)!r}" if company else ""
            print(f"  - vendor_id={int(vid)}{suffix} email={str(email)!r}")

    # 6) Sample a few matched but unusable date formats (NULL/blank/AUTO)
    cur.execute(
        f"""
        SELECT vc.vendor_id,
               COALESCE(v.company_name, '') AS company_name,
               LOWER(BTRIM(vc.email)) AS email_norm,
               a.date_format
        FROM "{schema}"."vendor_contacts" vc
        JOIN "{schema}"."authorized_senders" a
          ON LOWER(BTRIM(a.email)) = LOWER(BTRIM(vc.email))
        LEFT JOIN "{schema}"."{vendors_table}" v
          ON v.id = vc.vendor_id
        WHERE (a.date_format IS NULL OR BTRIM(a.date_format) = '' OR LOWER(BTRIM(a.date_format)) = 'auto')
          {role_and}
        ORDER BY vc.vendor_id ASC
        LIMIT 20
        """,
        role_params,
    )
    unusable = cur.fetchall()
    if unusable:
        print("[DATE-FORMAT][DEBUG] sample matched senders with unusable date_format (up to 20):")
        for vid, company, email, df in unusable:
            suffix = f" company={str(company)!r}" if company else ""
            print(f"  - vendor_id={int(vid)}{suffix} email={str(email)!r} date_format={df!r}")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="One-time migration: enable vendors.auto_reject_email_enabled and migrate vendors.date_format from authorized_senders for AM contacts."
    )
    ap.add_argument("--dry-run", action="store_true", help="Print what would change; do not write.")
    ap.add_argument("--apply", action="store_true", help="Write updates to DB.")
    ap.add_argument("--schema", default="public", help="DB schema (default: public)")
    ap.add_argument("--table", default="vendors", help="Table name (default: vendors)")
    ap.add_argument("--column", default="auto_reject_email_enabled", help="Column name (default: auto_reject_email_enabled)")
    ap.add_argument("--only-null", action="store_true", help="Only set TRUE where column is NULL (skip FALSE).")
    ap.add_argument("--skip-auto-reject", action="store_true", help="Skip enabling auto_reject_email_enabled.")
    ap.add_argument("--skip-date-format", action="store_true", help="Skip migrating vendors.date_format.")
    ap.add_argument(
        "--am-role",
        default="am contacts",
        help="vendor_contacts.role value to treat as AM contacts when NOT using --all-contacts (default: 'am contacts')",
    )
    ap.add_argument(
        "--all-contacts",
        action="store_true",
        help="Match authorized_senders emails against vendor_contacts for ANY role (ignores --am-role).",
    )
    ap.add_argument(
        "--overwrite-date-format",
        action="store_true",
        help="Overwrite vendors.date_format even if already set (default: only fills missing/auto).",
    )
    ap.add_argument(
        "--debug-date-format",
        action="store_true",
        help="Print diagnostics for date_format migration (useful when it reports 'No AM contacts matched...').",
    )
    args = ap.parse_args()

    if args.dry_run and args.apply:
        print("Choose only one: --dry-run OR --apply")
        return 2
    if not args.dry_run and not args.apply:
        # default to dry-run (safe)
        args.dry_run = True

    schema = _validate_ident(args.schema, "schema")
    table = _validate_ident(args.table, "table")
    column = _validate_ident(args.column, "column")

    # Where clause for rows that would change
    where_to_update = f'"{column}" IS NULL' if args.only_null else f'"{column}" IS DISTINCT FROM TRUE'

    with get_conn() as conn, conn.cursor() as cur:
        # Preconditions
        if not _has_table(cur, table, schema=schema):
            raise SystemExit(f"Missing table: {schema}.{table}")
        if not args.skip_auto_reject and not _has_column(cur, table, column, schema=schema):
            raise SystemExit(f"Missing column: {schema}.{table}.{column}")

        # Date format migration prerequisites (only if enabled)
        if not args.skip_date_format:
            for t in ("vendor_contacts", "authorized_senders"):
                if not _has_table(cur, t, schema=schema):
                    raise SystemExit(f"Missing table: {schema}.{t}")
            if not _has_column(cur, "authorized_senders", "date_format", schema=schema):
                raise SystemExit(f"Missing column: {schema}.authorized_senders.date_format")
            if not _has_column(cur, table, "date_format", schema=schema):
                raise SystemExit(f"Missing column: {schema}.{table}.date_format")
            if not _has_column(cur, "vendor_contacts", "role", schema=schema):
                raise SystemExit(f"Missing column: {schema}.vendor_contacts.role")
            if not _has_column(cur, "vendor_contacts", "email", schema=schema):
                raise SystemExit(f"Missing column: {schema}.vendor_contacts.email")

        has_updated_at = _has_column(cur, table, "updated_at", schema=schema)

        # --- Part 1: auto_reject_email_enabled ---
        total_vendors = _count(cur, schema, table)
        affected_auto_reject = 0
        will_update_auto_reject = 0
        if not args.skip_auto_reject:
            already_true = _count(cur, schema, table, where_sql=f'"{column}" IS TRUE')
            will_update_auto_reject = _count(cur, schema, table, where_sql=where_to_update)
            print(f"[AUTO-REJECT] schema={schema} table={table} column={column}")
            print(f"[AUTO-REJECT] total vendors: {total_vendors}")
            print(f"[AUTO-REJECT] already TRUE: {already_true}")
            print(f"[AUTO-REJECT] will update to TRUE: {will_update_auto_reject} (only_null={bool(args.only_null)})")

            if args.dry_run and will_update_auto_reject:
                sample = _sample_rows(cur, schema, table, f'"{column}"', where_to_update)
                print("[AUTO-REJECT][DRY-RUN] sample rows that would change (up to 50):")
                for vid, company, val in sample:
                    suffix = f" company={company!r}" if company else ""
                    print(f"  - vendor_id={vid}{suffix} current={val}")
            elif args.apply and will_update_auto_reject:
                if has_updated_at:
                    sql = f"""
                        UPDATE "{schema}"."{table}"
                        SET "{column}" = TRUE,
                            updated_at = NOW()
                        WHERE {where_to_update}
                    """
                else:
                    sql = f"""
                        UPDATE "{schema}"."{table}"
                        SET "{column}" = TRUE
                        WHERE {where_to_update}
                    """
                cur.execute(sql)
                affected_auto_reject = int(cur.rowcount or 0)
        else:
            print("[AUTO-REJECT] skipped.")

        # --- Part 2: vendors.date_format from authorized_senders for AM contacts ---
        affected_date_format = 0
        planned_date_updates: List[Tuple[int, str]] = []
        if not args.skip_date_format:
            contact_role = None if args.all_contacts else args.am_role
            rows = _fetch_sender_date_format_rows(cur, schema, table, contact_role=contact_role)
            if not rows:
                role_msg = "ANY role" if contact_role is None else f"role={contact_role!r}"
                print(f"[DATE-FORMAT] No vendor_contacts matched authorized_senders with non-empty date_format ({role_msg}).")
                if args.debug_date_format:
                    _debug_date_format_migration(cur, schema, table, contact_role)
            else:
                chosen, conflicts = _choose_one_format_per_vendor(rows)
                existing = _fetch_existing_vendor_date_formats(cur, schema, table)

                for vid, (fmt, _email) in chosen.items():
                    cur_fmt = existing.get(vid)
                    if args.overwrite_date_format or _is_missing_or_auto(cur_fmt):
                        planned_date_updates.append((vid, fmt))

                role_label = "ANY" if contact_role is None else repr(contact_role)
                print(f"[DATE-FORMAT] role={role_label}")
                print(f"[DATE-FORMAT] matched contact rows with date_format: {len(rows)}")
                print(f"[DATE-FORMAT] vendors with at least one mapped date_format: {len(chosen)}")
                print(f"[DATE-FORMAT] will update vendors.date_format: {len(planned_date_updates)} (overwrite={bool(args.overwrite_date_format)})")

                if conflicts:
                    print(f"[DATE-FORMAT][WARN] {len(conflicts)} vendor(s) have multiple different date_formats among AM contacts (showing up to 10):")
                    vid_to_company: Dict[int, str] = {}
                    for vid, company, _email, _fmt, _u in rows:
                        vid_to_company.setdefault(vid, company)
                    for vid, fmts in list(conflicts.items())[:10]:
                        print(f"  - vendor_id={vid} company={vid_to_company.get(vid)!r} formats={fmts}")

                if args.dry_run and planned_date_updates:
                    vid_to_company: Dict[int, str] = {}
                    for vid, company, _email, _fmt, _u in rows:
                        vid_to_company.setdefault(vid, company)
                    print("[DATE-FORMAT][DRY-RUN] sample planned vendor updates (up to 50):")
                    for vid, fmt in planned_date_updates[:50]:
                        print(f"  - {vid_to_company.get(vid)!r} (vendor_id={vid}) -> {fmt}")
                    if len(planned_date_updates) > 50:
                        print(f"  ... and {len(planned_date_updates) - 50} more")
                elif args.apply and planned_date_updates:
                    affected_date_format = _apply_vendor_date_format_updates(
                        cur,
                        schema,
                        table,
                        planned_date_updates,
                        has_updated_at=has_updated_at,
                    )
        else:
            print("[DATE-FORMAT] skipped.")

        if args.dry_run:
            print("[DRY-RUN] done (no DB writes).")
            return 0

        conn.commit()
        print(
            "Done."
            f" auto_reject_updated={affected_auto_reject}"
            f" date_format_updated={affected_date_format}"
            f" (planned_auto_reject={will_update_auto_reject} planned_date_format={len(planned_date_updates)})"
        )
        return 0


if __name__ == "__main__":
    raise SystemExit(main())


