"""
Set vendor_contacts.is_added_to_cc = TRUE for selected roles.

Use-case:
  Enable CC by default for roles:
    - "hayo am"
    - "am contacts"

Usage:
  Dry-run (no DB writes):
    python tools/set_vendor_contacts_cc_true_for_roles.py --dry-run

  Apply (writes to DB):
    python tools/set_vendor_contacts_cc_true_for_roles.py --apply

Optional:
  Limit to rows where is_added_to_cc is currently FALSE (default True):
    --only-false
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List

# Ensure project root is on sys.path so "import database" works when running from tools/
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import get_conn  # noqa: E402


TARGET_ROLES: List[str] = ["hayo am", "am contacts"]


def _count_matches(*, only_false: bool) -> int:
    sql = """
        SELECT COUNT(*)
        FROM vendor_contacts
        WHERE lower(role) = ANY(%s)
    """
    params = (TARGET_ROLES,)
    if only_false:
        sql += " AND is_added_to_cc IS NOT TRUE"
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, params)
        return int(cur.fetchone()[0] or 0)


def _apply_update(*, only_false: bool) -> int:
    sql = """
        UPDATE vendor_contacts
        SET is_added_to_cc = TRUE,
            updated_at = NOW()
        WHERE lower(role) = ANY(%s)
    """
    params = (TARGET_ROLES,)
    if only_false:
        sql += " AND is_added_to_cc IS NOT TRUE"
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, params)
        affected = int(cur.rowcount or 0)
        conn.commit()
        return affected


def main() -> int:
    ap = argparse.ArgumentParser(description="Set vendor_contacts.is_added_to_cc=TRUE for selected roles.")
    ap.add_argument("--dry-run", action="store_true", help="Show how many rows would be updated (no writes).")
    ap.add_argument("--apply", action="store_true", help="Apply the update in DB.")
    ap.add_argument(
        "--only-false",
        action="store_true",
        help="Update only rows where is_added_to_cc is not already TRUE.",
    )
    args = ap.parse_args()

    if args.apply and args.dry_run:
        print("Choose only one: --dry-run OR --apply")
        return 2

    mode = "apply" if args.apply else "dry-run"
    only_false = bool(args.only_false)

    n = _count_matches(only_false=only_false)
    print(f"[OK] mode={mode} roles={TARGET_ROLES} only_false={only_false}")
    print(f"[OK] matching rows: {n}")

    if not args.apply:
        return 0

    affected = _apply_update(only_false=only_false)
    print(f"[DONE] updated rows: {affected}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


