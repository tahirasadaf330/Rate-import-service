"""
Fetch distinct sender emails from the `rate_uploads` table.

Usage examples:
  python fetch_unique_rate_upload_emails.py
  python fetch_unique_rate_upload_emails.py --since 2025-12-01
  python fetch_unique_rate_upload_emails.py --since 2025-12-01 --until 2025-12-18 --format csv --out emails.csv
  python fetch_unique_rate_upload_emails.py --format xlsx --out unique_rate_upload_emails.xlsx
  python fetch_unique_rate_upload_emails.py --include-empty --include-null

DB connection:
  Uses the same environment variables as `database.get_conn()`:
    DB_HOST, DB_PORT, DB_DATABASE, DB_USERNAME, DB_PASSWORD
  Loads `.env` via python-dotenv (same as the rest of the repo).
"""

import argparse
import csv
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from dotenv import load_dotenv

# Ensure project root is importable when running as a script (Windows-friendly).
# This lets `python fetch_unique_rate_upload_emails.py` import `database.py`.
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from database import get_conn  # noqa: E402


def _parse_yyyy_mm_dd(s: str) -> datetime:
    # Interpret as local date boundary; DB compares against timestamps.
    # We keep it simple: pass as YYYY-MM-DD string to Postgres and let it cast.
    # This function validates format only.
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except ValueError as e:
        raise argparse.ArgumentTypeError(f"Invalid date {s!r}. Expected YYYY-MM-DD.") from e


def fetch_distinct_sender_emails(
    *,
    since: Optional[str],
    until: Optional[str],
    normalize: bool,
    include_null: bool,
    include_empty: bool,
    limit: Optional[int],
) -> List[str]:
    """
    Returns a list of distinct sender emails from rate_uploads.
    """
    # Base select (normalized or raw)
    if normalize:
        sel = "LOWER(BTRIM(sender_email)) AS sender_email"
        where_nonempty = "LOWER(BTRIM(sender_email)) <> ''"
    else:
        sel = "sender_email"
        where_nonempty = "BTRIM(sender_email) <> ''"

    where_parts = []
    params = []

    if not include_null:
        where_parts.append("sender_email IS NOT NULL")
    if not include_empty:
        where_parts.append(where_nonempty)

    # Prefer received_at if present; fallback to created_at in case older rows are missing received_at.
    # If your schema guarantees received_at, this still works fine.
    if since:
        where_parts.append("COALESCE(received_at, created_at) >= %s::date")
        params.append(since)
    if until:
        where_parts.append("COALESCE(received_at, created_at) <= %s::date")
        params.append(until)

    where_sql = ""
    if where_parts:
        where_sql = "WHERE " + " AND ".join(where_parts)

    limit_sql = ""
    if limit is not None:
        limit_sql = "LIMIT %s"
        params.append(int(limit))

    sql = f"""
        SELECT DISTINCT {sel}
        FROM rate_uploads
        {where_sql}
        ORDER BY 1
        {limit_sql}
    """

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, tuple(params))
        rows = cur.fetchall()

    return [r[0] for r in rows]


def main(argv: Optional[List[str]] = None) -> int:
    load_dotenv()

    p = argparse.ArgumentParser(description="Fetch distinct sender emails from rate_uploads.")
    p.add_argument("--since", type=_parse_yyyy_mm_dd, help="Only include uploads on/after this date (YYYY-MM-DD).")
    p.add_argument("--until", type=_parse_yyyy_mm_dd, help="Only include uploads on/before this date (YYYY-MM-DD).")
    p.add_argument("--raw", action="store_true", help="Do not normalize; return raw sender_email values.")
    p.add_argument("--include-null", action="store_true", help="Include NULL sender_email values.")
    p.add_argument("--include-empty", action="store_true", help="Include empty/whitespace sender_email values.")
    p.add_argument("--limit", type=int, default=None, help="Limit number of results.")
    p.add_argument("--format", choices=["text", "csv", "json", "xlsx"], default="text", help="Output format.")
    p.add_argument("--out", default=None, help="Output file path. Default: stdout (or default .xlsx name for xlsx).")
    p.add_argument("--count", action="store_true", help="Print total unique count to stderr.")

    args = p.parse_args(argv)

    since_s = args.since.strftime("%Y-%m-%d") if args.since else None
    until_s = args.until.strftime("%Y-%m-%d") if args.until else None

    emails = fetch_distinct_sender_emails(
        since=since_s,
        until=until_s,
        normalize=(not args.raw),
        include_null=args.include_null,
        include_empty=args.include_empty,
        limit=args.limit,
    )

    if args.count:
        print(f"unique_emails={len(emails)}", file=sys.stderr)

    if args.format == "xlsx":
        # Default output name if not provided
        out_path = args.out or "unique_rate_upload_emails.xlsx"
        if not out_path.lower().endswith(".xlsx"):
            out_path = out_path + ".xlsx"

        # Import lazily to keep CLI fast if user doesn't need Excel
        import pandas as pd  # type: ignore

        df = pd.DataFrame({"sender_email": emails})
        df.to_excel(out_path, index=False)
        print(f"✅ Wrote Excel file: {out_path}")
        return 0

    out_fh = open(args.out, "w", newline="", encoding="utf-8") if args.out else sys.stdout
    try:
        if args.format == "text":
            for e in emails:
                out_fh.write(f"{e}\n")
        elif args.format == "csv":
            w = csv.writer(out_fh)
            w.writerow(["sender_email"])
            for e in emails:
                w.writerow([e])
        else:  # json
            json.dump({"sender_emails": emails}, out_fh, ensure_ascii=False, indent=2)
            out_fh.write("\n")
    finally:
        if args.out:
            out_fh.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


