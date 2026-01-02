"""
Import vendors + vendor_contacts from an Excel workbook with multiple sheets.

Expected columns (case-insensitive, trimmed):
  - Company Name
  - AM contacts
  - Rates

Behavior:
  - Reads ALL sheets by default.
  - Upserts vendors by company_name into `vendors`.
  - Upserts vendor_contacts by email into `vendor_contacts` and assigns role:
      - "am"   for AM contacts column
      - "rate" for Rates column
  - Prefixes/extra text in the email cells are ignored; we extract valid emails via regex.

Requirements:
  - DB env vars set (.env): DB_HOST, DB_PORT, DB_DATABASE, DB_USERNAME, DB_PASSWORD
  - Tables exist:
      vendors(company_name UNIQUE, ...)
      vendor_contacts(email UNIQUE, vendor_id FK, contact_name NOT NULL, role, ...)

Usage:
  Dry-run (no DB writes):
    python tools/import_vendors_and_contacts_from_excel.py --file path/to/vendors.xlsx --dry-run

  Apply (writes to DB):
    python tools/import_vendors_and_contacts_from_excel.py --file path/to/vendors.xlsx --apply
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

# Ensure project root is on sys.path so "import database" works when running from tools/
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import get_conn  # noqa: E402


EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def _norm_col(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def extract_emails(cell) -> List[str]:
    """Extract unique emails from a cell (string/NaN), lowercased."""
    if cell is None:
        return []
    try:
        if pd.isna(cell):
            return []
    except Exception:
        pass
    text = str(cell)
    found = EMAIL_RE.findall(text)
    # de-dupe while preserving order
    out: List[str] = []
    seen = set()
    for e in found:
        el = e.strip().lower()
        if el and el not in seen:
            seen.add(el)
            out.append(el)
    return out


@dataclass(frozen=True)
class ContactRow:
    company_name: str
    email: str
    role: str  # "am" | "rate"


def _detect_columns(df: pd.DataFrame) -> Tuple[str, str, str]:
    """
    Return (company_col, am_col, rate_col).
    Raises ValueError if required columns aren't found.
    """
    cols = list(df.columns)
    norm2real: Dict[str, str] = {_norm_col(c): c for c in cols}

    company = None
    for k in ("company name", "company", "vendor", "vendor name"):
        if k in norm2real:
            company = norm2real[k]
            break
    if not company:
        raise ValueError("Missing required column: Company Name")

    am = None
    for k in ("am contacts", "am contact", "am", "account manager", "account manager contacts"):
        if k in norm2real:
            am = norm2real[k]
            break
    if not am:
        raise ValueError("Missing required column: AM contacts")

    rate = None
    for k in ("rates", "rate", "rates email", "rate email", "rates contacts"):
        if k in norm2real:
            rate = norm2real[k]
            break
    if not rate:
        raise ValueError("Missing required column: Rates")

    return company, am, rate


def load_rows_from_workbook(xlsx_path: Path, sheet_names: Optional[Sequence[str]] = None) -> Tuple[List[str], List[ContactRow]]:
    """
    Read all sheets (or subset) and return:
      - vendor_names: unique company_name list
      - contacts: list of (company_name, email, role)
    """
    xlsx = pd.ExcelFile(xlsx_path)
    sheets = list(sheet_names) if sheet_names else list(xlsx.sheet_names)

    vendor_names: List[str] = []
    vendor_seen = set()
    contacts: List[ContactRow] = []

    for sh in sheets:
        df = pd.read_excel(xlsx, sheet_name=sh)
        if df is None or df.empty:
            continue

        try:
            company_col, am_col, rate_col = _detect_columns(df)
        except ValueError as e:
            # Skip sheets that don't have the expected format (but keep going)
            print(f"[SKIP] sheet={sh!r}: {e}")
            continue

        for _, row in df.iterrows():
            company = str(row.get(company_col, "")).strip()
            if not company or company.lower() in {"company name", "company"}:
                continue

            if company not in vendor_seen:
                vendor_seen.add(company)
                vendor_names.append(company)

            for email in extract_emails(row.get(am_col)):
                contacts.append(ContactRow(company_name=company, email=email, role="am"))
            for email in extract_emails(row.get(rate_col)):
                contacts.append(ContactRow(company_name=company, email=email, role="rate"))

    return vendor_names, contacts


def _ensure_tables_exist() -> None:
    sql = """
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = 'public'
          AND table_name IN ('vendors', 'vendor_contacts')
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql)
        found = {r[0] for r in cur.fetchall()}
    missing = [t for t in ("vendors", "vendor_contacts") if t not in found]
    if missing:
        raise RuntimeError(f"Missing required tables in DB: {missing}")


def upsert_vendor(company_name: str) -> int:
    sql = """
        INSERT INTO vendors (company_name, created_at, updated_at)
        VALUES (%s, NOW(), NOW())
        ON CONFLICT (company_name) DO UPDATE SET
          updated_at = NOW()
        RETURNING id;
    """
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (company_name,))
        vid = cur.fetchone()[0]
        conn.commit()
        return int(vid)


def upsert_vendor_contact(vendor_id: int, email: str, role: str) -> None:
    sql = """
        INSERT INTO vendor_contacts
          (vendor_id, contact_name, email, role, is_added_to_cc, status, created_at, updated_at)
        VALUES
          (%s, %s, %s, %s, FALSE, TRUE, NOW(), NOW())
        ON CONFLICT (email) DO UPDATE SET
          vendor_id = EXCLUDED.vendor_id,
          role = EXCLUDED.role,
          contact_name = EXCLUDED.contact_name,
          updated_at = NOW();
    """

    def _contact_name_from_email(addr: str) -> str:
        """
        Derive a simple contact_name from an email address:
          - take the local-part (before @)
          - split on ., _, -, and whitespace
          - take the first token
        Fallback to full email if anything goes wrong.
        """
        try:
            local = (addr or "").split("@", 1)[0].strip()
            if not local:
                return addr
            token = re.split(r"[.\s_\-]+", local)[0].strip()
            return token or addr
        except Exception:
            return addr

    contact_name = _contact_name_from_email(email)
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(sql, (vendor_id, contact_name, email, role))
        conn.commit()


def main() -> int:
    ap = argparse.ArgumentParser(description="Import vendors + vendor_contacts from a multi-sheet Excel workbook.")
    ap.add_argument("--file", required=True, help="Path to Excel workbook (.xlsx).")
    ap.add_argument("--sheet", action="append", default=None, help="Optional: import only this sheet (repeatable).")
    ap.add_argument("--dry-run", action="store_true", help="Parse and print summary only (no DB writes).")
    ap.add_argument("--apply", action="store_true", help="Write to DB (upserts).")
    args = ap.parse_args()

    xlsx_path = Path(args.file).expanduser().resolve()
    if not xlsx_path.exists():
        print(f"File not found: {xlsx_path}")
        return 2
    if xlsx_path.suffix.lower() not in {".xlsx", ".xls"}:
        print("Expected an Excel file (.xlsx/.xls).")
        return 2

    if args.apply and args.dry_run:
        print("Choose only one: --dry-run OR --apply")
        return 2
    mode = "apply" if args.apply else "dry-run"

    vendor_names, contacts = load_rows_from_workbook(xlsx_path, sheet_names=args.sheet)

    # Basic parsing summary
    unique_contacts = {(c.company_name, c.email, c.role) for c in contacts}
    unique_emails = {c.email for c in contacts}
    print(f"[OK] mode={mode} file={xlsx_path.name!r}")
    print(f"[OK] vendors found: {len(vendor_names)}")
    print(f"[OK] contacts found: {len(contacts)} (unique triplets={len(unique_contacts)} unique_emails={len(unique_emails)})")

    # Detect email reused across companies/roles
    email_to_companies: Dict[str, set] = {}
    for c in contacts:
        email_to_companies.setdefault(c.email, set()).add(c.company_name)
    multi_company = {e: comps for e, comps in email_to_companies.items() if len(comps) > 1}
    if multi_company:
        sample = list(multi_company.items())[:10]
        print(f"[WARN] {len(multi_company)} email(s) appear under multiple companies (showing up to 10):")
        for e, comps in sample:
            print(f"  - {e}: {sorted(comps)}")

    if not args.apply:
        return 0

    # DB writes
    _ensure_tables_exist()

    company_to_vid: Dict[str, int] = {}
    for company in vendor_names:
        vid = upsert_vendor(company)
        company_to_vid[company] = vid

    # Upsert contacts (dedupe triplets to avoid redundant DB work)
    for (company, email, role) in sorted(unique_contacts):
        vid = company_to_vid.get(company)
        if not vid:
            # should not happen, but be safe
            vid = upsert_vendor(company)
            company_to_vid[company] = vid
        upsert_vendor_contact(vendor_id=int(vid), email=email, role=role)

    print(f"[DONE] upserted vendors={len(company_to_vid)} contacts_unique_triplets={len(unique_contacts)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


