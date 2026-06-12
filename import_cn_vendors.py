"""
Import CN vendors and vendor contacts from CSV file into the database.
CSV columns: Name, AM contacts, Call networks AM, NOC, Rates, INVOICES
Only the 'Rates' column emails are imported as vendor_contacts (these are the rate sheet senders).
"""

import csv
import os
import re
import psycopg2
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(dotenv_path=Path(__file__).with_name(".env"), override=True)

CSV_FILE = Path(__file__).parent / "CN Client Info_2026 - For Review(CN Client detail) (1).csv"

def get_conn():
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_DATABASE"),
        user=os.getenv("DB_USERNAME"),
        password=os.getenv("DB_PASSWORD"),
    )

def extract_emails(raw: str) -> list:
    """Extract all valid emails from a raw string."""
    if not raw or not raw.strip():
        return []
    return [e.strip().lower() for e in re.findall(r'[\w\.\+\-]+@[\w\.\-]+\.\w+', raw)]

def import_vendors():
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now()

    vendors_added = 0
    contacts_added = 0
    skipped = 0

    with open(CSV_FILE, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            company_name = (row.get("Name") or "").strip()
            company_name = company_name.encode("ascii", "ignore").decode("ascii").strip()
            rates_raw    = row.get("Rates") or ""
            am_raw       = row.get("AM contacts") or ""
            noc_raw      = row.get("NOC") or ""

            if not company_name:
                skipped += 1
                continue

            # Insert vendor
            cur.execute("""
                INSERT INTO vendors (company_name, status, auto_reject_email_enabled, created_at, updated_at)
                VALUES (%s, TRUE, FALSE, %s, %s)
                ON CONFLICT DO NOTHING
                RETURNING id
            """, (company_name, now, now))

            row_result = cur.fetchone()
            if row_result:
                vendor_id = row_result[0]
                vendors_added += 1
            else:
                # Already exists — fetch its id
                cur.execute("SELECT id FROM vendors WHERE company_name = %s", (company_name,))
                result = cur.fetchone()
                if not result:
                    print(f"  [WARN] Could not find or insert vendor: {company_name}")
                    skipped += 1
                    continue
                vendor_id = result[0]

            # Insert rate contacts (Rates column) — these are the verified senders
            rate_emails = extract_emails(rates_raw)
            for email in rate_emails:
                cur.execute("""
                    INSERT INTO vendor_contacts (vendor_id, email, role, status, is_added_to_cc, created_at, updated_at)
                    VALUES (%s, %s, 'Rates', TRUE, FALSE, %s, %s)
                    ON CONFLICT DO NOTHING
                """, (vendor_id, email, now, now))
                if cur.rowcount:
                    contacts_added += 1

            # Insert AM contacts
            for email in extract_emails(am_raw):
                cur.execute("""
                    INSERT INTO vendor_contacts (vendor_id, email, role, status, is_added_to_cc, created_at, updated_at)
                    VALUES (%s, %s, 'AM', TRUE, FALSE, %s, %s)
                    ON CONFLICT DO NOTHING
                """, (vendor_id, email, now, now))
                if cur.rowcount:
                    contacts_added += 1

            # Insert NOC contacts
            for email in extract_emails(noc_raw):
                cur.execute("""
                    INSERT INTO vendor_contacts (vendor_id, email, role, status, is_added_to_cc, created_at, updated_at)
                    VALUES (%s, %s, 'NOC', TRUE, FALSE, %s, %s)
                    ON CONFLICT DO NOTHING
                """, (vendor_id, email, now, now))
                if cur.rowcount:
                    contacts_added += 1

            print(f"  [OK] {company_name} — rates: {rate_emails}")

    conn.commit()
    cur.close()
    conn.close()

    print(f"\n=== IMPORT COMPLETE ===")
    print(f"Vendors added  : {vendors_added}")
    print(f"Contacts added : {contacts_added}")
    print(f"Skipped        : {skipped}")

if __name__ == "__main__":
    import_vendors()
