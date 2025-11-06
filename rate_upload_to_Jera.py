#!/usr/bin/env python3
"""
Non-destructive connectivity test to JeraSoft:
- REST: PATCH /rates/tables/{id} to update only the 'tag' field
- (Optional) JSON-RPC fallback: rates.tables.update

Usage:
  # PowerShell: set your API key for this session
  #   $env:JERA_SOFT_API_KEY = "your-key-here"
  # Run (table 4330, default host)
  #   python test_push_jerasoft.py
  #
  # Custom tag:
  #   python test_push_jerasoft.py --table-id 4330 --tag "TEST_FROM_SCRIPT"
  #
  # Dry run (shows the request only):
  #   python test_push_jerasoft.py --dry-run
  #
  # If your Windows trust store blocks TLS, try:
  #   python test_push_jerasoft.py --insecure
"""

import os
import json
import argparse
import http.client
import ssl
from urllib.parse import urlparse
from datetime import datetime, timezone
from datetime import datetime

tag_value = f"TEST_{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}"
payload = {"tag": tag_value}
print("Generated tag:", tag_value)


try:
    import certifi  # for a reliable CA bundle on Windows
    CERT_PATH = certifi.where()
except Exception:
    CERT_PATH = None

DEFAULT_REST_BASE = "https://billing.voipsystem.org"
DEFAULT_RPC_URL   = "http://billing.voipsystem.org:3080"
DEFAULT_TABLE_ID  = 4330


def _utc_tag(prefix: str = "TEST") -> str:
    return f"{prefix}_{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}"


def rest_patch_table(base_url: str,
                     api_key: str,
                     table_id: int,
                     payload: dict,
                     dry_run: bool = False,
                     insecure: bool = False):
    """
    PATCH /rates/tables/{id} with http.client, providing proper TLS context.
    """
    parsed = urlparse(base_url)
    if parsed.scheme not in ("https", "http"):
        raise ValueError(f"Unsupported scheme in base URL: {base_url}")

    host = parsed.netloc or parsed.path
    path = f"/rates/tables/{table_id}"

    body = json.dumps(payload)
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Api-Key": api_key,
    }

    print("\n[REST] PATCH request")
    print("  Host:", host)
    print("  Path:", path)
    print("  Headers:", {k: ("<hidden>" if k.lower().startswith("x-api-key") else v) for k, v in headers.items()})
    print("  Body:", body)

    if dry_run:
        print("  (dry-run) Skipping network call.")
        return None, None, None

    if parsed.scheme == "https":
        if insecure:
            ctx = ssl._create_unverified_context()
        else:
            ctx = ssl.create_default_context(cafile=CERT_PATH) if CERT_PATH else ssl.create_default_context()
        conn = http.client.HTTPSConnection(host, timeout=60, context=ctx)
    else:
        conn = http.client.HTTPConnection(host, timeout=60)

    try:
        conn.request("PATCH", path, body=body, headers=headers)
        res = conn.getresponse()
        data = res.read()
        print("\n[REST] Response")
        print("  Status:", res.status, res.reason)
        try:
            decoded = data.decode("utf-8", errors="replace")
            print("  Body:", json.dumps(json.loads(decoded), indent=2))
        except Exception:
            print("  Body (raw):", data[:1000])
        return res.status, res.reason, data
    finally:
        conn.close()


def rpc_update_table(rpc_url: str,
                     api_key: str,
                     table_id: int,
                     tag_value: str,
                     dry_run: bool = False,
                     insecure: bool = False):
    """
    JSON-RPC fallback to rates.tables.update (matches your jerasoft.py style).
    """
    payload = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "rates.tables.update",
        "params": {
            "AUTH": api_key,
            "id": int(table_id),
            "tag": tag_value,
        },
    }

    parsed = urlparse(rpc_url)
    if parsed.scheme not in ("https", "http"):
        raise ValueError(f"Unsupported scheme in rpc URL: {rpc_url}")

    host = parsed.netloc or parsed.path
    path = parsed.path or "/"

    body = json.dumps(payload)
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    print("\n[JSON-RPC] POST request")
    print("  Host:", host)
    print("  Path:", path)
    print("  Body:", body)

    if dry_run:
        print("  (dry-run) Skipping network call.")
        return None, None, None

    if parsed.scheme == "https":
        if insecure:
            ctx = ssl._create_unverified_context()
        else:
            ctx = ssl.create_default_context(cafile=CERT_PATH) if CERT_PATH else ssl.create_default_context()
        conn = http.client.HTTPSConnection(host, timeout=60, context=ctx)
    else:
        conn = http.client.HTTPConnection(host, timeout=60)

    try:
        conn.request("POST", path, body=body, headers=headers)
        res = conn.getresponse()
        data = res.read()
        print("\n[JSON-RPC] Response")
        print("  Status:", res.status, res.reason)
        try:
            decoded = data.decode("utf-8", errors="replace")
            print("  Body:", json.dumps(json.loads(decoded), indent=2))
        except Exception:
            print("  Body (raw):", data[:1000])
        return res.status, res.reason, data
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser(description="Test pushing to JeraSoft (safe tag update).")
    ap.add_argument("--table-id", type=int, default=DEFAULT_TABLE_ID, help="Rate table ID to patch (default: 4330)")
    ap.add_argument("--rest-base", default=DEFAULT_REST_BASE, help="REST base URL (default: https://billing.voipsystem.org)")
    ap.add_argument("--rpc-url", default=os.getenv("JERASOFT_RPC_URL", DEFAULT_RPC_URL), help="JSON-RPC URL (optional fallback)")
    ap.add_argument("--tag", default=None, help="Tag to set; default = TEST_<UTC timestamp>")
    ap.add_argument("--dry-run", action="store_true", help="Print the request but do not send")
    ap.add_argument("--no-rpc", action="store_true", help="Skip JSON-RPC fallback")
    ap.add_argument("--insecure", action="store_true", help="Bypass TLS verification (NOT for production)")
    args = ap.parse_args()

    api_key = os.getenv("JERA_SOFT_API_KEY")
    if not api_key:
        raise SystemExit("Missing env var JERA_SOFT_API_KEY")

    tag_value = args.tag or _utc_tag("TEST")

    # REST (doc-style)
    try:
        rest_status, _, _ = rest_patch_table(
            base_url=args.rest_base,
            api_key=api_key,
            table_id=args.table_id,
            payload={"tag": tag_value},
            dry_run=args.dry_run,
            insecure=args.insecure,
        )
    except Exception as e:
        print(f"\n[REST] Error: {e}")
        rest_status = None

    # JSON-RPC fallback
    rpc_status = None
    if not args.no_rpc:
        try:
            rpc_status, _, _ = rpc_update_table(
                rpc_url=args.rpc_url,
                api_key=api_key,
                table_id=args.table_id,
                tag_value=tag_value,
                dry_run=args.dry_run,
                insecure=args.insecure,
            )
        except Exception as e:
            print(f"\n[JSON-RPC] Error: {e}")
            rpc_status = None

    print("\n=== SUMMARY ===")
    print("REST status:", rest_status)
    print("RPC status :", rpc_status)
    print("Tag tested :", tag_value)
    if rest_status == 200 and rpc_status == 200:
        print("Note: If you see a JSON-RPC error like 'tag-already_assigned', the table is assigned and JeraSoft blocks tag changes. Try another field or an unassigned table for this test.")


if __name__ == "__main__":
    main()
