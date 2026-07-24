"""
Jerasoft rates exporter (functions-only)

Usage (example):
    export_rates_by_query(
        target_query="Quickcom tel PRM trunk Prefix:1001 USD",
        output_path="quickcom_rates.xlsx",
    )

The entrypoint `export_rates_by_query()`:
- Finds the best-matching TERM* rate table for `target_query`.
- Fetches active current & future rates from that table.
- Saves a tidy Excel file at `output_path`.
- Returns a small result dict with table_id, score, and counts.

Configuration:
- API URL & KEY are read from env by default: JERA_SOFT_API_KEY, JERASOFT_API_URL
- You can also pass `api_url` / `api_key` explicitly to any function.
"""

from __future__ import annotations
import os
import re
import json
from pathlib import Path
from datetime import datetime
from difflib import SequenceMatcher
from typing import Dict, List, Tuple, Optional
import pandas as pd
import requests
from dotenv import load_dotenv
from json import JSONDecodeError

load_dotenv()

# -------------------------------------------------------------------
# Defaults / Env
# -------------------------------------------------------------------
DEFAULT_API_URL = os.getenv("JERASOFT_API_URL", "http://billing.voipsystem.org:3080")
DEFAULT_API_KEY = os.getenv("JERA_SOFT_API_KEY")
DEFAULT_HEADERS = {"Content-Type": "application/json", "Accept": "application/json"}

# _session = requests.Session()

# --- robust session with retries for transient upstream issues ---
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_session = requests.Session()
_retry = Retry(
    total=4,
    backoff_factor=0.4,
    status_forcelist=[502, 503, 504],
    allowed_methods=["POST"],
)
_session.mount("http://", HTTPAdapter(max_retries=_retry))
_session.mount("https://", HTTPAdapter(max_retries=_retry))

#__________________________Enforce prefix_____________________________
# --- prefix utilities (add near your string utils) ---
# Allow table names like '... PREFIX:#2223' by ignoring the '#', and also handle textual NONE
_prefix_in_name_pat = re.compile(r'(?:prefix|prfx)[:\s\-#]*(\d+|none)\b', re.I)

def normalize_prefix(prefix) -> Optional[str]:
    if prefix is None:
        return None
    s = str(prefix).strip()
    if not s:
        return None
    # Preserve leading zeros: return first digit run as-is, or map textual NONE.
    if "none" in s.lower():
        return "NONE"
    m = re.search(r"\d+", s)
    return m.group(0) if m else None

def table_prefix_from_name(name: str) -> Optional[str]:
    """Extract numeric prefix from a table name like '... PREFIX:33' or 'PRFX-33'."""
    m = _prefix_in_name_pat.search(name or "")
    if not m:
        return None
    val = m.group(1)
    return "NONE" if val.lower() == "none" else val

def table_has_prefix(name: str, prefix_code: str) -> bool:
    """True if table name contains the exact prefix number."""
    tp = table_prefix_from_name(name)
    return tp == normalize_prefix(prefix_code)

# ----- Trunk matching (aliases + whole-word match in table name) -----
TRUNK_ALIASES: Dict[str, List[str]] = {
    "STD":     ["std", "standard", "gold", "wholesale"],
    "PRM":     ["prm", "premium", "silver", "prs"],
    "CC":      ["cc", "call center"],
    "ORTP":    ["ortp"],
    "TDM":     ["tdm"],
    "DID":     ["did"],
    "SPECIAL": ["special"],
    "ATX":     ["atx"],
}

def canonicalize_trunk(trunk: Optional[str]) -> Optional[str]:
    if not trunk:
        return None
    key = str(trunk).strip().lower()
    if not key:
        return None
    for canonical, aliases in TRUNK_ALIASES.items():
        if key in (a.lower() for a in aliases):
            return canonical
    return str(trunk).strip().upper()

def table_matches_trunk(name: str, trunk: str) -> bool:
    """True if table name contains the trunk (or any alias) as a whole word."""
    canonical = canonicalize_trunk(trunk)
    if not canonical:
        return False
    aliases = TRUNK_ALIASES.get(canonical, [canonical.lower()])
    hay = (name or "").lower()
    for alias in aliases:
        pat = r"(?<![A-Za-z0-9])" + re.escape(alias.lower()) + r"(?![A-Za-z0-9])"
        if re.search(pat, hay):
            return True
    return False

#_____________________________________________________________________________________

# -------------------------------------------------------------------
# String utils
# -------------------------------------------------------------------
_norm_space = re.compile(r"\s+")
_norm_hashes = re.compile(r"[#]+")

def normalize(s: str) -> str:
    s = s or ""
    s = s.lower().strip()
    s = _norm_hashes.sub("", s)
    s = s.replace(".", " ")
    s = _norm_space.sub(" ", s)
    return s.strip()

def fuzzy_score(a: str, b: str) -> float:
    """Blend of SequenceMatcher and Jaccard, in [0,1]."""
    a_n, b_n = normalize(a), normalize(b)
    base = SequenceMatcher(None, a_n, b_n).ratio()
    a_tokens, b_tokens = set(a_n.split()), set(b_n.split())
    jacc = (len(a_tokens & b_tokens) / len(a_tokens | b_tokens)) if (a_tokens and b_tokens) else 0.0
    return 0.85 * base + 0.15 * jacc

def name_starts_with_term(name: str) -> bool:
    prefix = os.getenv("JERA_TABLE_PREFIX", "TERM").upper()
    return str(name).lstrip().upper().startswith(prefix)

def name_contains_company(name: str, company_kw: str) -> bool:
    return company_kw in normalize(name)

def extract_company_keyword(query: str) -> str:
    """Derive company keyword from the first token (alnum, ., _, -)."""
    m = re.search(r"[A-Za-z0-9._-]+", query or "")
    return (m.group(0).lower() if m else "").strip()

def _post_json(api_url: str, payload: Dict, timeout: int = 500) -> Dict:
    resp = _session.post(api_url, headers=DEFAULT_HEADERS, json=payload, timeout=timeout)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "")
    try:
        return resp.json()
    except JSONDecodeError as e:
        body = resp.text or ""
        clen = resp.headers.get("Content-Length")
        raise RuntimeError(
            f"JSON decode failed at pos {e.pos}: {e.msg}. "
            f"status={resp.status_code} content_type={ctype} content_length={clen} "
            f"body_len={len(body)} body_start={body[:1000]!r} body_end={body[-500:]!r}"
        )

def fetch_all_tables(api_url: Optional[str] = None, api_key: Optional[str] = None, page_size: int = 500) -> List[Dict]:
    """Fetch all rate tables via pagination."""
    api_url = api_url or DEFAULT_API_URL
    api_key = api_key or DEFAULT_API_KEY
    if not api_key:
        raise ValueError("Missing API key (env JERA_SOFT_API_KEY or pass api_key).")

    all_tables: List[Dict] = []
    offset = 0
    while True:
        payload = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "rates.tables.search",
            "params": {"AUTH": api_key, "limit": page_size, "offset": offset},
        }
        data = _post_json(api_url, payload)
        if "result" not in data:
            raise RuntimeError(f"API error: {data.get('error')}")
        page = data["result"]
        if not page:
            break
        all_tables.extend(page)
        offset += page_size
    return all_tables

def find_best_term_table(
    target_query: str,
    subject: str,
    api_url: Optional[str] = None,
    api_key: Optional[str] = None,
    top_k: int = 5,
    prefix_code: Optional[str] = None,
    trunk_code: Optional[str] = None,
) -> Tuple[int, Dict, List[Tuple[float, Dict]]]:

    if not target_query:
        raise ValueError("target_query must be non-empty")

    company_kw = extract_company_keyword(target_query)
    if not company_kw:
        raise ValueError("Could not extract a company keyword from target_query.")

    tables = fetch_all_tables(api_url=api_url, api_key=api_key)
    candidates = [
        t for t in tables
        if name_starts_with_term(t.get("name", "")) and name_contains_company(t.get("name", ""), company_kw)
    ]
    if not candidates:
        return f"No TERM* tables found containing company '{company_kw}'.", "", ""

    # Enforce explicit prefix if provided
    norm_pref = normalize_prefix(prefix_code)
    if norm_pref:
        exact_prefix = [t for t in candidates if table_has_prefix(t.get("name", ""), norm_pref)]
        if not exact_prefix:
            return (f"No TERM* tables found for company '{company_kw}' with PREFIX:{norm_pref}.",
                    "", "")
        candidates = exact_prefix

    # If still ambiguous after company+prefix, narrow down using the trunk filter
    if len(candidates) >= 2 and trunk_code:
        trunk_filtered = [t for t in candidates if table_matches_trunk(t.get("name", ""), trunk_code)]
        if len(trunk_filtered) == 1:
            candidates = trunk_filtered
        elif len(trunk_filtered) == 0:
            return (
                f"No TERM* tables found for company '{company_kw}' "
                f"with PREFIX:{norm_pref or 'ANY'} and trunk '{trunk_code}'.",
                "",
                "",
            )
        else:
            table_names = [t.get("name", "") for t in trunk_filtered]
            return (
                f"Multiple tables found after company+prefix+trunk filtering — "
                f"please assign a table manually. Tables: {table_names}",
                "",
                "",
            )

    # Safety net: still >1 (e.g. trunk not supplied, or trunk alias matches many)
    if len(candidates) > 1:
        table_names = [t.get("name", "") for t in candidates]
        return (
            f"Multiple tables found after filtering — please refine your input. "
            f"Tables: {table_names}",
            "",
            "",
        )

    best_table = candidates[0]
    best_id = best_table.get("id")
    if best_id is None:
        raise KeyError("Best table did not include an 'id' field")

    # Return with score 1.0 since it's the only match
    return int(best_id), best_table, [(1.0, best_table)]
def get_table_id_by_name(
    table_name: str,
    api_url: Optional[str] = None,
    api_key: Optional[str] = None,
) -> Optional[int]:
    """
    Look up a Jerasoft rate table ID by its exact name (case-insensitive).
    Returns the table ID if found, otherwise None.

    Example:
        tid = get_table_id_by_name("TERM Quickcom tel PRM trunk Prefix:1001 USD")
    """
    api_url = api_url or DEFAULT_API_URL
    api_key = api_key or DEFAULT_API_KEY
    if not api_key:
        raise ValueError("Missing API key (env JERA_SOFT_API_KEY or pass api_key).")

    # We'll reuse your fetch_all_tables for reliability
    tables = fetch_all_tables(api_url=api_url, api_key=api_key)
    normalized_target = normalize(table_name)

    for t in tables:
        if normalize(t.get("name", "")) == normalized_target:
            return t.get("id")

    # If no exact match, optionally try fuzzy match
    best_match = max(
        ((fuzzy_score(table_name, t.get("name", "")), t) for t in tables),
        key=lambda x: x[0],
        default=(0, None),
    )

    if best_match[0] > 0.9:  # strong fuzzy match threshold
        print(f"⚠️ No exact match, using fuzzy match with score={best_match[0]:.3f}")
        return best_match[1].get("id")

    print(f"❌ Table '{table_name}' not found.")
    return None


def get_table_code_deck_id(
    table_id: int,
    api_url: Optional[str] = None,
    api_key: Optional[str] = None,
) -> Optional[int]:
    """
    Return the code_decks_id a rate table is linked to, or None if not found.
    Used so imports send the table's OWN code deck instead of a hardcoded value
    (JeraSoft rejects a code deck the table isn't on: "Code Deck is not allowed
    for this import" — e.g. CN tables are on deck 66, Hayo tables on deck 19).
    """
    for t in fetch_all_tables(api_url=api_url, api_key=api_key):
        if t.get("id") == table_id:
            return t.get("code_decks_id")
    return None


def fetch_active_current_future_rates(
    table_id: int,
    api_url: Optional[str] = None,
    api_key: Optional[str] = None,
    when_utc: Optional[datetime] = None,
    page_limit: int = 500,  # start conservative; adaptive logic will adjust
) -> pd.DataFrame:
    """
    Fetch active current & future rates with adaptive pagination.
    If the server returns truncated JSON with HTTP 200, we retry the same page,
    and if needed, halve the page size until it parses.
    """
    api_url = api_url or DEFAULT_API_URL
    api_key = api_key or DEFAULT_API_KEY
    if not api_key:
        raise ValueError("Missing API key (env JERA_SOFT_API_KEY or pass api_key).")

    when_utc = when_utc or datetime.utcnow()

    offset = 0
    all_records: List[Dict] = []

    # never go below this; below 100 is just suffering
    MIN_LIMIT = 100
    current_limit = max(MIN_LIMIT, int(page_limit))

    base_params = {
        "AUTH": api_key,
        "rate_tables_id": table_id,
        "state": "current_future",
        "status": "active",
        "__tz": "UTC",
        "dt": when_utc.strftime("%Y-%m-%d %H:%M:%S"),
        "order": ["+code", "-effective_from"],
    }

    while True:
        params = dict(base_params, limit=current_limit, offset=offset)
        payload = {"jsonrpc": "2.0", "id": 1, "method": "rates.search", "params": params}

        # retry decode at this offset/limit before shrinking
        attempts = 0
        while True:
            attempts += 1
            try:
                data = _post_json(api_url, payload, timeout=500)
                break  # parsed ok
            except RuntimeError as e:
                # Our _post_json wraps JSONDecodeError in RuntimeError with a clear message
                if "JSON decode failed" in str(e):
                    if attempts < 2:
                        # try once more at the same limit, could be a hiccup
                        continue
                    if current_limit > MIN_LIMIT:
                        # shrink page and retry same offset
                        current_limit = max(MIN_LIMIT, current_limit // 2)
                        attempts = 0
                        continue
                # not a truncation case or can't shrink further — bubble up
                raise

        result = data.get("result")
        if not isinstance(result, list):
            raise RuntimeError(f"Unexpected response or error: {str(data)[:400]}")

        if not result:
            break  # all done

        all_records.extend(result)
        offset += len(result)
        # keep current_limit as adapted; if the server handled it, we keep using it

    if not all_records:
        return pd.DataFrame(columns=[
            "Dst Code", "Dst Code Name", "Rate", "Effective Date", "Billing Increment"
        ])

    df = pd.DataFrame(all_records)

    min_vol = df.get("min_volume")
    pay_int = df.get("pay_interval")
    min_vol_str = min_vol.where(min_vol.notna(), "").astype(str) if min_vol is not None else ""
    pay_int_str = pay_int.where(pay_int.notna(), "").astype(str) if pay_int is not None else ""
    df["Billing Increment"] = (
        (min_vol_str if isinstance(min_vol_str, pd.Series) else "") + "/" +
        (pay_int_str if isinstance(pay_int_str, pd.Series) else "")
    ).str.strip("/")

    keep = ["code", "code_name", "value", "effective_from", "Billing Increment"]
    keep_existing = [c for c in keep if c in df.columns]
    df_selected = df[keep_existing].rename(columns={
        "code": "Dst Code",
        "code_name": "Dst Code Name",
        "value": "Rate",
        "effective_from": "Effective Date",
    })

    return df_selected.reset_index(drop=True)


def save_rates_to_excel(df: pd.DataFrame, output_path: str) -> str:
    """Save DataFrame to Excel, ensuring parent folder exists. Returns absolute path."""
    if not output_path:
        raise ValueError("output_path must be provided")

    out_path = Path(output_path).expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Let pandas pick a writer engine that's available (xlsxwriter/openpyxl)
    df.to_excel(out_path, index=False)
    return str(out_path)

def export_rates_by_query(
    target_query: str,
    output_path: str,
    subject: str,
    prefix_code: Optional[str] = None,
    trunk_code: Optional[str] = None,
    *,
    api_url: Optional[str] = None,
    api_key: Optional[str] = None,
    return_debug: bool = True,
    force_table_name: Optional[str] = None,   # <-- NEW
) -> Dict:
    # If a specific table name is approved/forced, use it directly.
    if force_table_name:
        tid = get_table_id_by_name(force_table_name, api_url=api_url, api_key=api_key)
        if not tid:
            return f"Forced table '{force_table_name}' not found."
        best_table = {"id": tid, "name": force_table_name}
        top_scored = [(1.0, best_table)]
        table_id = tid
    else:
        table_id, best_table, top_scored = find_best_term_table(
            target_query=target_query, api_url=api_url, api_key=api_key,
            subject=subject, prefix_code=prefix_code, trunk_code=trunk_code,
        )
        if best_table == "" and top_scored == "":
            return table_id  # string error from find_best_term_table

    print(f"DEBUG: Best table: ID={table_id} NAME='{best_table.get('name')}' SCORE={top_scored[0][0]:.3f}")

    df = fetch_active_current_future_rates(table_id=table_id, api_url=api_url, api_key=api_key)
    print(f"DEBIG: Fetched {df.shape[0]} active current & future rates.")

    saved_to = save_rates_to_excel(df, output_path)

    result = {
        "table_id": table_id,
        "rows": int(df.shape[0]),
        "saved_to": saved_to,
    }

    if return_debug:
        result.update({
            "best_table_name": best_table.get("name"),
            "top_candidates": [
                {"score": round(score, 3), "id": t.get("id"), "name": t.get("name")}
                for score, t in top_scored
            ],
        })

    return result

if __name__ == "__main__":
    # Example quick-start (reads API key from env):
    info = export_rates_by_query(
        subject="[Sipstatus Global LTD] [Retail] [62750] [USD]",
        target_query="[Sipstatus Global LTD] [Retail] [62750] [USD]",
        output_path="quickcom_rates.xlsx",
    )
    print(json.dumps(info, indent=2))
