# ───────────────────────── Push comparison results to JeraSoft ─────────────────────────
from typing import Optional, Dict, List, Tuple
import re
import pandas as pd
import os
import json
import requests

# Reuse your existing env vars
J_API_URL = os.getenv("JERASOFT_API_URL", "http://billing.voipsystem.org:3080")
J_API_KEY = os.getenv("JERA_SOFT_API_KEY")

# Shared session with retries (optional but recommended)
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
_session_push = requests.Session()
_session_push.mount("http://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=0.4, status_forcelist=[502,503,504])))
_session_push.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=0.4, status_forcelist=[502,503,504])))

def _rpc_call(method: str, params: Dict, api_url: Optional[str] = None) -> Dict:
    """Low-level JSON-RPC helper."""
    api_url = api_url or J_API_URL
    payload = {"jsonrpc": "2.0", "id": 1, "method": method, "params": params}
    r = _session_push.post(api_url, headers={"Content-Type":"application/json","Accept":"application/json"}, json=payload, timeout=300)
    r.raise_for_status()
    data = r.json()
    if "error" in data and data["error"]:
        # JeraSoft error comes here
        raise RuntimeError(f"JeraSoft RPC error in {method}: {json.dumps(data['error'])}")
    return data.get("result")

_bi_re = re.compile(r"^\s*(\d+)\s*/\s*(\d+)\s*$")
def _parse_billing_increment(bi: Optional[str]) -> Tuple[Optional[int], Optional[int]]:
    """'60/60' -> (60,60); returns (None,None) if blank/invalid."""
    if not isinstance(bi, str):
        return None, None
    m = _bi_re.match(bi.strip())
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))

def _fmt_date(d: object) -> Optional[str]:
    """Accepts datetime/str; returns 'YYYY-MM-DD' or None."""
    if d is None or (isinstance(d, float) and pd.isna(d)):
        return None
    if isinstance(d, str):
        d = d.strip()
        if not d:
            return None
        try:
            return str(pd.to_datetime(d, errors="coerce").date())
        except Exception:
            return None
    try:
        return str(pd.to_datetime(d, errors="coerce").date())
    except Exception:
        return None

def _search_rate_id(table_id: int, code: str, effective_from: str, api_url: Optional[str]=None) -> Optional[int]:
    """
    Find an existing rate row id by table + code + effective_from.
    Returns None if not found.
    """
    params = {
        "AUTH": J_API_KEY,
        "rate_tables_id": table_id,
        "code": code,
        "effective_from": effective_from,
        "state": "all",
        "status": "all",
        "limit": 1,
        "offset": 0,
    }
    res = _rpc_call("rates.search", params, api_url=api_url)
    if isinstance(res, list) and res:
        rid = res[0].get("id")
        return int(rid) if rid is not None else None
    return None

def _create_rate(table_id: int, code: str, value: float, eff: str,
                 min_vol: Optional[int], pay_int: Optional[int],
                 api_url: Optional[str]=None) -> int:
    """Create a rate row and return its id."""
    params = {
        "AUTH": J_API_KEY,
        "rate_tables_id": table_id,
        "code": str(code),
        "effective_from": eff,
        "value": float(value),
        "status": "active",
    }
    if min_vol is not None: params["min_volume"] = int(min_vol)
    if pay_int is not None: params["pay_interval"] = int(pay_int)
    res = _rpc_call("rates.create", params, api_url=api_url)
    rid = res.get("id") if isinstance(res, dict) else None
    if rid is None:
        raise RuntimeError(f"rates.create returned no id: {res}")
    return int(rid)

def _update_rate(rate_id: int, value: Optional[float]=None,
                 min_vol: Optional[int]=None, pay_int: Optional[int]=None,
                 status: Optional[str]=None, api_url: Optional[str]=None) -> None:
    """Update a rate row by id (only provided fields are changed)."""
    params = {"AUTH": J_API_KEY, "id": int(rate_id)}
    if value is not None:    params["value"] = float(value)
    if min_vol is not None:  params["min_volume"] = int(min_vol)
    if pay_int is not None:  params["pay_interval"] = int(pay_int)
    if status is not None:   params["status"] = status
    _ = _rpc_call("rates.update", params, api_url=api_url)

def push_comparison_to_jerasoft(df: pd.DataFrame,
                                table_id: int,
                                *,
                                api_url: Optional[str]=None,
                                api_key: Optional[str]=None,
                                accepted_statuses: Tuple[str, ...]=("Accepted",),
                                create_if_missing: bool=True,
                                update_if_exists: bool=True,
                                dry_run: bool=False) -> Dict[str,int]:
    """
    Push comparison results into a JeraSoft rate table (upsert).
    Expects columns:
      'Code', 'New Rate', 'Effective Date', 'New Billing Increment' (optional), 'Status' (optional)

    Only rows with Status ∈ accepted_statuses are uploaded.

    Returns a summary dict.
    """
    if api_key:
        # allow overriding process-wide key
        global J_API_KEY
        J_API_KEY = api_key

    if not J_API_KEY:
        raise ValueError("Missing JERA_SOFT_API_KEY (env) or pass api_key=)")

    req_cols = ["Code","New Rate","Effective Date"]
    for c in req_cols:
        if c not in df.columns:
            raise ValueError(f"Missing required column in DataFrame: {c}")

    summary = {"processed": 0, "created": 0, "updated": 0, "skipped_status": 0, "skipped_invalid": 0, "errors": 0}

    for idx, row in df.iterrows():
        summary["processed"] += 1

        # Filter by status (optional)
        if "Status" in df.columns and accepted_statuses:
            st = str(row.get("Status") or "").strip()
            if st not in accepted_statuses:
                summary["skipped_status"] += 1
                continue

        code = str(row.get("Code") or "").strip()
        rate = row.get("New Rate")
        eff  = _fmt_date(row.get("Effective Date"))
        bi   = row.get("New Billing Increment")
        mv, pi = _parse_billing_increment(bi)

        if not code or eff is None:
            summary["skipped_invalid"] += 1
            continue
        try:
            rate = float(rate)
        except Exception:
            summary["skipped_invalid"] += 1
            continue

        if dry_run:
            print(f"[DRY-RUN] upsert code={code} eff={eff} rate={rate} min_vol={mv} pay_int={pi}")
            continue

        try:
            rid = _search_rate_id(table_id, code, eff, api_url=api_url)

            if rid is None:
                if not create_if_missing:
                    summary["skipped_invalid"] += 1
                    continue
                _ = _create_rate(table_id, code, rate, eff, mv, pi, api_url=api_url)
                summary["created"] += 1
            else:
                if update_if_exists:
                    _update_rate(rid, value=rate, min_vol=mv, pay_int=pi, api_url=api_url)
                    summary["updated"] += 1
                else:
                    # exists but we chose not to update
                    pass

        except Exception as e:
            summary["errors"] += 1
            print(f"[push-error] row #{idx} code={code} eff={eff}: {e}")

    return summary




table_id = 1234  # your target table
summary = push_comparison_to_jerasoft(
    cmp_df,
    table_id=table_id,
    accepted_statuses=("Accepted",),  # only push Accepted rows
    create_if_missing=True,           # create missing rates
    update_if_exists=True,            # update if already present
    dry_run=False                     # set True to preview without writing
)
print("Upload summary:", summary)