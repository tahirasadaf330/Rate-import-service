"""
Process Outlook Inbox via Microsoft Graph using app permissions:
- Verify sender against VERIFIED_SENDERS
- Validate Subject via SUBJECT_PATTERN
- Save only allowed attachment types (.csv, .xlsx, etc.) with safe filenames
- Make per-message folders named <sender>_<UTC timestamp>
- Log short body text (HTML → text if needed)
- Write per-message metadata.json alongside saved attachments

Env (.env next to script):
  TENANT_ID=...
  CLIENT_ID=...
  CLIENT_SECRET=...
  USER_EMAIL=target_mailbox@domain.com
Optional:
  VERBOSE=1  # to print decoded token roles
"""

import os, sys, re, json, base64, time, unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Set, Tuple, Optional, List, Dict

# from valid_emails import VERIFIED_SENDERS  as verified_senders

import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
from urllib.parse import quote
# from open_ai import validate_subject_openai

from valid_emails import get_verified_senders  # list of allowed sender emails
from msal import ConfidentialClientApplication
import logging
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
# ========================= Debug Toggles =========================
DEBUG = False         # ultra-verbose prints for every step
DRY_RUN = False      # if True, do not write files; only log decisions
# ================================================================



# ==========================================================
# Code for loging failed emails

FAILED_EMAILS_PATH = Path(__file__).with_name("failed_emails.json")

_FAILED_SCHEMA = {
    "version": 1,
    "last_updated_utc": None,
    "totals": {
        "unverified_sender": 0,
        "subject_invalid": 0,
        "no_attachments": 0,
        "no_valid_attachments": 0,
        "all": 0
    },
    "buckets": {
        "unverified_sender": [],
        "subject_invalid": [],
        "no_attachments": [],
        "no_valid_attachments": []
    }
}

def _load_failed_log() -> dict:
    try:
        if FAILED_EMAILS_PATH.exists():
            with open(FAILED_EMAILS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    # light schema heal
                    for k in _FAILED_SCHEMA:
                        data.setdefault(k, _FAILED_SCHEMA[k])
                    for b in _FAILED_SCHEMA["buckets"]:
                        data["buckets"].setdefault(b, [])
                    for t in _FAILED_SCHEMA["totals"]:
                        data["totals"].setdefault(t, 0)
                    return data
    except Exception as e:
        print(f"(warn) could not read {FAILED_EMAILS_PATH}: {e}", file=sys.stderr)
    return json.loads(json.dumps(_FAILED_SCHEMA))

def _recalc_totals_inplace(data: dict) -> None:
    buckets = data.get("buckets", {})
    totals  = data.get("totals", {})
    keys = ("unverified_sender", "subject_invalid", "no_attachments", "no_valid_attachments")
    for k in keys:
        totals[k] = len(buckets.get(k, []))
    totals["all"] = sum(totals[k] for k in keys)

def _atomic_write_failed_log(data: dict) -> None:
    _recalc_totals_inplace(data)  # <-- add this line
    tmp = FAILED_EMAILS_PATH.with_suffix(FAILED_EMAILS_PATH.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, FAILED_EMAILS_PATH)

def _entry_exists(bucket: list, key: Optional[str]) -> bool:
    if not key:
        return False
    for item in bucket:
        if item.get("id") == key or item.get("internetMessageId") == key:
            return True
    return False

def log_failed_email(reason: str, payload: dict) -> None:
    """
    reason ∈ {'unverified_sender','subject_invalid','no_attachments','no_valid_attachments'}
    payload will be appended into the corresponding bucket if not duplicate.
    """
    data = _load_failed_log()
    buckets = data["buckets"]
    totals  = data["totals"]

    if reason not in buckets:
        # future-proof: create unknown reason bucket if needed
        buckets[reason] = []

    bucket = buckets[reason]

    # dedupe by Graph id first, then by internetMessageId if provided
    dedupe_key = payload.get("id") or payload.get("internetMessageId")

    if not _entry_exists(bucket, dedupe_key):
        bucket.append(payload)
  

    data["last_updated_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    _atomic_write_failed_log(data)

# ==========================================================

def dbg(*args, **kwargs):
    if DEBUG:
        print("[DEBUG]", *args, **kwargs)

# ------------------------- Config & Regex -------------------------

SCOPES = ["https://graph.microsoft.com/.default"]

# VERIFIED = {e.lower().strip() for e in VERIFIED_SENDERS}

# Where to store subjects that fail validation (next to this script)
FAILED_SUBJECTS_PATH = Path(__file__).with_name("failed_subjects.json")

# ------------------------- Subject Normalization & Parsing -------------------------

def _strip_date_time_tokens_for_invalid_subject(subj: str) -> str:
    """
    Canonicalize subject for invalid-subject approval:
    remove date/time tokens so date-format-only differences map to the same key.
    """
    if not subj:
        return ""

    # Normalize separators similar to _normalize_subject, but KEEP ':' so time like "14:04"
    # can be stripped before we collapse punctuation.
    subj = unicodedata.normalize("NFKC", str(subj))
    subj = re.sub(r"[;|,/\\]+", " ", subj)
    subj = re.sub(r"\s+", " ", subj).strip()

    # Remove time tokens
    subj = re.sub(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", "", subj)  # HH:MM or HH:MM:SS
    subj = re.sub(r"\b\d{6}\b", "", subj)  # HHMMSS
    subj = re.sub(r"\b(?:am|pm)\b", "", subj, flags=re.IGNORECASE)

    # Remove ISO-8601 datetimes (covers most cases)
    subj = re.sub(
        r"\b(?:19|20)\d{2}[-/\.](?:0?[1-9]|1[0-2])[-/\.](?:0?[1-9]|[12]\d|3[01])"
        r"(?:[T\s]"
        r"(?:[01]\d|2[0-3])[:\.]?[0-5]\d"
        r"(?::?[0-5]\d)?"
        r"(?:\.\d+)?"
        r"(?:\s*(?:Z|[+-](?:[01]\d|2[0-3]):?[0-5]\d))?"
        r")\b",
        "",
        subj,
        flags=re.IGNORECASE,
    )

    # Drop common timezone offset tokens if they appear alone
    # (use whitespace-boundaries, not \b, because '+'/'-' are non-word chars)
    subj = re.sub(r"(?<!\S)[+-](?:[01]\d|2[0-3]):?[0-5]\d(?!\S)", "", subj)

    # Drop day-of-week tokens (common in RFC-like dates)
    subj = re.sub(
        r"\b(?:mon|tue(?:s)?|wed(?:nes)?|thu(?:rs)?|fri|sat(?:ur)?|sun)(?:day)?\b",
        "",
        subj,
        flags=re.IGNORECASE,
    )

    # Drop common timezone tokens that often follow times
    subj = re.sub(r"\b(?:UTC|GMT|EST|EDT|CST|CDT|MST|MDT|PST|PDT)\b", "", subj, flags=re.IGNORECASE)

    # Remove ISO-like / common date formats:
    # - YYYY-MM-DD, YYYY/M/D, YYYY.MM.DD
    subj = re.sub(r"\b(?:19|20)\d{2}[-/.](?:0?[1-9]|1[0-2])[-/.](?:0?[1-9]|[12]\d|3[01])\b", "", subj)
    # - DD-MM-YYYY, D/M/YYYY, DD.MM.YYYY
    subj = re.sub(r"\b(?:0?[1-9]|[12]\d|3[01])[-/.](?:0?[1-9]|1[0-2])[-/.](?:19|20)\d{2}\b", "", subj)
    # - compact yyyymmdd (very common in filenames)
    subj = re.sub(r"\b(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\b", "", subj)
    # - spaced variants after normalization: YYYY M D or D M YYYY
    subj = re.sub(r"\b(?:19|20)\d{2}\s+(?:0?[1-9]|1[0-2])\s+(?:0?[1-9]|[12]\d|3[01])\b", "", subj)
    subj = re.sub(r"\b(?:0?[1-9]|[12]\d|3[01])\s+(?:0?[1-9]|1[0-2])\s+(?:19|20)\d{2}\b", "", subj)
    # - month-name dates like "December 1 2025" / "Dec 1 2025" (commas already normalized away)
    month = r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    subj = re.sub(rf"\b{month}\s+(?:0?[1-9]|[12]\d|3[01])(?:st|nd|rd|th)?\s+(?:19|20)\d{{2}}\b", "", subj, flags=re.IGNORECASE)
    subj = re.sub(rf"\b(?:0?[1-9]|[12]\d|3[01])(?:st|nd|rd|th)?\s+{month}\s+(?:19|20)\d{{2}}\b", "", subj, flags=re.IGNORECASE)
    # - hyphenated month-name forms: 01-Dec-2025, Dec-01-25, 01-Dec-25
    subj = re.sub(rf"\b(?:0?[1-9]|[12]\d|3[01])[-\s]{month}[-\s](?:\d{{2}}|\d{{4}})\b", "", subj, flags=re.IGNORECASE)
    subj = re.sub(rf"\b{month}[-\s](?:0?[1-9]|[12]\d|3[01])[-\s](?:\d{{2}}|\d{{4}})\b", "", subj, flags=re.IGNORECASE)
    # - month + year (e.g. "Dec 2025") and year + month (e.g. "2025 Dec")
    subj = re.sub(rf"\b{month}\s+(?:19|20)\d{{2}}\b", "", subj, flags=re.IGNORECASE)
    subj = re.sub(rf"\b(?:19|20)\d{{2}}\s+{month}\b", "", subj, flags=re.IGNORECASE)

    # If a month-name date was partially stripped (e.g. "01-Dec" left behind), remove the remainder too
    subj = re.sub(rf"\b(?:0?[1-9]|[12]\d|3[01])[-\s]{month}\b", "", subj, flags=re.IGNORECASE)
    subj = re.sub(rf"\b{month}[-\s](?:0?[1-9]|[12]\d|3[01])\b", "", subj, flags=re.IGNORECASE)

    # Remove any leftover standalone '+'/'-' tokens after stripping offsets
    subj = re.sub(r"(?<!\S)[+-](?!\S)", "", subj)

    return re.sub(r"\s+", " ", subj).strip()

def _normalize_subject(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"[:;|,/\\]+", " ", s)            # common separators → space
    s = re.sub(r"[^A-Za-z0-9_\-\s]", " ", s)     # drop ~!@#$%^&*()[]{}<>'"?.+= etc.
    s = re.sub(r"\s+", " ", s).strip()
    return s

# Strict/ordered patterns
_FREEFORM = re.compile(r"""
    ^\s*
    (?P<company>[A-Za-z0-9_\-\s]+?)\s+
    (?P<trunk>[A-Za-z][\w\-]*)
    (?:\s+trunk\b)?\s+prefix\b\s*(?P<prefix>none|\d+)\s+
    (?P<currency>[A-Za-z]{3})\s*$
""", re.IGNORECASE | re.VERBOSE)

_BRACKETED_LIKE = re.compile(r"""
    ^\s*
    (?P<company>[A-Za-z0-9_\-\s]+?)\s+
    (?P<trunk>\S+)\s+
    (?P<prefix>none|\d+)\s+
    (?P<currency>[A-Za-z]{3})\s*$
""", re.IGNORECASE | re.VERBOSE)

_NO_LABELS = re.compile(r"""
    ^\s*
    (?P<company>.+?)\s+
    (?P<trunk>[A-Za-z][\w\-]*)\s+
    (?P<prefix>\d+|none)\s+
    (?P<currency>[A-Za-z]{3})\s*$
""", re.IGNORECASE | re.VERBOSE)

_ANYORDER = re.compile(r"^(?=.*\bprefix\b)(?=.*\b[A-Za-z]{3}\b).*$", re.IGNORECASE)

# Strict subject format required (branch requirement):
#   [company] [trunk] [prefix] [currency]
# Rules:
# - Exactly 4 bracket groups, and nothing else outside them (except whitespace).
# - No bracket may be empty after trimming (so [] / [ ] invalid).
# - Prefix: either "None"/"none" OR must contain at least one digit; we extract the first digit-run
#   and preserve leading zeros (e.g., "Prefix 001" -> "001", "1117#" -> "1117").
# - Currency must be exactly 3 letters (e.g., USD).
_STRICT_4_BRACKETS = re.compile(
    r"""
    ^\s*
    \[(?P<company>[^\[\]]+?)\]\s*
    \[(?P<trunk>[^\[\]]+?)\]\s*
    \[(?P<prefix>[^\[\]]+?)\]\s*
    \[(?P<currency>[A-Za-z]{3})\]\s*
    $
    """,
    re.VERBOSE,
)

# Freeform subject format (allowed as a fallback):
#   <company words> <TRUNK> trunk Prefix:1234 USD
#
# Notes:
# - We apply this to the *normalized* subject (punctuation collapsed to spaces),
#   so "Prefix:1001" and "Prefix 1001" both work.
# - Trunk is the word immediately before the keyword "trunk".
# - Company is everything before that trunk word.
# - Prefix is digits after "prefix".
# - Currency is a 3-letter code at the end.
_FREEFORM_TRUNK_PREFIX = re.compile(
    r"""
    ^\s*
    (?P<company>.+?)\s+
    (?P<trunk>[A-Za-z][\w\-]*)\s+
    trunk\b
    .*?
    prefix\b\s*[:\s-]*\s*(?P<prefix>\d+)
    \s+
    (?P<currency>[A-Za-z]{3})
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

def _first_3letter_currency(tokens):
    for t in reversed(tokens):
        if re.fullmatch(r"[A-Za-z]{3}", t):
            return t.upper()
    return None

def _normalize_output(company: str, trunk: str, prefix, currency: str) -> Optional[Dict[str, object]]:
    """
    Normalize extracted subject parts.
    Notes:
    - company/trunk are preserved (trimmed) because vendor matching uses subject company names.
    - prefix is stored as:
        - None (if explicitly "none"/"None")
        - otherwise the first digit-run string (preserving leading zeros)
    - currency is normalized to uppercase.
    """
    company = (company or "").strip()
    trunk = (trunk or "").strip()
    currency = (currency or "").strip().upper()
    if not company or not trunk or not re.fullmatch(r"[A-Z]{3}", currency):
        return None

    # Normalize prefix
    if isinstance(prefix, int):
        prefix = str(prefix)
    if isinstance(prefix, str):
        p = prefix.strip()
        if not p:
            return None
        # Allow only "NONE" or "PREFIX NONE" (case-insensitive) to mean None.
        # Do NOT accept other phrases like "no prefix".
        if p.lower() == "none" or re.fullmatch(r"prefix\s+none", p, flags=re.IGNORECASE):
            prefix = None
        else:
            m = re.search(r"\d+", p)
            if not m:
                return None
            prefix = m.group(0)  # keep leading zeros
    elif prefix is not None:
        return None

    return {"company": company, "trunk": trunk, "prefix": prefix, "currency": currency}

def _extract_anyorder(subject: str) -> Optional[Dict[str, object]]:
    tokens = subject.split()
    currency = _first_3letter_currency(tokens)
    if not currency:
        return None
    prefix = None
    for i, t in enumerate(tokens):
        if t.lower() == "prefix":
            cand = tokens[i+1] if i + 1 < len(tokens) else ""
            val = cand.lower()
            if val == "none":
                prefix = None
                break
            if re.fullmatch(r"\d+", val):
                prefix = int(val)
                break
    if prefix is None:
        try:
            cur_idx = len(tokens) - 1 - list(reversed(tokens)).index(currency)
            if cur_idx - 1 >= 0 and re.fullmatch(r"\d+", tokens[cur_idx - 1]):
                prefix = int(tokens[cur_idx - 1])
        except ValueError:
            pass
    trunk = None
    for i, t in enumerate(tokens):
        if t.lower() == "trunk":
            if i - 1 >= 0 and re.fullmatch(r"[A-Za-z][\w\-]*", tokens[i-1]):
                trunk = tokens[i-1]
                break
            if i + 1 < len(tokens) and re.fullmatch(r"[A-Za-z][\w\-]*", tokens[i+1]):
                trunk = tokens[i+1]
                break
    if trunk is None:
        for t in tokens:
            if t.lower() in {"prefix", "none", currency.lower()}:
                continue
            if re.fullmatch(r"[A-Za-z][\w\-]*", t):
                trunk = t
                break
    company = None
    if trunk:
        try:
            trunk_idx = tokens.index(trunk)
            head = [w for w in tokens[:trunk_idx] if w.lower() not in {"trunk", "prefix"}]
            company = " ".join(head).strip()
        except ValueError:
            pass
    if not company:
        drop = {currency.lower(), "trunk", "prefix", str(prefix) if prefix is not None else ""}
        rest = [w for w in tokens if w.lower() not in drop and not re.fullmatch(r"\d+", w)]
        company = " ".join(rest).strip()
    return _normalize_output(company, trunk, prefix, currency)

def validate_subject(subject: Optional[str]) -> Optional[Dict[str, object]]:
    if not subject:
        return None
    raw = str(subject).strip()

    # 1) Preferred: strict 4-bracket subjects.
    m = _STRICT_4_BRACKETS.match(raw)
    if m:
        gd = m.groupdict()
        return _normalize_output(
            gd.get("company", ""),
            gd.get("trunk", ""),
            gd.get("prefix", ""),
            gd.get("currency", ""),
        )

    # 2) Fallback: freeform "<company> <trunk> trunk Prefix:123 USD"
    s = _normalize_subject(raw)
    if not s:
        return None
    m2 = _FREEFORM_TRUNK_PREFIX.match(s)
    if not m2:
        return None
    gd2 = m2.groupdict()
    return _normalize_output(
        gd2.get("company", ""),
        gd2.get("trunk", ""),
        gd2.get("prefix", ""),
        gd2.get("currency", ""),
    )

def subject_ok(s: Optional[str]) -> bool:
    return validate_subject(s) is not None

# ------------------------- Utilities -------------------------

def load_env():
    env_path = Path(__file__).with_name(".env")
    load_dotenv(dotenv_path=env_path, override=True)
    missing = [k for k in ["TENANT_ID","CLIENT_ID","CLIENT_SECRET","USER_EMAIL"] if not os.getenv(k)]
    if missing:
        raise SystemExit(f"Missing env vars: {missing}. Check {env_path}")
    return {
        "TENANT_ID": os.getenv("TENANT_ID"),
        "CLIENT_ID": os.getenv("CLIENT_ID"),
        "CLIENT_SECRET": os.getenv("CLIENT_SECRET"),
        "USER_EMAIL": os.getenv("USER_EMAIL"),
        "VERBOSE": os.getenv("VERBOSE","0") == "1",
    }

def b64url_json(seg: str):
    seg = seg + "=" * (-len(seg) % 4)
    return json.loads(base64.urlsafe_b64decode(seg.encode()))

def get_token(tenant_id: str, client_id: str, client_secret: str, verbose=False) -> str:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }

    logger.info(f"Auth token URL: {token_url}")
    try:
        resp = requests.post(token_url, data=data, timeout=(5, 120))  # 5s connect, 120s read
    except requests.exceptions.Timeout as e:
        logger.error(f"Token request timed out: {e}")
        raise
    except requests.exceptions.RequestException as e:
        logger.error(f"Token request failed: {e!r}")
        raise

    if resp.status_code != 200:
        try:
            logger.error(f"Token error: {json.dumps(resp.json(), indent=2)}")
        except Exception:
            logger.error(f"Token error text: {resp.text[:2000]}")
        raise RuntimeError("Failed to get token from Microsoft")

    tok = resp.json()
    access_token = tok.get("access_token")
    if not access_token:
        logger.error(f"Token response missing access_token: {tok}")
        raise RuntimeError("No access_token in token response")

    if verbose or DEBUG:
        parts = access_token.split(".")
        if len(parts) > 1:
            try:
                payload = b64url_json(parts[1])
                logger.info(f"Token roles: {payload.get('roles')}")
                logger.info(f"Token exp (unix): {payload.get('exp')}")
            except Exception as e:
                logger.warning(f"Token payload decode failed: {repr(e)}")

    return access_token

def get_session(access_token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    })
    dbg("Session default headers set.")
    return s

def get_with_retries(session: requests.Session, url: str, timeout=30, max_retries=5) -> requests.Response:
    last = None
    for attempt in range(max_retries):
        dbg(f"HTTP GET attempt {attempt+1}/{max_retries}:", url)
        r = session.get(url, timeout=timeout)
        dbg("HTTP status:", r.status_code)
        if r.status_code in (429, 503, 504):
            wait = int(r.headers.get("Retry-After", "0")) or (2 ** attempt)
            dbg(f"Throttled / transient error. Waiting {wait}s then retrying...")
            time.sleep(min(wait, 30))
            last = r
            continue
        return r
    dbg("Max retries reached. Returning last response.")
    return last if last is not None else r

def iso_range(after: Optional[str], before: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Convert YYYY-MM-DD strings into Graph ISO instants:
      after => start of day Z
      before => end of day Z
    """
    def start_of_day(s: str) -> str:
        dt = datetime.strptime(s, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        return dt.isoformat().replace("+00:00", "Z")
    def end_of_day(s: str) -> str:
        dt = datetime.strptime(s, "%Y-%m-%d").replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        return dt.isoformat().replace("+00:00", "Z")
    a = start_of_day(after) if after else None
    b = end_of_day(before) if before else None
    dbg("ISO range:", a, "→", b)
    return a, b

def build_messages_url(user_email: str, top: int, after_iso: Optional[str], before_iso: Optional[str], unread_only: bool) -> str:
    user_path = quote(user_email)
    base = f"https://graph.microsoft.com/v1.0/users/{user_path}/mailFolders/Inbox/messages"
    selects = "id,internetMessageId,subject,from,receivedDateTime,isRead,hasAttachments"
    order = "receivedDateTime desc"
    filters = []
    if after_iso:
        filters.append(f"receivedDateTime ge {after_iso}")
    if before_iso:
        filters.append(f"receivedDateTime le {before_iso}")
    if unread_only:
        filters.append("isRead eq false")
    filter_q = f"&$filter={' and '.join(filters)}" if filters else ""
    url = f"{base}?$select={selects}&$orderby={order}&$top={top}{filter_q}"
    dbg("Built list URL:", url)
    return url

def safe_filename(name: str) -> str:
    name = os.path.basename(name or "")
    name = unicodedata.normalize('NFKC', name)
    name = re.sub(r'[^A-Za-z0-9._-]+', '_', name).strip('._')
    return name or "attachment"

def unique_path(dirpath: str, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    i = 1
    path = os.path.join(dirpath, filename)
    while os.path.exists(path):
        path = os.path.join(dirpath, f"{base}({i}){ext}")
        i += 1
    return path

def strip_html(html: str) -> str:
    from html import unescape
    html = re.sub(r'(?is)<(script|style).*?>.*?</\1>', '', html)
    html = re.sub(r'(?i)</?(br|p|div|li|tr|h[1-6])[^>]*>', '\n', html)
    text = re.sub(r'(?s)<[^>]+>', '', html)
    text = unescape(text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text).strip()
    return text

# ------------------------- Core Processing -------------------------

def message_sender(msg: dict) -> str:
    frm = (msg.get("from") or {}).get("emailAddress") or {}
    return (frm.get("address") or "").lower().strip()

def save_matching_attachments_for_user(session: requests.Session, user_email: str, msg_id: str,
                                       allowed_exts: Set[str], save_dir: str,
                                       max_bytes: int = 50*1024*1024, min_bytes: int = 1000
                                      ) -> Tuple[bool, int, int, List[str], List[Dict[str, str]]]:
    """
    Returns (saved_any, considered_count, skipped_count, saved_filenames, skip_details)
    skip_details: list of { "name": ..., "ext": ..., "reason": ... }
    """
    user_path = quote(user_email)
    list_url = (
        f"https://graph.microsoft.com/v1.0/users/{user_path}/messages/{quote(msg_id)}/attachments"
        "?$select=id,name,contentType,size"
    )

    saved = False
    considered = 0
    skipped = 0
    saved_files: List[str] = []
    skip_details: List[Dict[str, str]] = []

    if not DRY_RUN:
        os.makedirs(save_dir, exist_ok=True)
    else:
        dbg("DRY_RUN: would create dir", save_dir)

    url = list_url
    while url:
        dbg("Fetching attachments (list):", url)
        r = get_with_retries(session, url)
        if r.status_code != 200:
            try:
                print("Attachment list error:", json.dumps(r.json(), indent=2), file=sys.stderr)
            except Exception:
                print("Attachment list error text:", r.text[:2000], file=sys.stderr)
            r.raise_for_status()

        data = r.json()
        atts = data.get("value", [])
        dbg(f"[DEBUG] Attachments returned: {len(atts)}")

        for att in atts:
            considered += 1
            raw_name = att.get("name") or "attachment"
            filename = safe_filename(raw_name)
            _, ext = os.path.splitext(filename)
            size_meta = att.get("size")

            att_id = att.get("id")
            if not att_id:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "missing_attachment_id"})
                continue

            att_url = f"https://graph.microsoft.com/v1.0/users/{user_path}/messages/{quote(msg_id)}/attachments/{quote(att_id)}"
            dbg("   GET attachment detail:", att_url)
            r2 = get_with_retries(session, att_url)
            if r2.status_code != 200:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": f"http_{r2.status_code}"})
                continue

            fatt = r2.json()
            otype = fatt.get("@odata.type", "")

            if "fileAttachment" not in otype:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "not_fileAttachment"})
                continue

            if ext.lower() not in allowed_exts:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "extension_not_allowed"})
                continue

            b64 = fatt.get("contentBytes")
            if not b64:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "missing_contentBytes"})
                continue

            try:
                blob = base64.b64decode(b64)
            except Exception:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "base64_decode_failed"})
                continue

            sz = len(blob)
            if sz < min_bytes:
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "too_small"})
                continue
            if sz > max_bytes:
                print(f"   ↳ skipped {filename}: {sz} bytes exceeds limit")
                skipped += 1
                skip_details.append({"name": raw_name, "ext": ext.lower(), "reason": "too_large"})
                continue

            path = unique_path(save_dir, filename)
            if DRY_RUN:
                dbg("   -> DRY_RUN: would save to", path)
            else:
                with open(path, "wb") as f:
                    f.write(blob)
                print(f"   ↳ saved attachment: {path}")
                saved_files.append(os.path.basename(path))
            saved = True

        url = data.get("@odata.nextLink")
        dbg("Next attachments page?", bool(url))

    return saved, considered, skipped, saved_files, skip_details

def fetch_message_body_text(session: requests.Session, user_email: str, msg_id: str) -> str:
    user_path = quote(user_email)
    url = f"https://graph.microsoft.com/v1.0/users/{user_path}/messages/{quote(msg_id)}?$select=body,bodyPreview"
    dbg("Fetching body:", url)
    r = get_with_retries(session, url)
    if r.status_code != 200:
        try:
            print("Body fetch error:", json.dumps(r.json(), indent=2), file=sys.stderr)
        except Exception:
            print("Body fetch error text:", r.text[:2000], file=sys.stderr)
        r.raise_for_status()
    msg = r.json()
    body = (msg.get("body") or {})
    content_type = (body.get("contentType") or "").lower()
    content = body.get("content") or ""
    dbg("Body contentType:", content_type, "| content length:", len(content))
    if not content:
        preview = (msg.get("bodyPreview") or "")
        dbg("Using bodyPreview length:", len(preview))
        return preview
    if content_type == "html":
        stripped = strip_html(content)
        dbg("HTML stripped length:", len(stripped))
        return stripped
    return content

def write_metadata(save_dir: str, meta: Dict[str, object]) -> None:
    """Write metadata.json in save_dir (idempotent overwrite)."""
    if DRY_RUN:
        dbg("DRY_RUN: would write metadata:", os.path.join(save_dir, "metadata.json"))
        dbg("DRY_RUN meta:", meta)
        return
    try:
        os.makedirs(save_dir, exist_ok=True)
        path = os.path.join(save_dir, "metadata.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(f"  ↳ wrote metadata: {path}")
    except Exception as e:
        print(f"  (warn) failed to write metadata for {save_dir}: {e}", file=sys.stderr)

# ___________________ Failed Emails subjects handeling ___________________
def _read_json_dict(path: Path) -> Dict[str, object]:
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception as e:
        print(f"(warn) could not read {path}: {e}", file=sys.stderr)
    return {}

def _atomic_write_json(path: Path, data: Dict[str, object]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def log_failed_subject(sender_email: str, raw_subject: str) -> None:
    """Append the failed subject under the sender key. Values are lists to avoid overwrites."""
    if DRY_RUN:
        dbg(f"DRY_RUN: would log failed subject for {sender_email!r}: {raw_subject!r}")
        return
    data = _read_json_dict(FAILED_SUBJECTS_PATH)
    current = data.get(sender_email)
    if current is None:
        data[sender_email] = [raw_subject]
    elif isinstance(current, list):
        if raw_subject not in current:
            current.append(raw_subject)
    else:
        # normalize pre-existing non-list value
        if current != raw_subject:
            data[sender_email] = [current, raw_subject]
    try:
        _atomic_write_json(FAILED_SUBJECTS_PATH, data)
    except Exception as e:
        print(f"(warn) failed to update {FAILED_SUBJECTS_PATH}: {e}", file=sys.stderr)
#_________________________________

def process_inbox(session: requests.Session, user_email: str, after: Optional[str], before: Optional[str],
                  page_size: int, allowed_exts: Set[str], attachments_base: str, unread_only: bool, verified_set: Set[str]):
    after_iso, before_iso = iso_range(after, before)
    url = build_messages_url(user_email, page_size, after_iso, before_iso, unread_only)

    total_fetched = 0
    total_pages = 0
    matched_messages = 0
    saved_messages = 0
    skipped_sender = skipped_subject = skipped_no_attach = skipped_ext = 0
    skipped_read = 0
    skipped_hayo_replace = 0  # counter for HAYO FULL A-Z REPLACE ORIG subjects
    skipped_existing_dir = 0  # counter for idempotent skip

    print(f"Query: after={after or '(none)'} | before={before or '(none)'} | page size={page_size} | unread_only={unread_only}")
    print(f"Allowed filetypes: {', '.join(sorted(allowed_exts))}")
    print(f"Verified senders count: {len(verified_set)}")
    print("-" * 60)

    while url:
        total_pages += 1
        print(f"\n--- PAGE {total_pages} ------------------------------------------------")
        r = get_with_retries(session, url)
        if r.status_code != 200:
            try:
                print("HTTP error payload:", json.dumps(r.json(), indent=2), file=sys.stderr)
            except Exception:
                print("HTTP error text:", r.text[:2000], file=sys.stderr)
            r.raise_for_status()

        data = r.json()
        items = data.get("value", [])
        page_count = len(items)
        total_fetched += page_count
        print(f"Fetched {page_count} message(s) on this page.")

        if not items:
            print("(No messages returned on this page.)")


        for i, m in enumerate(items, 1):
            msg_id = m.get("id")
            internet_msg_id = m.get("internetMessageId")  # NEW: for stable dedupe/audit
            sender = message_sender(m)
            subject = m.get("subject") or ""
            dt_str = m.get("receivedDateTime") or ""
            has_attachments = bool(m.get("hasAttachments"))
            is_unread = not m.get("isRead")

            # ensure this exists for later metadata (even if not used)
            override_table_name: Optional[str] = None
            source = "regex"

            # 1) unread gate
            if unread_only and not is_unread:
                print("  -> skip: message is read but unread_only=True")
                skipped_read += 1
                continue

            # 2) ignore emails with subject containing "HAYO FULL A-Z REPLACE ORIG" (including reply prefixes)
            subject_upper = subject.upper()
            if (subject_upper.startswith("HAYO FULL A-Z REPLACE ORIG") or 
                "HAYO FULL A-Z REPLACE ORIG" in subject_upper):
                print(f"  -> skip: subject contains 'HAYO FULL A-Z REPLACE ORIG': {subject!r}")
                skipped_hayo_replace += 1
                continue

            # 3) unverified sender
            if sender not in verified_set:
                print("  -> skip: sender NOT in VERIFIED_SENDERS")
                skipped_sender += 1
                try:
                    log_failed_email("unverified_sender", {
                        "id": msg_id,
                        "internetMessageId": internet_msg_id,
                        "sender": sender,
                        "subject": subject,
                        "receivedDateTime": dt_str,
                        "logged_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                        "details": { "verified_list_size": len(verified_set) }
                    })
                except Exception as e:
                    print(f"(warn) failed to update failed_emails.json: {e}", file=sys.stderr)
                continue

            # 4) subject invalid → DB handling/override support
            parsed = validate_subject(subject)
            if not parsed:
                print(f"  -> subject does not match required fields: {subject!r}")
                # JSON logs you already have
                try:
                    log_failed_subject(sender, subject)
                except Exception as e:
                    print(f"  (warn) failed to log failed subject: {e}", file=sys.stderr)
                skipped_subject += 1
                try:
                    log_failed_email("subject_invalid", {
                        "id": msg_id,
                        "internetMessageId": internet_msg_id,
                        "sender": sender,
                        "subject": subject,
                        "receivedDateTime": dt_str,
                        "logged_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                        "details": { "note": "did not match required fields" }
                    })
                except Exception as e:
                    print(f"(warn) failed to update failed_emails.json: {e}", file=sys.stderr)

                # DB flow: allow override if approved
                try:
                    from database import (
                        get_or_create_invalid_subject,
                        insert_invalid_subject_detail,
                        find_invalid_subject_detail,
                    )
                    try:
                        rcvd_dt = datetime.fromisoformat((dt_str or "").replace("Z","+00:00")) if dt_str else None
                    except Exception:
                        rcvd_dt = None

                    parent_id = get_or_create_invalid_subject(
                        email=sender or "",
                        received_at=rcvd_dt,
                        processed_at=datetime.now(timezone.utc),
                        status="pending"
                    )

                    def strip_date_from_subject(subj: str) -> str:
                        return _strip_date_time_tokens_for_invalid_subject(subj)
                    
                    subject_nodate = strip_date_from_subject(subject)
                    found = find_invalid_subject_detail(parent_id, subject_nodate)
                    if found:
                        detail_id, jera_table_name = found
                        if jera_table_name:
                            # APPROVED via override → proceed normally (do NOT continue)
                            override_table_name = jera_table_name
                            print(f"  -> override APPROVED via invalid_subject_details id={detail_id}, table={jera_table_name}")
                        else:
                            print(f"  -> subject already logged but no JeraSoft table yet; waiting for approval")
                            continue  # still pending approval
                    else:
                        # New subject for this sender → record & wait for approval
                        _ = insert_invalid_subject_detail(parent_id, subject_nodate, None)
                        print(f"  -> logged new invalid subject for approval (date ignored)")
                        continue

                except Exception as e:
                    print(f"(warn) invalid-subject DB handling failed: {e}", file=sys.stderr)
                    continue  # fail-safe: skip for now

            # 5) no attachments flag
            if not has_attachments:
                print("  -> skip: no attachments (hasAttachments=False)")
                skipped_no_attach += 1
                try:
                    log_failed_email("no_attachments", {
                        "id": msg_id,
                        "internetMessageId": internet_msg_id,
                        "sender": sender,
                        "subject": subject,
                        "receivedDateTime": dt_str,
                        "logged_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                        "details": { "hasAttachments_flag": False }
                    })
                except Exception as e:
                    print(f"(warn) failed to update failed_emails.json: {e}", file=sys.stderr)
                continue

            # --- Passed sender, subject (or approved override), and has_attachments ---
            matched_messages += 1

            # Directory name from sender + UTC timestamp
            try:
                dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            except Exception as e:
                dbg("datetime parse failed:", repr(e), "raw:", dt_str)
                dt = datetime.now(timezone.utc)

            date_time_str = dt.astimezone(timezone.utc).strftime('%Y%m%d_%H%M%S')
            date_only = dt.astimezone(timezone.utc).strftime('%Y-%m-%d')
            time_only = dt.astimezone(timezone.utc).strftime('%H:%M:%S')
            safe_sender = sender.replace('@', '_at_')
            save_dir = os.path.join(attachments_base, f"{safe_sender}_{date_time_str}")
            print("  save_dir:", save_dir)

            # Idempotency: if this exact message-dir already exists, skip
            if os.path.isdir(save_dir):
                print("  -> skip: directory already exists, assuming this message was processed before")
                skipped_existing_dir += 1
                continue

            # Save attachments
            saved_any, considered_count, skipped_count, saved_files, skip_details = save_matching_attachments_for_user(
                session, user_email, msg_id, allowed_exts, save_dir
            )
            print(f"  attachments considered: {considered_count} | skipped: {skipped_count} | saved_any={saved_any}")

            # 6) none valid after checking
            if not saved_any:
                print("  -> skip: no attachment passed extension/size checks")
                skipped_ext += 1
                # clean empty dir
                try:
                    if os.path.isdir(save_dir) and not os.listdir(save_dir):
                        os.rmdir(save_dir)
                        dbg("  (cleanup) removed empty dir:", save_dir)
                except Exception:
                    pass
                try:
                    log_failed_email("no_valid_attachments", {
                        "id": msg_id,
                        "internetMessageId": internet_msg_id,
                        "sender": sender,
                        "subject": subject,
                        "receivedDateTime": dt_str,
                        "logged_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                        "details": {
                            "considered": considered_count,
                            "skipped": skipped_count,
                            "saved": 0,
                            "allowed_exts": sorted(list(allowed_exts)),
                            "size_limits_bytes": {"min": 1000, "max": 50*1024*1024},
                            "per_attachment": skip_details
                        }
                    })
                except Exception as e:
                    print(f"(warn) failed to update failed_emails.json: {e}", file=sys.stderr)
                continue

            # Fetch body text (optional for logs)
            body_text = fetch_message_body_text(session, user_email, msg_id)

            # Write metadata.json
            meta = {
                "subject": subject,
                "sender": sender,
                "company": (parsed.get("company") if parsed else None),
                "prefix": (parsed.get("prefix") if parsed else None),
                "date_utc": date_only,
                "time_utc": time_only,
                "directory": os.path.abspath(save_dir),
                "internetMessageId": internet_msg_id,
                "attachment_count": len(saved_files),
                "attachments": saved_files,
                "receivedDateTime_raw": dt_str,
                "processed_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                "subject_parsing_source": source
            }
            if override_table_name:
                meta["force_jerasoft_table_name"] = override_table_name  # <-- NEW

            write_metadata(save_dir, meta)

            print("\n✅ VERIFIED EMAIL (with allowed attachment)")
            print(f"From: {sender}")
            print(f"Subject: {subject}")
            print(f"Received: {dt_str}")
            print(f"Body (first 300 chars):\n{(body_text or '')[:300]}")
            print("-" * 60)
            saved_messages += 1

        url = data.get("@odata.nextLink")
        dbg("Next page link present?" , bool(url))

    print("\n======================== SUMMARY ========================")
    print(f"Pages fetched:          {total_pages}")
    print(f"Messages fetched total: {total_fetched}")
    print(f"Matched messages:       {matched_messages}")
    print(f"Saved messages:         {saved_messages}")
    print(f"Skipped (sender):       {skipped_sender}")
    print(f"Skipped (HAYO replace): {skipped_hayo_replace}")
    print(f"Skipped (subject):      {skipped_subject}")
    print(f"Skipped (no attach):    {skipped_no_attach}")
    print(f"Skipped (ext/size):     {skipped_ext}")
    print(f"Skipped (read):         {skipped_read}")
    print(f"Skipped (existing dir): {skipped_existing_dir}")
    print("========================================================\n")

# ------------------------- Main (hardcoded settings) -------------------------

def verify_fetch_emails(after: str, before: str, unread_only: bool = True) -> None:
    cfg = load_env()

    # build the verified set here (fresh every run)
    verified_senders = get_verified_senders()    
    verified_set = {e.lower().strip() for e in verified_senders}
   
    page_size = 50                    # number of messages per API call
    filetypes = ".csv,.xlsx,.xls"     # allowed file extensions
    attachments_dir = "attachments"   # base directory where attachments are saved
              # only process unread emails

    # Normalize extensions
    allowed_exts = {e.strip().lower() for e in filetypes.split(",") if e.strip()}
    if not allowed_exts:
        allowed_exts = {".csv", ".xlsx"}

    token = get_token(cfg["TENANT_ID"], cfg["CLIENT_ID"], cfg["CLIENT_SECRET"], verbose=cfg["VERBOSE"])
    session = get_session(token)

    print(f"Querying Inbox for: {cfg['USER_EMAIL']}")
    print(f"VERIFIED_SENDERS count: {len(verified_set)}")
    if DEBUG:
        print("VERIFIED_SENDERS sample:", list(sorted(verified_set))[:10], "..." if len(verified_set) > 10 else "")

    process_inbox(
        session=session,
        user_email=cfg["USER_EMAIL"],
        after=after,
        before=before,
        page_size=page_size,
        allowed_exts=allowed_exts,
        attachments_base=attachments_dir,
        unread_only=unread_only,
        verified_set=verified_set
    )

if __name__ == "__main__":
    after = "2025-08-19"              # only include emails on/after this date (YYYY-MM-DD) or None
    before = "2025-08-19"             # only include emails on/before this date (YYYY-MM-DD) or None
    unread_only = False    
    verify_fetch_emails(after, before, unread_only)
