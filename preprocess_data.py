import pandas as pd, os, re, numpy as np
from datetime import datetime
from dateutil import parser as dparse
from openpyxl import load_workbook
import re, unicodedata

REQUIRED_COLS = [
    'Dst Code', 'Rate', 'Effective Date', 'Billing Increment'
]

ROW_FALLBACK_THRESHOLD = 100  # trigger pandas fallback if streaming returns < 100 rows


BILLING_PAIRS = [
    ('initial_period', 'recurring_period'),
    ('initial_period', 'subsequent_period'),
    ('initial_increment', 'next_increment'),
    ('min_bill', 'billing_step'),
    ('first_increment', 'second_increment'),
    ('first_increment', 'additional_increment'),
    ('interval_1', 'interval_n'),
    ('min_volume','interval'),
    ('mc', 'ci'),
    ('rounding', 'rounding'),
    ('min' ,'inc'),
]

EXCEL_EPOCH = datetime(1899, 12, 30)          # Excel’s epoch (PC versions)

DEBUG = True
def dbg(*a, **k):
    if DEBUG:
        print(*a, **k)

def _codepoints(s):
    s = str(s)
    return " ".join(f"U+{ord(ch):04X}" for ch in s)


# ─────────────────────────── Dst Code expander ───────────────────────────

_DASHES_RE = re.compile(r'[–—−]')  # normalize en/em/minus to simple "-"

def _parse_dst_code_list(code_str: str) -> list[str]:
    """Split 'Dst Code' strings on comma/semicolon; expand hyphen ranges inclusive.
       Treat everything that isn't a digit or '-' as noise. Preserve zero padding."""
    if pd.isna(code_str):
        return []
    s = str(code_str).strip()
    if not s:
        return []

    s = _DASHES_RE.sub('-', s)  # unify weird dashes
    parts = re.split(r'[;,]', s)  # comma/semicolon = separate codes

    out: list[str] = []
    for part in parts:
        token = re.sub(r'[^0-9\-]', '', part.strip())  # keep digits and hyphen only
        if not token:
            continue

        if '-' in token:
            lo_str, hi_str = token.split('-', 1)
            if lo_str.isdigit() and hi_str.isdigit():
                lo, hi = int(lo_str), int(hi_str)
                if hi < lo:  # be forgiving if they reversed it
                    lo, hi = hi, lo
                width = max(len(lo_str), len(hi_str))
                out.extend(f"{n:0{width}d}" for n in range(lo, hi + 1))
                continue

        # single code
        out.append(token)

    return out

def expand_dst_code_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Explode rows that have multiple/ranged DST codes into one row per code."""
    out = df.copy()
    out['Dst Code'] = out['Dst Code'].astype(str).str.strip()
    out['__dst_list'] = out['Dst Code'].apply(_parse_dst_code_list)
    out = out.explode('__dst_list', ignore_index=True)
    out['Dst Code'] = out['__dst_list'].fillna('')
    out = out[out['Dst Code'] != ''].drop(columns='__dst_list').reset_index(drop=True)
    return out

# ---- Header pre-cleaner: strip currency symbols/codes and other noise ----

_CURRENCY_SYMBOLS_RE = r'[$£€¥₹₩₽₺₫₪₴₦₱₲₵₡₭฿₮₸₼]'
_CURRENCY_CODES = (
    "usd|eur|gbp|jpy|cny|rmb|cad|aud|nzd|inr|chf|sar|aed|rub|brl|mxn|zar|"
    "sek|nok|dkk|pln|huf|try|ils|krw|idr|myr|thb|php|vnd|uah|ron|czk|ars|"
    "clp|cop|pen|twd|hkd|sgd|ngn|kzt"
)
_CURRENCY_CODES_RE = re.compile(rf"(?i)\b(?:{_CURRENCY_CODES})\b")
_CURRENCY_SYMBOLS_RE = re.compile(r"[$£€¥₹₩₽₺₫₪₴₦₱₲₵₡₭฿₮₸₼]")

def _preclean_header_token(s: str) -> str:
    """Strip currency symbols and common junk from HEADER labels."""
    s = str(s).replace("\n", " ")
    s = _CURRENCY_SYMBOLS_RE.sub(" ", s)          # $, €, etc.
    s = re.sub(r"\(.*?\)", " ", s)                # drop parentheticals like (USD)
    # 🚫 DO NOT blindly strip currency codes from whole header like 'USD'
    # s = _CURRENCY_CODES_RE.sub(" ", s)          # ← remove this line or comment it out
    s = re.sub(r"(?i)\bper\s*(min(?:ute)?|sec(?:ond)?)\b", " ", s)
    s = re.sub(r"[/\\|:]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _strip_currency_words_from_key(key: str) -> str:
    """
    Remove currency codes only when they are part of a longer key (rate_usd),
    but if the entire key is just 'usd', keep it so it maps to Rate.
    """
    # If the whole key is exactly a currency code, keep it
    if re.fullmatch(rf"(?:{_CURRENCY_CODES})", key, flags=re.I):
        return key

    # Otherwise strip currency suffixes/prefixes
    cleaned = re.sub(
        rf"(?:^|_)(?:{_CURRENCY_CODES})(?=_|$)", 
        "", 
        key, 
        flags=re.I
    )
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")

    # If stripping removed everything, keep original key
    return cleaned or key



#________________________________handeling seperate/duplicate billing increment columns ──────────────────────────────────

def _coalesce_billing_increment_dupes(df: pd.DataFrame, col_name: str = 'Billing Increment') -> pd.DataFrame:
    """
    If there are 2+ columns named `col_name`, combine the FIRST TWO into a single 'a/b' string per row:
      - a = last integer from the first dup column
      - b = last integer from the second dup column
      - If only one side exists on a row, duplicate it as a/a or b/b
    Drops all original dup columns and inserts a single consolidated column at the
    position of the first duplicate. No-op if <2 duplicates exist.
    """
    # locate duplicates by position, not by label
    mask = (df.columns == col_name)
    idxs = np.flatnonzero(mask)
    if len(idxs) < 2:
        return df  # nothing to do

    # grab the first two duplicate columns by position
    block = df.iloc[:, idxs]
    left_raw = block.iloc[:, 0]
    right_raw = block.iloc[:, 1]

    # reuse your existing number extractor
    def _num_last(x) -> str:
        return _last_num(x)  # already defined in your file

    left = left_raw.map(_num_last).fillna('')
    right = right_raw.map(_num_last).fillna('')

    # build the consolidated a/b values
    out = pd.Series('', index=df.index, dtype='object')
    both = (left != '') & (right != '')
    only_l = (left != '') & (right == '')
    only_r = (left == '') & (right != '')

    out.loc[both] = left[both] + '/' + right[both]
    out.loc[only_l] = left[only_l] + '/' + left[only_l]
    out.loc[only_r] = right[only_r] + '/' + right[only_r]
    # rows where both sides empty remain ''

    # drop all dup columns, insert the consolidated one at the original first position
    first_pos = int(idxs[0])
    df = df.drop(columns=block.columns)
    df.insert(min(first_pos, len(df.columns)), col_name, out)

    return df


def _make_unique_header_names(headers: list[str]) -> list[str]:
    """
    Make header names unique while preserving normalization/alias matching.

    We suffix duplicates as " (2)", " (3)", ... and rely on `_preclean_header_token`
    stripping parentheticals so aliasing still sees the original base header text.

    Example:
      ["BILLING TERMS", "BILLING TERMS"] -> ["BILLING TERMS", "BILLING TERMS (2)"]
    """
    seen: dict[str, int] = {}
    out: list[str] = []
    for h in headers:
        base = "" if h is None else str(h)
        n = seen.get(base, 0) + 1
        seen[base] = n
        out.append(base if n == 1 else f"{base} ({n})")
    return out


def _last_num(s: str) -> str:
    m = re.findall(r'\d+', str(s))
    return m[-1] if m else ''

def _synthesize_billing_increment(df: pd.DataFrame) -> pd.DataFrame:

    df = _coalesce_billing_increment_dupes(df)

    # If Billing Increment is already present and non-empty anywhere, we normally keep it.
    # But if we ALSO have known "pair" columns (e.g., Interval 1 / Interval N), that is ambiguous:
    # reject the file with a clear error instead of silently ignoring the extra columns.
    if 'Billing Increment' in df.columns and df['Billing Increment'].astype(str).str.strip().ne('').any():
        norm2real = {_norm(c): c for c in df.columns}
        pair_present = [(a, b) for (a, b) in BILLING_PAIRS if a in norm2real and b in norm2real]
        if pair_present:
            # show the real column names found
            cols = []
            for a, b in pair_present:
                cols.append(norm2real[a])
                cols.append(norm2real[b])
            cols = sorted(set(cols))
            raise ValueError(
                "Ambiguous Billing Increment sources: found 'Billing Increment' column with values, "
                f"and also found billing-increment pair columns {cols}. "
                "Please keep only one source (either provide a single Billing Increment column, "
                "OR provide the pair columns like Interval 1/Interval N and remove Billing Increment)."
            )
        return df
    
    print("\n\n\nSynthesizing 'Billing Increment' from other columns...\n\n\n")

    # map normalized -> actual column name
    norm2real = { _norm(c): c for c in df.columns }

    pairs = [
        ('initial_period', 'recurring_period'),
        ('initial_period', 'subsequent_period'),
        ('initial_increment', 'next_increment'),
        ('min_bill', 'billing_step'),
        ('first_increment', 'second_increment'),
        ('first_increment', 'additional_increment'),
        ('interval_1', 'interval_n'),
        ('min_volume','interval'),
        ('mc', 'ci'),
        ('rounding', 'rounding'),
        ('min' ,'inc'),
    ]

    for a, b in pairs:
        if a in norm2real and b in norm2real:
            ax = df[norm2real[a]].map(_last_num)
            bx = df[norm2real[b]].map(_last_num)
            mask = (ax != '') & (bx != '')
            if mask.any():
                df['Billing Increment'] = np.where(mask, ax + '/' + bx, '')
                return df

    # fallback: single-column duplication if we only have one increment-ish column
    # Extract all possible singles from pairs dynamically
    potential_singles = set()
    for a, b in pairs:
        potential_singles.add(a)
        potential_singles.add(b)
    
    # Add explicit singles
    explicit_singles = ['increment']
    potential_singles.update(explicit_singles)
    
    for k in potential_singles:
        if k in norm2real:
            col_values = df[norm2real[k]].astype(str).str.strip()
            
            # Check if values already contain "/" format (like "1/1", "60/60")
            has_slash_format = col_values.str.contains(r'\d+/\d+', regex=True, na=False)
            
            if has_slash_format.any():
                # Use existing slash format directly where it exists
                df['Billing Increment'] = np.where(
                    has_slash_format & (col_values != ''), 
                    col_values, 
                    ''
                )
                
                # For values without slash format, extract number and duplicate
                one = col_values.map(_last_num)
                no_slash_mask = ~has_slash_format & (one != '')
                if no_slash_mask.any():
                    df.loc[no_slash_mask, 'Billing Increment'] = one[no_slash_mask] + '/' + one[no_slash_mask]
                
                return df
            else:
                # Original logic: extract last number and duplicate
                one = col_values.map(_last_num)
                if one.ne('').any():
                    df['Billing Increment'] = np.where(one != '', one + '/' + one, '')
                    return df

    # ensure the column exists so downstream selection doesn't explode
    df['Billing Increment'] = ''
    return df

#________________________________ read raw ──────────────────────────────────

# def _raw_from_ws(ws) -> pd.DataFrame:
#     try:
#         # Try to calculate dimensions without forcing
#         ws.calculate_dimension()
#     except ValueError:
#         # If it fails, use calculate_dimension with force=True
#         print("DEBUG: Extraction failed without force; retrying with force=True")
#         ws.calculate_dimension(force=True)

#     rows = list(ws.iter_rows(values_only=True))
#     raw = pd.DataFrame(rows)
#     raw.dropna(how="all", inplace=True)
#     raw.dropna(axis=1, how="all", inplace=True)
#     raw.reset_index(drop=True, inplace=True)
#     return raw.astype("string")


# def _raw_from_ws(ws) -> pd.DataFrame:
#     # read cells, not values_only
#     rows_as_text = []
#     for row in ws.iter_rows(values_only=False):
#         out = []
#         for c in row:
#             v = c.value
#             # if it's a datetime/date, render as M/D/YYYY (unpadded like Excel often shows)
#             if hasattr(c, "is_date") and c.is_date and isinstance(v, (datetime, )):
#                 out.append(f"{v.month}/{v.day}/{v.year}")
#             else:
#                 out.append("" if v is None else str(v))
#         rows_as_text.append(out)

#     raw = pd.DataFrame(rows_as_text)
#     raw.dropna(how="all", inplace=True)
#     raw.dropna(axis=1, how="all", inplace=True)
#     raw.reset_index(drop=True, inplace=True)
#     return raw.astype("string")


from datetime import date, datetime

def _raw_from_ws(ws) -> pd.DataFrame:
    rows_as_text = []
     # Read the entire worksheet without an artificial row cap.
    # Keep a reasonable empty-row streak break to avoid trailing empties in very large sheets.
    empty_row_limit = 500  # Stop if 500 consecutive empty rows (acts near end-of-data)
    empty_count = 0
    for row in ws.iter_rows(values_only=False):
        out = []
        for c in row:
            v = c.value
            if v is None:
                out.append("")
            else:
                out.append(str(v))
        # Skip all-empty rows
        if all(cell.strip() == "" for cell in out):
            empty_count += 1
            if empty_count >= empty_row_limit:
                break
            continue
        else:
            empty_count = 0
        rows_as_text.append(out)

    raw = pd.DataFrame(rows_as_text)
    raw.dropna(how="all", inplace=True)
    raw.dropna(axis=1, how="all", inplace=True)
    raw.reset_index(drop=True, inplace=True)
    return raw.astype("string")

# ── PATCH 1: make the pandas reader actually try multiple engines & the target sheet ──
def _raw_from_excel_pandas(path: str, sheet) -> pd.DataFrame:
    """
    Read a sheet using pandas.read_excel with multiple engines.
    Returns a raw grid (no headers assigned), dropping all-empty rows/columns.
    """
    engines = ("calamine", "openpyxl")  # try calamine first; fall back to openpyxl
    last_exc = None
    for eng in engines:
        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet if sheet is not None else 0,
                engine=eng,
                header=None,     # we want a raw grid
                dtype=str
            )
            # If pandas returns a dict (when sheet_name=None), normalize to the first sheet
            if isinstance(df, dict):
                # pick the first item deterministically
                first_key = sorted(df.keys())[0]
                df = df[first_key]

            df = pd.DataFrame(df)
            df.dropna(how="all", inplace=True)
            df.dropna(axis=1, how="all", inplace=True)
            df.reset_index(drop=True, inplace=True)
            return df.astype("string")
        except Exception as e:
            dbg(f"[pandas engine={eng}] read failed: {e}")
            last_exc = e
            continue
    raise last_exc if last_exc else RuntimeError("read_excel failed with both calamine and openpyxl")

# ── PATCH 2: strengthen _read_raw_matrix to try pandas when there are 0 worksheets
#             and as a final fallback when header detection fails on all sheets. ──
def _read_raw_matrix(path: str, sheet=0) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm', '.xls'):
        # Optional: route legacy .xls straight to pandas, since openpyxl can't read BIFF .xls
        if ext == '.xls':
            try:
                print("DEBUG: Trying the read excel pandas method (.xls fast path)")
                raw_pd = _raw_from_excel_pandas(path, sheet if sheet is not None else 0)
                print("DEBUG: Gotcha using the read excel pandas method (.xls fast path)")
                return raw_pd
            except Exception:
                # fall through to the openpyxl loop anyway, in case the file is misnamed
                pass

        print("DEBUG: Trying the workbook method")
        wb = load_workbook(path, data_only=True, read_only=True)

        # If there are ZERO normal worksheets (chartsheet-only, corrupted, or placeholder),
        # try a pandas fallback immediately.
        if len(wb.worksheets) == 0:
            dbg("[excel] 0 worksheets detected by openpyxl → trying pandas fallback")
            try:
                raw_pd = _raw_from_excel_pandas(path, sheet if sheet is not None else 0)
                wb.close()
                return raw_pd
            except Exception as e:
                wb.close()
                raise ValueError(f"No worksheets found and pandas fallback failed: {e}")

        # try requested sheet first, then all others
        try_order = []
        if isinstance(sheet, int) and 0 <= sheet < len(wb.worksheets):
            try_order.append(sheet)
        elif isinstance(sheet, str):
            try_order += [i for i, ws in enumerate(wb.worksheets) if ws.title == sheet]
        try_order += [i for i in range(len(wb.worksheets)) if i not in try_order]

        # helper: pandas-all-sheets fallback (even when row count >= threshold)
        last_header_err: str | None = None

        def _try_pandas_all_sheets():
            nonlocal last_header_err
            for j in range(len(wb.worksheets)):
                try:
                    raw_pd = _raw_from_excel_pandas(path, j)
                    try:
                        _ = detect_header_row(raw_pd)
                        dbg(f"[excel-fallback] pandas finally found headers on sheet #{j}")
                        return raw_pd
                    except ValueError as e:
                        # remember the most recent, more detailed header error
                        last_header_err = str(e)
                        continue
                except Exception as e:
                    dbg(f"[excel-fallback] pandas read failed on sheet #{j}: {e}")
            return None

        for i in try_order:
            raw_stream = _raw_from_ws(wb.worksheets[i])

            # Fallback trigger BEFORE header detection if the streaming read seems too short
            chosen = raw_stream
            if raw_stream.shape[0] < ROW_FALLBACK_THRESHOLD:
                try:
                    raw_pd = _raw_from_excel_pandas(path, i)  # same sheet index
                    if raw_pd.shape[0] > raw_stream.shape[0]:
                        dbg(f"[excel-fallback] streaming rows={raw_stream.shape[0]} < {ROW_FALLBACK_THRESHOLD}; "
                            f"pandas rows={raw_pd.shape[0]} -> using pandas for sheet #{i}")
                        chosen = raw_pd
                    else:
                        dbg(f"[excel-fallback] streaming rows={raw_stream.shape[0]} < {ROW_FALLBACK_THRESHOLD}; "
                            f"pandas rows={raw_pd.shape[0]} not larger -> keeping streaming for sheet #{i}")
                except Exception as e:
                    dbg(f"[excel-fallback] pandas read failed on sheet #{i}: {e}")

            try:
                print('\n\nDEBUG: Calling the detect header row function from read_raw_matrix\n\n')
                _ = detect_header_row(chosen)  # will raise if not found
                wb.close()
                print("\n\nDEBUG: Gotcha using the read excel openpyxl method")
                return chosen  # this sheet has the headers; use it
            except ValueError as e:
                # remember the most recent, more detailed header error so we
                # can surface it if all sheets fail.
                last_header_err = str(e)
                print("\n\nDEBUG: Failed the detect_header_row check")
                # dont return yet; try other sheets
                continue

        # If we exhausted all openpyxl sheets without finding headers, try pandas across all sheets
        dbg("[excel] openpyxl could not find headers on any sheet → trying pandas across all sheets")
        raw_pd_any = _try_pandas_all_sheets()
        wb.close()
        if raw_pd_any is not None:
            return raw_pd_any

        # If we have a detailed header error from detect_header_row, prefer
        # that over a generic message so callers/meta/status can show exactly
        # which canonical columns were missing.
        if last_header_err:
            raise ValueError(last_header_err)

        raise ValueError("Header not found on any sheet (openpyxl and pandas fallbacks failed).")
    else:
        # Non-Excel → treat as CSV/TSV/etc.
        try:
            # First try: default CSV reading
            df = pd.read_csv(path, header=None, dtype=str)
            print(f"DEBUG: Successfully read CSV with default engine. Shape: {df.shape}")
            return df
        except Exception as e:
            print(f"DEBUG: Default CSV reading failed: {e}")
            print(f"DEBUG: Error details: {type(e).__name__}: {e}")
            
            # Try to find where the actual CSV data starts by skipping header lines
            print("DEBUG: Trying to skip header lines and find CSV data...")
            for skip_lines in range(0, 10):  # Try skipping 0-9 lines
                try:
                    df = pd.read_csv(path, header=None, dtype=str, skiprows=skip_lines)
                    if df.shape[1] >= 4:  # CSV should have at least 4 columns
                        print(f"DEBUG: Found CSV data starting at line {skip_lines + 1}. Shape: {df.shape}")
                        print(f"DEBUG: First 5 rows after skipping {skip_lines} lines:")
                        print(df.head(5))
                        return df
                except Exception as skip_e:
                    print(f"DEBUG: Skipping {skip_lines} lines failed: {skip_e}")
                    continue
            
            # If still failing, try with Python engine and error handling
            print("DEBUG: Trying Python engine with error handling as final fallback...")
            df = pd.read_csv(path, header=None, dtype=str, engine='python', on_bad_lines='skip')
            print(f"DEBUG: Python engine result. Shape: {df.shape}")
            return df

# In the detect_header_row function:
def detect_header_row(raw: pd.DataFrame) -> int:
    """
    Find the row index that, after normalization + alias mapping,
    contains ALL required canonical columns, but stops after 1000 rows.
    """
    targets = set(REQUIRED_COLS)

    best_count = -1
    best_row = -1
    best_covered = set()

    # Add a max row limit of 1000
    max_row_limit = 1000

    for idx, row in raw.iterrows():
        if idx >= max_row_limit:
            break  # Stop searching after 1000 rows

        # Skip rows that are all empty or all NaN
        if all((pd.isna(x) or str(x).strip() == "") for x in row):
            continue

        # raw cell texts on this row (skip NaN)
        cells = [x for x in row if pd.notna(x)]
        precleaned = [_preclean_header_token(x) for x in cells]
        normed = [_norm(x) for x in precleaned]
        # Exact mapping only (substring disabled)
        normed_keys = [_strip_currency_words_from_key(n) for n in normed]
        mapped = [ALIAS_MAP.get(n) for n in normed_keys]

        covered = {m for m in mapped if m}

        if idx <=10:
            print(f"DEBUG TARGETS: {targets}")
            print("DEBUG IN THE DETECT HEADER ROW: covered so far", covered)
            print(f"DEBUG ROW {idx} CELLS: {cells}")
            print(f"DEBUG ROW {idx} NORMALIZED: {normed}")
            print(f"DEBUG ROW {idx} MAPPED: {mapped}")
            if covered:
                print(f"DEBUG ROW {idx} FOUND MAPPINGS: {[(c, m) for c, m in zip(cells, mapped) if m]}")
            print("---")

        ###################
        # ---------- tolerate missing Billing Increment if related headers exist ----------
        if 'Billing Increment' not in covered:
            norm_set = set(normed)

            def _has_key(key: str) -> bool:
                # fuzzy: exact, prefix, suffix, or underscore-delimited infix
                return any(
                    n == key
                    or n.startswith(key + '_')
                    or n.endswith('_' + key)
                    or ('_' + key + '_') in ('_' + n + '_')
                    for n in norm_set
                )

            # Pairs allow perfect synthesis later
            pair_hit = any(_has_key(a) and _has_key(b) for (a, b) in BILLING_PAIRS)

            # Singles are also enough, because _synthesize_billing_increment can duplicate a single
            # Extract all possible singles from pairs dynamically
            potential_singles = set()
            for a, b in BILLING_PAIRS:
                potential_singles.add(a)
                potential_singles.add(b)
            
            # Add explicit singles
            explicit_singles = ['increment']
            potential_singles.update(explicit_singles)
            
            # Check if any single column from pairs exists (even if its pair doesn't)
            single_hit = any(_has_key(k) for k in potential_singles)

            if pair_hit or single_hit:
                covered.add('Billing Increment')

            dbg("  billing_pair_hit:", pair_hit, "singles_hit:", single_hit,
                "pairs_checked:", BILLING_PAIRS, "singles_checked:", list(potential_singles))
        # ----------------------------------------------------------------------------- 
        ##################

        missing = targets - covered

        # DEBUG dump
        dbg(f"[hdr-scan] row={idx}")
        dbg("  cells:", [repr(x) for x in cells])
        dbg("  precleaned:", [repr(x) for x in precleaned])
        dbg("  normed:", normed)
        dbg("  mapped:", mapped)
        if missing:
            dbg("  still-missing:", sorted(missing))
        else:
            dbg(f"[hdr-scan] FOUND header row -> {idx}")
            return idx

        # track best partial coverage to help when failing
        if len(covered & targets) > best_count:
            best_count = len(covered & targets)
            best_row = idx
            best_covered = covered & targets

    # If we got here, we failed. Print the best candidate with codepoints.
    dbg(f"[hdr-scan] best coverage: {best_count} on row {best_row} -> {sorted(best_covered)}")
    if best_row != -1:
        best_cells = [x for x in raw.iloc[best_row] if pd.notna(x)]
        dbg("[hdr-scan] best row cells (repr):", [repr(x) for x in best_cells])
        dbg("[hdr-scan] best row cells (_norm):", [_norm(x) for x in best_cells])
        dbg("[hdr-scan] best row cells (codepoints):", [_codepoints(str(x)) for x in best_cells])

    # Work out which canonical columns were never satisfied so the error/meta
    # can explicitly list them (e.g. ['Rate', 'Effective Date']). Keep the
    # message short so it surfaces cleanly in processing_statuses.status.
    missing_overall = sorted(targets - best_covered) if best_row != -1 else sorted(targets)

    raise ValueError(
        "Missing required canonical columns: " f"{missing_overall}"
    )


# kill zero-widths and collapse all whitespace/slashes/hyphens; strip punctuation
_ZW = r'[\u200B-\u200D\uFEFF]'

def _norm(s: str) -> str:
    s = unicodedata.normalize('NFKC', str(s)).lower()
    s = re.sub(_ZW, '', s)                 # remove zero-width chars
    s = re.sub(r'[\s/\\\-]+', ' ', s)      # any whitespace, slash, hyphen -> single space
    s = re.sub(r'[^\w ]+', '', s)          # drop punctuation like ., :, (), etc.
    s = ' '.join(s.split())                # collapse to single spaces
    return s.replace(' ', '_')

# Map of normalized header variants -> canonical
ALIAS_MAP = {
    # Dst Code
    'dst_code': 'Dst Code',
    'dstcode': 'Dst Code',
    'code': 'Dst Code',
    'destination_code': 'Dst Code',
    'dest_code': 'Dst Code',
    'area_code': 'Dst Code',
    'dial_code': 'Dst Code',
    'dialcode': 'Dst Code',
    'codes': 'Dst Code',
    'dial_codes': 'Dst Code',
    'area_code': 'Dst Code',
    'prefix': 'Dst Code',
    'dial_code': 'Dst Code',
    'breakout' : 'Dst Code',
    'country_code': 'Dst Code',
    'dialstring': 'Dst Code',
    # Dst Code Name (destination label)
    'dst_code_name': 'Dst Code Name',
    'dstcodename': 'Dst Code Name',
    'code_name': 'Dst Code Name',
    'codename': 'Dst Code Name',
    'destination': 'Dst Code Name',
    'destinations': 'Dst Code Name',
    'destination_name': 'Dst Code Name',
    'dest_name': 'Dst Code Name',
    'dst_name': 'Dst Code Name',
    'dstname': 'Dst Code Name',
    'country_name': 'Dst Code Name',
    # Rate
    'rate': 'Rate',
    'rates': 'Rate',
    'price': 'Rate',
    'new_rates': 'Rate',
    'new_rate': 'Rate',
    'cost': 'Rate',
    'pricing': 'Rate',
    'price': 'Rate',
    'pricing_in': 'Rate',
    'rate_min($)': 'Rate',
    'new_price': 'Rate',
    'price_peak': 'Rate',
    'pricemin': 'Rate',
    'recurring_charge': 'Rate',
    'usd': 'Rate',
    'future_rate': 'Rate',
    'price_min': 'Rate',
    'rate_min': 'Rate',
    'standard_price': 'Rate',
    'allday': 'Rate',
    # Effective Date
    'effective_date': 'Effective Date',
    'effective': 'Effective Date',
    'eff_date': 'Effective Date',
    'effective_from': 'Effective Date',
    'start_date': 'Effective Date',
    'valid_from': 'Effective Date',
    'valid_date': 'Effective Date',
    'date': 'Effective Date',
    'effectivedate': 'Effective Date',
    'efective_date': 'Effective Date',
    'activation_date': 'Effective Date',
    'future_rate_effective_date': 'Effective Date',
    'application_date': 'Effective Date',
    'effective_date_time': 'Effective Date',
    'effective_date_and_time': 'Effective Date',
    'effective_since': 'Effective Date',
    # Billing Increment
    'billing_increment': 'Billing Increment',
    'billing_increament': 'Billing Increment',   # common typo
    'billing_increments': 'Billing Increment',
    'billing_inc': 'Billing Increment',
    'min_inc': 'Billing Increment',  # e.g. "Min./Inc." column
    'billing': 'Billing Increment',
    'billingincrement': 'Billing Increment',
    'rounding_rules': 'Billing Increment',
    'rounding': 'Billing Increment',
    'billing_terms': 'Billing Increment',
    'increment': 'Billing Increment',
    'increments': 'Billing Increment',
    'minimum_increments': 'Billing Increment',
    'bill_incrmnt': 'Billing Increment',
    'initial_recurring': 'Billing Increment',
    'billing_interval': 'Billing Increment',
    'incr_sec': 'Billing Increment',
}


##############################################
# new code for handeling the junk that compes with the col names

# --- Substring alias matching helpers ---
def _normalize_header_key(s: str) -> str:
    """
    Lowercase, replace non-alphanumerics with underscores, collapse repeats,
    and strip edges. Keeps behavior tight and predictable.
    """
    s = unicodedata.normalize('NFKC', str(s)).lower()
    s = re.sub(r'[^0-9a-z]+', '_', s)
    s = re.sub(r'_+', '_', s).strip('_')
    return s

# Build a normalized alias map once. Keys should already be "alias-style"
# but this makes it robust to future edits.
ALIAS_MAP_NORM = { _normalize_header_key(k): v for k, v in ALIAS_MAP.items() }

def _match_alias_substring(normalized_key: str, alias_map: dict = ALIAS_MAP_NORM):
    """
    Substring matching is intentionally disabled.

    We keep the function (with the old implementation commented out) so it’s easy to
    restore later, but the current pipeline uses **exact ALIAS_MAP matches only**.
    """
    return None

    # --- Old substring implementation (commented out) ---
    # if normalized_key in alias_map:
    #     return alias_map[normalized_key]
    #
    # for alias in sorted(alias_map.keys(), key=len, reverse=True):
    #     if alias and alias in normalized_key:
    #         return alias_map[alias]
    # return None


#################################################

def _canonicalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    original = list(df.columns)

    # 1) Preclean labels (kill $, €, USD, etc.)
    preclean_map = {c: _preclean_header_token(c) for c in original}
    # 2) Normalize
    norm_map = {c: _norm(preclean_map[c]) for c in original}
    # 3) Strip currency words that survived normalization (e.g., rate_usd -> rate)
    key_map = {c: _strip_currency_words_from_key(norm_map[c]) for c in original}
    
    # 4) Build mapping from canonical names to potential column matches
    canonical_to_columns = {}
    column_to_canonical = {}
    
    # First pass: collect all potential matches (exact and substring)
    for c in original:
        # Special cases first (explicit destination name columns)
        pc_l = preclean_map[c].lower()
        if pc_l in (
            'dst code name',
            'code name',
            'destination',
            'destinations',
            'destination name',
            'destination country',
        ):
            column_to_canonical[c] = 'Dst Code Name'
            canonical_to_columns.setdefault('Dst Code Name', []).append(c)
            continue
        # Heuristic: map any column that looks like "destination ... name"
        # to Dst Code Name. This is intentionally conservative to avoid ambiguity with Dst Code.
        if ("name" in pc_l) and (("destination" in pc_l) or ("dst" in pc_l)):
            column_to_canonical[c] = 'Dst Code Name'
            canonical_to_columns.setdefault('Dst Code Name', []).append(c)
            continue

        if preclean_map[c].lower() == 'current rate usd':
            column_to_canonical[c] = 'CURRENT RATE USD'
            canonical_to_columns.setdefault('CURRENT RATE USD', []).append(c)
            continue

        key = key_map[c]
        canonical_name = None
        
        # Exact match only (substring disabled)
        if key in ALIAS_MAP:
            canonical_name = ALIAS_MAP[key]
        else:
            canonical_name = None
        
        if canonical_name:
            column_to_canonical[c] = canonical_name
            canonical_to_columns.setdefault(canonical_name, []).append(c)
    
    # Second pass: resolve conflicts.
    #
    # IMPORTANT: For some canonicals (especially 'Dst Code'), multiple matches are ambiguous
    # (e.g., a sheet containing both "Code" and "Prefix"). In these cases we prefer to fail
    # fast with a clear error instead of silently choosing one.
    final_alias_hit = {}
    
    for canonical_name, column_candidates in canonical_to_columns.items():
        if len(column_candidates) == 1:
            # No conflict, use the single match
            final_alias_hit[column_candidates[0]] = canonical_name
        else:
            # Multiple columns map to same canonical name.
            # For Billing Increment we intentionally allow duplicates (e.g., Min/Inc split into two cols)
            # because `_synthesize_billing_increment` will coalesce them into a single "min/inc" string.
            if canonical_name == 'Billing Increment':
                # Only allow when the candidates are genuinely duplicates of the SAME header key,
                # e.g. "BILLING TERMS" repeated twice. If they come from different keys (like
                # "increment" + "billing_increment"), that is ambiguous and should still fail fast.
                cand_keys_set = {key_map.get(cc) for cc in column_candidates}
                if len(cand_keys_set) == 1:
                    for cc in column_candidates:
                        final_alias_hit[cc] = canonical_name
                else:
                    cand_keys = {c: key_map.get(c) for c in column_candidates}
                    raise ValueError(
                        f"Ambiguous columns for '{canonical_name}': {column_candidates}. "
                        f"These all map to '{canonical_name}' after normalization/aliasing. "
                        f"Normalized keys: {cand_keys}. "
                        f"Please keep only ONE column for '{canonical_name}' and remove/rename the others."
                    )
            else:
                # always ambiguous; fail fast.
                cand_keys = {c: key_map.get(c) for c in column_candidates}
                raise ValueError(
                    f"Ambiguous columns for '{canonical_name}': {column_candidates}. "
                    f"These all map to '{canonical_name}' after normalization/aliasing. "
                    f"Normalized keys: {cand_keys}. "
                    f"Please keep only ONE column for '{canonical_name}' and remove/rename the others."
                )
    
    # Ensure all columns have an entry
    for c in original:
        if c not in final_alias_hit:
            final_alias_hit[c] = None

    # DEBUG
    dbg("[canon] original -> preclean -> norm -> key_strip -> alias:")
    for c in original:
        dbg(f"  {repr(c)}  ->  {repr(preclean_map[c])}  ->  {norm_map[c]}  ->  {key_map[c]}  ->  {final_alias_hit[c]}")
        if not final_alias_hit[c]:
            dbg("    codepoints(original):", _codepoints(c))

    # If alias matches, use canonical; otherwise keep the cleaned label
    mapped = {c: (final_alias_hit[c] if final_alias_hit[c] else preclean_map[c]) for c in original}
    df = df.rename(columns=mapped)

    # ---------- NEW: tolerate missing Billing Increment if a known pair exists ----------
    missing = [c for c in REQUIRED_COLS if c not in df.columns]

    if 'Billing Increment' in missing:
        # normalize the CURRENT df.columns (post-rename) for fuzzy matching
        normed_current = {_norm(_preclean_header_token(c)) for c in df.columns}

        def _has_key(key: str) -> bool:
            # fuzzy: exact, prefix, suffix, or underscore-delimited infix
            return any(
                n == key or
                n.startswith(key + '_') or
                n.endswith('_' + key) or
                ('_' + key + '_') in ('_' + n + '_')
                for n in normed_current
            )

        pair_hit = any(_has_key(a) and _has_key(b) for (a, b) in BILLING_PAIRS)
        dbg("[canon] billing_pair_hit:", pair_hit, "pairs_checked:", BILLING_PAIRS)

        dbg("[canon] billing_pair_hit:", pair_hit, "pairs_checked:", BILLING_PAIRS)

        if pair_hit:
            # don’t count BI as missing; create placeholder so later selection won’t crash
            missing = [m for m in missing if m != 'Billing Increment']
            if 'Billing Increment' not in df.columns:
                df['Billing Increment'] = ''   # _synthesize_billing_increment will fill this later
            dbg("[canon] Billing Increment satisfied via header pair; will synthesize values later.")
    # -----------------------------------------------------------------------------------

    # Final guard
    if missing:
        dbg("[canon] df.columns:", list(df.columns))
        dbg("[canon] missing required:", missing)
        dbg("[canon] precleaned originals:", preclean_map)
        dbg("[canon] normalized originals:", norm_map)
        dbg("[canon] stripped keys:", key_map)
        raise ValueError(f"Missing required columns: {missing}.")
    return df


# ──────────────────────────────── footer ─────────────────────────────────

# Customize the keywords if you want to add more
_NOTES_RE = re.compile(r'^(note|notes|remark|remarks|disclaimer|footer|terms and conditions)\b', re.IGNORECASE)

def _is_empty_row(row: pd.Series) -> bool:
    """True if every cell in the row is NaN or only whitespace."""
    for v in row:
        if pd.isna(v):
            continue
        if isinstance(v, str):
            if v.strip() != "":
                return False
        else:
            # non-string, non-NaN value counts as data
            return False
    return True

def trim_after_notes_and_strip_blank_above(df: pd.DataFrame) -> pd.DataFrame:
    """
    - Find the first row where ANY cell starts with one of the notes keywords.
    - Cut the DataFrame to everything ABOVE that row.
    - Then remove trailing blank rows immediately above the notes marker.
    If no notes marker is found, returns df unchanged.
    """
    if df.empty:
        return df

    # Build a normalized view for matching
    # norm = df.applymap(lambda x: "" if pd.isna(x) else str(x).strip())

    try:
        norm = df.map(lambda x: "" if pd.isna(x) else str(x).strip())
    except AttributeError:
        norm = df.applymap(lambda x: "" if pd.isna(x) else str(x).strip())


    # Row-wise: does ANY cell start with a notes-like keyword?
    marker_mask = norm.apply(lambda r: any(_NOTES_RE.match(cell) for cell in r), axis=1)

    if not marker_mask.any():
        # no notes marker -> return as-is
        return df

    # Position of first notes marker (iloc)
    marker_index_label = marker_mask.idxmax()  # first True by index order
    marker_iloc = df.index.get_loc(marker_index_label)

    # Everything strictly ABOVE the marker
    end_iloc = marker_iloc - 1

    # Strip trailing blank rows immediately above the marker
    while end_iloc >= 0 and _is_empty_row(df.iloc[end_iloc]):
        end_iloc -= 1

    # If nothing left, return empty frame with same columns
    if end_iloc < 0:
        return df.iloc[0:0].copy()

    return df.iloc[:end_iloc + 1].copy()

# ───────────────────────────────── helpers ──────────────────────────────────

def _normalize_excel_writer_path(path: str) -> tuple[str, dict]:
    """
    Make writer extension predictable and attach the right engine.
    - Lowercase xlsx/xlsm and use openpyxl.
    - For .xls or unknown junk, upgrade to .xlsx and use openpyxl.
    """
    root, ext = os.path.splitext(path)
    ext_low = ext.lower()
    if ext_low in ('.xlsx', '.xlsm'):
        return root + ext_low, {'engine': 'openpyxl'}
    if ext_low == '.xls':
        new_path = root + '.xlsx'
        dbg(f"[writer] upgrading output extension from {ext} to .xlsx")
        return new_path, {'engine': 'openpyxl'}
    # no or weird extension → force .xlsx
    new_path = (root if ext else path) + '.xlsx'
    dbg(f"[writer] forcing .xlsx for unknown ext {ext or '(none)'}")
    return new_path, {'engine': 'openpyxl'}

def normalise_date_any(val, date_format: str | None = None) -> pd.Timestamp:
    """Return pandas.Timestamp (UTC-naive) or NaT if invalid.
       If date_format is provided, parse strictly to that format; else auto-guess."""
    print(f"\n\n\n Date Format received: {date_format},, {val} \n\n\n")
    if val is None or (isinstance(val, float) and np.isnan(val)) or str(val).strip() == '':
        return pd.NaT

    s = str(val).strip()

    # Excel serial (integer days since 1899-12-30) – allowed in both modes
    if re.fullmatch(r'\d{1,6}', s):
        try:
            serial = int(s)
            if serial > 0:
                return EXCEL_EPOCH + pd.Timedelta(days=serial)
        except Exception:
            return pd.NaT

    # Strip timezone suffixes and unify separators
    s = re.sub(r'\s*\+\d{4}$', '', s).rstrip('Zz').strip()

    if date_format:
        # STRICT mode: parse exactly as given format
        s_norm = s.replace('/', '-').replace('.', '-')
        fmt_raw = date_format.strip().lower().replace('/', '-').replace('.', '-')

        # accept either python strptime tokens or simple tokens like yyyy-mm-dd, mm-dd-yy, etc.
        if '%' in fmt_raw:
            fmt_py = fmt_raw
        else:
            # replace longer tokens first to avoid overlaps
            fmt_py = (fmt_raw
                      .replace('yyyy', '%Y')
                      .replace('yy',   '%y')
                      .replace('mm',   '%m')
                      .replace('dd',   '%d'))
        try:
            dt = datetime.strptime(s_norm, fmt_py)
            print(f"\n\n\n Debug: Date parsed: {dt} \n\n\n")
            return pd.Timestamp(dt.date())
        except Exception:
            return pd.NaT

    # AUTO mode (original behavior)
    s_auto = s.replace('/', '-').replace('.', '-')
    for dayfirst in (True, False):
        try:
            dt = dparse.parse(s_auto, dayfirst=dayfirst, fuzzy=True, default=datetime(1900, 1, 1))
            return pd.Timestamp(dt.date())
        except Exception:
            continue

    return pd.NaT


def clean_billing_increment(val) -> str:
    """
    Normalize vendor increments with the rule:
      - If there are exactly two numbers -> keep as-is (no changes).
      - If there are 3+ numbers and any zeros -> drop all zeros, then take the first two remaining.
      - If only one number -> duplicate it (n/n).
      - If nothing useful remains (all zeros, or empty) -> return ''.

    Examples:
      "0/1/1"   -> "1/1"
      "1/0/60"  -> "1/60"
      "60/60"   -> "60/60"     (unchanged)
      "60"      -> "60/60"
      "0/0/0"   -> ""
    """
    if pd.isna(val):
        return ''
    nums = [int(n) for n in re.findall(r'\d+', str(val))]

    if not nums:
        return ''

    if len(nums) == 2:
        # leave exactly-two-value cases untouched (per your spec)
        return f"{nums[0]}/{nums[1]}"

    if len(nums) >= 3:
        # remove all zeros, then keep the first two remaining
        nz = [n for n in nums if n != 0]
        if len(nz) >= 2:
            return f"{nz[0]}/{nz[1]}"
        if len(nz) == 1:
            return f"{nz[0]}/{nz[0]}"
        return ''  # all zeros like "0/0/0"

    # len(nums) == 1
    return f"{nums[0]}/{nums[0]}"


# def normalize_dates(df, column_name, date_format_email):
#     """
#     Normalize the date column to the required format and strip the time based on the date_format_email.
#     """

#     # Remove any time-related part after the date using regex
#     print(f"\n\n\nBefore extracting date part:\n{df[column_name]}\n\n\n")

#     # Replace . or / with -
#     df[column_name] = df[column_name].str.replace(r'[./]', '-', regex=True)

#     print(f"\n\n\nAfter replacing . or / with -:\n{df[column_name]}\n\n\n")

#     df[column_name] = df[column_name].astype(str).str.extract(r'(\d{1,4}[-./]\d{1,2}[-./]\d{1,4})')[0]
   

#     print(f"\n\n\nAfter extracting date part:\n{df[column_name]}\n\n\n")

#     if date_format_email == 'MM-DD-YYYY':
#         # Handle MM-DD-YYYY format, convert to YYYY-MM-DD
#         df[column_name] = pd.to_datetime(df[column_name], format='%m-%d-%Y', errors='coerce')
#         df[column_name] = df[column_name].dt.strftime('%Y-%m-%d')

#     elif date_format_email == 'DD-MM-YYYY':
#         # Handle DD-MM-YYYY format, convert to YYYY-MM-DD
#         df[column_name] = pd.to_datetime(df[column_name], format='%d-%m-%Y', errors='coerce')
#         df[column_name] = df[column_name].dt.strftime('%Y-%m-%d')

#     elif date_format_email == 'YYYY-MM-DD':
#         # Handle YYYY-MM-DD format, no conversion needed
#         df[column_name] = pd.to_datetime(df[column_name], format='%Y-%m-%d', errors='coerce')
#         df[column_name] = df[column_name].dt.strftime('%Y-%m-%d')

#     else:
#         # If the date format doesn't match any of the three options, raise an exception
#         raise ValueError(f"Unsupported date format: {date_format_email}")

#     return df



def normalize_dates(df: pd.DataFrame, column_name: str, date_format_email: str | None) -> pd.DataFrame:
    """
    Normalize to 'YYYY-MM-DD'. Supports:
      - Numeric: MM-DD-YYYY, DD-MM-YYYY, YYYY-MM-DD
      - Month names: DD-MMM-YYYY, MMM-DD-YYYY, DD-MMMM-YYYY, MMMM-DD-YYYY, YYYY-MMM-DD, YYYY-MMMM-DD
      - Strips time/zone tails and handles Excel serial dates.
      - If date_format_email is None or 'AUTO', it tries multiple formats.
    """
    s = df[column_name].astype(str).str.strip()

    # Excel serials first (days since 1899-12-30)
    def _excel_serial_to_date(text: str):
        if re.fullmatch(r'\d{1,6}', text or ''):
            try:
                serial = int(text)
                if serial > 0:
                    return (EXCEL_EPOCH + pd.Timedelta(days=serial)).date()
            except Exception:
                pass
        return None

    # Clean: drop tz tokens, unify separators, remove trailing time (optionally with offset)
    s = s.str.replace(r'(?:\s+(?:UTC|Z|zz)|\s+[+\-]\d{2}:?\d{2}|\s+[+\-]\d{4})\s*$', '', regex=True)
    s = s.str.replace(r'[./]', '-', regex=True)
    # Also handle cases like "2026-02-09 00:00:00+0000" and "2026-02-04T10:30:00Z"
    # by allowing an optional +HHMM/+HH:MM or trailing Z after the time
    s = s.str.replace(r'[T ]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?(?:Z|[+\-]\d{2}:?\d{2}|[+\-]\d{4})?$', '', regex=True)

    def _to_strptime(fmt: str) -> str:
        f = fmt.strip().lower().replace('/', '-').replace('.', '-')
        return (f.replace('yyyy', '%Y')
                 .replace('yy',   '%y')
                 .replace('mmmm', '%B')
                 .replace('mmm',  '%b')
                 .replace('mm',   '%m')
                 .replace('dd',   '%d'))

    explicit = {
        'MM-DD-YYYY','DD-MM-YYYY','YYYY-MM-DD',
        'DD-MMM-YYYY','MMM-DD-YYYY','DD-MMMM-YYYY','MMMM-DD-YYYY',
        'YYYY-MMM-DD','YYYY-MMMM-DD'
    }
    auto_formats = [
        '%Y-%m-%d','%d-%m-%Y','%m-%d-%Y',
        '%d-%b-%Y','%b-%d-%Y','%Y-%b-%d',
        '%d-%B-%Y','%B-%d-%Y','%Y-%B-%d',
    ]

    out = []
    # Track failures when an explicit date_format_email is provided so we
    # can raise instead of silently auto-parsing to a wrong day/month.
    explicit_parse_failures = []

    for raw in s.tolist():
        if not raw:
            out.append(None); continue

        ser = _excel_serial_to_date(raw)
        if ser is not None:
            out.append(ser.strftime('%Y-%m-%d')); continue

        parsed = None
        req = (date_format_email or '').strip()

        # helper: additional fallbacks to try if strict parse fails
        fallback_formats = [
            '%Y-%m-%d','%d-%m-%Y','%m-%d-%Y',
            '%d-%b-%Y','%b-%d-%Y','%Y-%b-%d',
            '%d-%B-%Y','%B-%d-%Y','%Y-%B-%d',
        ]

        if not req or req.upper() == 'AUTO':
            for fmt in fallback_formats:
                try:
                    parsed = datetime.strptime(raw, fmt); break
                except ValueError:
                    pass
            if parsed is None:
                try:
                    parsed = dparse.parse(raw, dayfirst=True, fuzzy=True, default=datetime(1900,1,1))
                except Exception:
                    parsed = None
        else:
            # EXPLICIT FORMAT MODE: when user passes date_format_email we
            # must not auto-guess other layouts, because that can swap
            # month/day. Try only the requested format; if it fails, mark
            # this value as an error and leave it unparsed.
            fmt = _to_strptime(req) if '%' not in req else req
            try:
                parsed = datetime.strptime(raw, fmt)
            except ValueError:
                parsed = None

        if req and req.upper() != 'AUTO':
            # Expand the requested format into a small set of equivalent patterns
            # (same logical order) before giving up. This keeps strictness while
            # allowing common separators and month-name variants.
            def _equivalent_formats(template: str) -> list[str]:
                key = template.strip().lower().replace(',', '')
                key = key.replace('/', '-').replace('.', '-')
                # If caller already passed a strptime pattern, just return it.
                if '%' in key:
                    return [template]

                def _with_month_first() -> list[str]:
                    return [
                        # 4-digit year
                        '%m-%d-%Y', '%m/%d/%Y', '%m.%d.%Y',
                        '%b-%d-%Y', '%B-%d-%Y',
                        '%b %d %Y', '%B %d %Y',
                        '%b %d, %Y', '%B %d, %Y',
                        '%b/%d/%Y', '%B/%d/%Y',
                        # 2-digit year variants for flexibility
                        '%m-%d-%y', '%m/%d/%y', '%m.%d.%y',
                        '%b-%d-%y', '%B-%d-%y',
                        '%b %d %y', '%B %d %y',
                        '%b %d, %y', '%B %d, %y',
                        '%b/%d/%y', '%B/%d/%y',
                    ]

                def _with_day_first() -> list[str]:
                    return [
                        # 4-digit year
                        '%d-%m-%Y', '%d/%m/%Y', '%d.%m.%Y',
                        '%d-%b-%Y', '%d-%B-%Y',
                        '%d %b %Y', '%d %B %Y',
                        '%d %b, %Y', '%d %B, %Y',
                        '%d/%b/%Y', '%d/%B/%Y',
                        # 2-digit year variants for flexibility
                        '%d-%m-%y', '%d/%m/%y', '%d.%m.%y',
                        '%d-%b-%y', '%d-%B-%y',
                        '%d %b %y', '%d %B %y',
                        '%d %b, %y', '%d %B, %y',
                        '%d/%b/%y', '%d/%B/%y',
                    ]

                def _with_year_first() -> list[str]:
                    return [
                        # 4-digit year
                        '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d',
                        '%Y-%b-%d', '%Y-%B-%d',
                        '%Y %b %d', '%Y %B %d',
                        # 2-digit year at start (yy-mm-dd, yy-bb-dd)
                        '%y-%m-%d', '%y/%m/%d', '%y.%m.%d',
                        '%y-%b-%d', '%y-%B-%d',
                        '%y %b %d', '%y %B %d',
                    ]

                if key == 'mm-dd-yyyy':
                    return _with_month_first()
                if key == 'dd-mm-yyyy':
                    return _with_day_first()
                if key == 'yyyy-mm-dd':
                    return _with_year_first()
                # Fallback: just the normalized strptime translation
                return [_to_strptime(template)]

            if raw and parsed is None:
                for alt in _equivalent_formats(req):
                    try:
                        parsed = datetime.strptime(raw, alt)
                        break
                    except ValueError:
                        continue

            if raw and parsed is None:
                explicit_parse_failures.append(raw)

        out.append(parsed.strftime('%Y-%m-%d') if parsed else None)

    df[column_name] = out
    # If a specific format was requested and any non-empty value could not
    # be parsed, raise instead of silently returning wrong/guessed dates.
    if date_format_email and date_format_email.upper() != 'AUTO' and explicit_parse_failures:
        unique_bad = sorted(set(explicit_parse_failures))
        raise ValueError(
            f"Failed to parse Effective Date values with format '{date_format_email}': "
            f"{unique_bad[:5]}"
        )
    return df



def _normalize_excel_writer_path(path: str) -> tuple[str, dict]:
    """
    Make writer extension predictable and attach the right engine.
    - Keep .xlsx/.xlsm and use openpyxl.
    - If .xls or unknown/blank extension, switch to .xlsx (openpyxl can’t write .xls).
    """
    root, ext = os.path.splitext(path)
    ext_low = ext.lower()
    if ext_low in ('.xlsx', '.xlsm'):
        return root + ext_low, {'engine': 'openpyxl'}
    if ext_low == '.xls':
        new_path = root + '.xlsx'
        print(f"[writer] upgrading output extension from {ext} to .xlsx")
        return new_path, {'engine': 'openpyxl'}
    # no or weird extension → force .xlsx
    new_path = (root if ext else path) + '.xlsx'
    print(f"[writer] forcing .xlsx for unknown ext {ext or '(none)'}")
    return new_path, {'engine': 'openpyxl'}

def load_clean_rates(path: str, output_path: str, sheet=None, date_format_email: str | None = None) -> pd.DataFrame:
    """
    Robust loader:
      1) Read raw grid (openpyxl for Excel; pandas for CSV/TXT)
      2) Detect header row (using your detect_header_row)
      3) Build a proper DataFrame with headers
      4) Canonicalize headers, trim notes/footer, select required cols
      5) Clean fields and write to Excel
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f'File not found: {path}')

    # 1) raw grid
    raw = _read_raw_matrix(path, sheet=sheet)

    print(f"\n\n\nInitial raw data read from file:\n{raw}\n\n\n")
    # return raw

    print('\n\nDEBUG: Calling the detect header row function from load_clean_rates\n\n')
    # 2) detect header row in the raw grid
    header_row_idx = detect_header_row(raw)

    # 3) construct DF: header = that row; data = rows below it
    header_values = list(raw.iloc[header_row_idx].fillna('').astype(str))
    header_values = _make_unique_header_names(header_values)
    df = raw.iloc[header_row_idx+1:].copy()
    df.columns = header_values

    # drop columns/rows that are completely empty after slicing
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # 4) canonicalize & trim
    df = _canonicalize_headers(df)
    df = _synthesize_billing_increment(df)
    df = trim_after_notes_and_strip_blank_above(df)

    required_cols = list(REQUIRED_COLS)
    # Preserve Dst Code Name when present for vendor files too (needed for New rows UI/DB/Jera upload).
    if "Dst Code Name" in df.columns and "Dst Code Name" not in required_cols:
        required_cols = required_cols + ["Dst Code Name"]

    df = df[required_cols].copy()

    # Early prune: drop any row where a required field is missing/blank.
    # This strips footer/notes rows before further parsing/normalization.
    early_missing = pd.Series(False, index=df.index)
    for col in required_cols:
        if col not in df.columns:
            continue
        s = df[col]
        col_missing = s.isna()
        try:
            t = s.astype(str).str.strip().str.lower()
            col_missing = col_missing | t.eq("") | t.eq("nan") | t.eq("none") | t.eq("nat")
        except Exception:
            pass
        early_missing = early_missing | col_missing

    if early_missing.any():
        df = df.loc[~early_missing].copy()
        df.reset_index(drop=True, inplace=True)

    # 5) clean fields
    # Dst Code: keep as string, strip; drop truly empty codes
    df['Dst Code'] = df['Dst Code'].astype(str).str.strip()
    df = df[df['Dst Code'].ne("")]

    # Rate: strip symbols/spaces and coerce to float
    s = df['Rate'].astype(str).str.strip()
    # replace comma with 
    s = s.str.replace(',', '.', regex=False)
    s = (s
         .str.replace(r'[\$\£\€]', '', regex=True)
         .str.replace(r'\s+', '', regex=True))

 
    df['Rate'] = pd.to_numeric(s, errors='coerce')

    # Billing Increment: normalize to "x/y"
    df['Billing Increment'] = df['Billing Increment'].astype(str).str.strip().apply(clean_billing_increment)

    # Effective Date: robust parse (your helper returns Timestamp or NaT)
    # df['Effective Date'] = df['Effective Date'].apply(normalise_date_any)


    print(f"\n\n\nBefore normalizing Effective Date normalize_dates:\n{df['Effective Date']}\n\n\n")

    # If date_format_email is provided, normalize the dates
    if date_format_email:
        df = normalize_dates(df, 'Effective Date', date_format_email= date_format_email)


    # df['Effective Date'] = df['Effective Date'].apply(lambda v: normalise_date_any(v, date_format=date_format_email))

    print(f"\n\n\nAfter normalizing Effective Date:\n{df['Effective Date']}\n\n\n")

    df = expand_dst_code_rows(df)

    # Strip any non-digits left in Dst Code (kills '-', ';', spaces, exotic dashes, etc.)
    df['Dst Code'] = (
        df['Dst Code']
        .astype(str)
        .str.strip()
        .str.replace(r'\D+', '', regex=True)  # keep only 0-9
    )

    # Drop rows that are incomplete: if ANY required field is empty/NaN/NaT.
    # (User request: remove rows that are fully empty OR have even one empty value.)
    missing_any = pd.Series(False, index=df.index)
    for col in REQUIRED_COLS:
        if col not in df.columns:
            continue
        s = df[col]
        col_missing = s.isna()
        # treat blank strings and common stringified nulls as missing too
        try:
            t = s.astype(str).str.strip().str.lower()
            col_missing = col_missing | t.eq("") | t.eq("nan") | t.eq("none") | t.eq("nat")
        except Exception:
            pass
        missing_any = missing_any | col_missing

    if missing_any.any():
        df = df.loc[~missing_any].copy()
        df.reset_index(drop=True, inplace=True)

    # Guard against "false cleaning" – if almost everything was stripped out,
    # treat this as a failure instead of silently writing a tiny file.
    row_count = len(df)
    if row_count <= 3:
        raise ValueError(
            f"Cleaning produced only {row_count} data rows after validation; "
            "treating this as a failed/false cleaning run."
        )

    # finally, write the cleaned sheet
    out_path, writer_kwargs = _normalize_excel_writer_path(output_path)
    df.to_excel(out_path, index=False, **writer_kwargs)
    return df
# ──────────────────────────── quick test ─────────────────────────────────────
if __name__ == '__main__':
    PATH = "C:/Users/Tahira Sadaf/Documents/attachments/DGS_CLI_SERVICE_Ratesheet_2026-02-18.xlsx"
    OUT_PATH = "C:/Users/Tahira Sadaf/Documents/attachments/cleaned.xlsx"
    FILE_PATH = PATH
    OUTPUT_FILE_PATH = OUT_PATH 
    cleaned = load_clean_rates(FILE_PATH, OUTPUT_FILE_PATH, 0, date_format_email='YYYY-MM-DD')
   
    print('✅ Cleaned and saved.')
