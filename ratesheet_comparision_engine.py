import os
from typing import Optional, List
import pandas as pd
import numpy as np

COL_CODE = "Dst Code"
COL_RATE = "Rate"
COL_EDATE = "Effective Date"
COL_BI = "Billing Increment"
COL_NAME = "Dst Code Name"

OUT_COLS = [
    "Code", "Dst Code Name",
    "Old Rate", "New Rate",
    "Old Billing Increment", "New Billing Increment",  
    "Effective Date",
    "Status", "Change Type", "Notes"
]


def read_table(path: str, sheet: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xls"]:
        df = pd.read_excel(path, engine="calamine")
    elif ext in [".xlsx"]:
        df = pd.read_excel(path)
    elif ext in [".csv", ".txt"]:
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    

    #########################
    # Adding other code name col

    expected = [COL_CODE, COL_RATE, COL_EDATE, COL_BI]
    optional = [COL_NAME]  # keep if present

    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns {missing} in {path}. Found: {list(df.columns)}")

    present = expected + [c for c in optional if c in df.columns]
    df = df[present].copy()

    # optional light cleanup
    if COL_NAME in df.columns:
        df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

    #########################



    df[COL_CODE] = df[COL_CODE].apply(lambda x: '' if pd.isna(x) else str(x).strip())
    df[COL_CODE] = df[COL_CODE].str.replace(r'\.0+$', '', regex=True)
    
    df["_rate_raw"] = df[COL_RATE]
    df[COL_RATE] = pd.to_numeric(df[COL_RATE], errors="coerce")
    df["_edate_raw"] = df[COL_EDATE]

    df[COL_EDATE] = (
    pd.to_datetime(df[COL_EDATE], errors="coerce", utc=True)
      .dt.tz_convert(None)     # drop timezone
    #   .dt.normalize()          # <-- force to 00:00:00
    )

    df[COL_BI] = df[COL_BI].astype(str).str.strip()

    df.dropna(how='all', inplace=True)

    return df

def keep_latest_per_code(df: pd.DataFrame) -> pd.DataFrame:
    tmp = df.copy()
    tmp["_edate_for_rank"] = tmp[COL_EDATE].fillna(pd.Timestamp.min)
    idx = tmp.groupby(COL_CODE)["_edate_for_rank"].idxmax()
    return df.loc[idx].copy()

def validate_row(row: pd.Series) -> List[str]:
    reasons = []
    code = row.get(COL_CODE)
    rate = row.get(COL_RATE)
    edate = row.get(COL_EDATE)
    bi = row.get(COL_BI)
    if not (isinstance(code, str) and code.isdigit()):
        reasons.append("invalid code")
    if pd.isna(rate):
        reasons.append("invalid rate format")
    elif not np.isfinite(rate):
        reasons.append("invalid rate format")
    elif rate < 0:
        reasons.append("negative rate")
    if pd.isna(edate):
        reasons.append("invalid effective date")
    if not (isinstance(bi, str) and bi.count('/') == 1):
        reasons.append("invalid billing increment")
    else:
        a, b = bi.split("/")
        if not (a.isdigit() and b.isdigit()):
            reasons.append("invalid billing increment")
    return reasons

def effective_note(new_date: Optional[pd.Timestamp], as_of: pd.Timestamp, notice_days: int) -> str:
    if pd.isna(new_date):
        return "invalid effective date"
    if new_date < as_of:
        return "immediate effective date"
    if new_date >= as_of + pd.Timedelta(days=notice_days):
        return "proper 6-day notice"
    return "new without 6-day notice"

#############################
# Counting stats like number of rows, increase, decrease etc..
############################

CHANGE_TYPES = {
    "new": "New",
    "increase": "Increase",
    "decrease": "Decrease",
    "unchanged": "Unchanged",
    "closed": "Closed",
    "stashed": "Stashed",
    "backdated_increase": "Backdated Increase",
    "backdated_decrease": "Backdated Decrease",
    "billing_increment_changes": "Billing Increments Changes",
}

def summarize_changes(df: pd.DataFrame) -> dict:
    ct = df["Change Type"].astype(str)
    return {
        "total_rows": int(len(df)),
        **{key: int((ct == label).sum()) for key, label in CHANGE_TYPES.items()},
    }

def compare(
    left: pd.DataFrame,
    right: pd.DataFrame,
    as_of_date: Optional[str],
    notice_days: int,
    rate_tol: float,
    *,
    is_partial: bool = False,
    progress_every: int = 0,
) -> pd.DataFrame:
    # Keep originals for stash logic (Jera extract can contain multiple future rates per code).
    left_all = left.copy()
    right_all = right.copy()

    left_dedup = keep_latest_per_code(left_all)
    right_dedup = keep_latest_per_code(right_all)

    merged = left_dedup.merge(right_dedup, on=COL_CODE, how="outer", suffixes=("_old", "_new"), indicator=True)

    left_only = merged["_merge"].eq("left_only")
    right_only = merged["_merge"].eq("right_only")
    both = merged["_merge"].eq("both")
    as_of = pd.to_datetime(as_of_date, errors="coerce")
    if pd.isna(as_of):
        as_of = pd.Timestamp.now().normalize()
    else:
        as_of = as_of.normalize()
    rows = []
    old_rate = merged[f"{COL_RATE}_old"].astype(float)
    new_rate = merged[f"{COL_RATE}_new"].astype(float)
    can_compare_rate = both & old_rate.notna() & new_rate.notna()
    bi_changed = both & (merged[f"{COL_BI}_old"].astype(str).str.strip() != merged[f"{COL_BI}_new"].astype(str).str.strip())

    # -----------------------------
    # Precompute stash candidates (vectorized) to avoid per-row dataframe scans
    # -----------------------------
    stashed_old_row_by_code: dict[str, pd.Series] = {}
    right_min_date_by_code: dict[str, pd.Timestamp] = {}
    try:
        # earliest NEW effective date per code (use all vendor rows)
        right_min_date = (
            right_all[[COL_CODE, COL_EDATE]]
            .dropna(subset=[COL_EDATE])
            .assign(**{COL_CODE: right_all[COL_CODE].astype(str).str.strip()})
            .groupby(COL_CODE, dropna=False)[COL_EDATE]
            .min()
        )
        right_min_date_by_code = {str(k): v for k, v in right_min_date.items()}

        # old rows with effective date later than earliest NEW date => stash candidate,
        # but only if that old effective date is STILL in the future relative to as_of.
        left_future = (
            left_all.dropna(subset=[COL_EDATE])
            .assign(**{COL_CODE: left_all[COL_CODE].astype(str).str.strip()})
            [[COL_CODE, COL_RATE, COL_EDATE, COL_BI] + ([COL_NAME] if COL_NAME in left_all.columns else [])]
        )
        right_min_df = right_min_date.reset_index().rename(columns={COL_EDATE: "n_min"})
        cand = left_future.merge(right_min_df, on=COL_CODE, how="inner")
        # Require: old_date > new_min_date AND old_date >= as_of (future-dated rate)
        cand = cand[(cand[COL_EDATE] > cand["n_min"]) & (cand[COL_EDATE] >= as_of)]
        if not cand.empty:
            cand = cand.sort_values(by=[COL_CODE, COL_EDATE], kind="mergesort")
            cand = cand.groupby(COL_CODE, as_index=False).head(1)
            stashed_old_row_by_code = {str(row[COL_CODE]): row for _, row in cand.iterrows()}
    except Exception:
        stashed_old_row_by_code = {}
        right_min_date_by_code = {}

    print(f"▶ Starting compare: {len(merged)} merged rows (progress_every={progress_every or 'off'})")

    for i, r in merged.iterrows():
        if progress_every and i and (int(i) % int(progress_every) == 0):
            print(f"  ... compared {i}/{len(merged)} merged rows")
        code = r[COL_CODE]
        o_rate = r.get(f"{COL_RATE}_old", np.nan)
        n_rate = r.get(f"{COL_RATE}_new", np.nan)
        o_date = r.get(f"{COL_EDATE}_old", pd.NaT)
        n_date = r.get(f"{COL_EDATE}_new", pd.NaT)
        notes: List[str] = []

        # Stash detection (fast path): use precomputed per-code info
        code_key = str(code).strip() if isinstance(code, str) else str(code)
        n_date_for_stash = right_min_date_by_code.get(code_key, n_date)
        old_future_row = stashed_old_row_by_code.get(code_key)
        stash_case = bool(both[i]) and (old_future_row is not None) and pd.notna(n_date_for_stash)

        #########################
        # ading the code name col
        
        name_new   = r.get(f"{COL_NAME}_new")
        name_old   = r.get(f"{COL_NAME}_old")
        name_plain = r.get(COL_NAME)  # unsuffixed column when only one side had it

        dst_name = (name_new if isinstance(name_new, str) and name_new.strip()
                    else name_old if isinstance(name_old, str) and name_old.strip()
                    else name_plain if isinstance(name_plain, str) and name_plain.strip()
                    else None)

        # Billing Increment (old/new)
        bi_old = r.get(f"{COL_BI}_old")
        bi_new = r.get(f"{COL_BI}_new")

        def _clean_bi(x):
            if isinstance(x, str):
                s = x.strip()
                return s if s else None
            return None

        bi_old = _clean_bi(bi_old)
        bi_new = _clean_bi(bi_new)

        ########################

        # Stashed scenario:
        # If the old rate has a later effective date than the new rate for the same code,
        # AND the rate changed, add an extra row for the old one marked "Stashed".
        # The NEW vendor row will still be classified via the normal Increase/Decrease logic below.
        if stash_case:
            # Override the "old" fields with the nearest-future old row we are going to stash.
            try:
                o_rate = old_future_row.get(COL_RATE, o_rate)
                o_date = old_future_row.get(COL_EDATE, o_date)
                bi_old = _clean_bi(old_future_row.get(COL_BI, bi_old))
                # Prefer code name from the stashed old row if present
                try:
                    ofn = old_future_row.get(COL_NAME)
                    if isinstance(ofn, str) and ofn.strip():
                        dst_name = ofn.strip()
                except Exception:
                    pass
            except Exception:
                pass
            rows.append({
                "Code": code,
                "Dst Code Name": dst_name,
                "Old Rate": o_rate,
                "New Rate": np.nan,
                "Old Billing Increment": bi_old,
                "New Billing Increment": np.nan,
                "Effective Date": o_date,         # the rate that should be stashed
                # Default to Rejected: require manual approval before stashing in JeraSoft.
                "Status": "Rejected",
                "Change Type": "Stashed",
                "Notes": f"old effective {pd.to_datetime(o_date).strftime('%Y-%m-%d')} stashed by new effective {pd.to_datetime(n_date_for_stash).strftime('%Y-%m-%d')}",
            })

        if right_only[i]:
            change_type = "New"
            eff_note = effective_note(n_date, as_of, notice_days)
            status = "Accepted" if eff_note == "proper 6-day notice" else "Rejected"
            notes.append(eff_note)
            reasons = validate_row(pd.Series({COL_CODE: r[COL_CODE], COL_RATE: n_rate, COL_EDATE: n_date, COL_BI: r.get(f"{COL_BI}_new", "")}))
            if reasons:
                status = "Rejected"
                notes.extend(reasons)
            rows.append({"Code": code, "Dst Code Name": dst_name, "Old Rate": o_rate, "New Rate": n_rate, "Old Billing Increment": bi_old, "New Billing Increment": bi_new, "Effective Date": n_date, "Status": status, "Change Type": change_type, "Notes": "; ".join(dict.fromkeys(notes))})
            continue

        if left_only[i]:
            # For partial vendor sheets, missing codes should NOT be treated as "Closed".
            if is_partial:
                continue
            rows.append({"Code": code, "Dst Code Name": dst_name, "Old Rate": o_rate, "New Rate": n_rate,  "Old Billing Increment": bi_old, "New Billing Increment": bi_new, "Effective Date": n_date, "Status": "Rejected", "Change Type": "Closed", "Notes": "present in current system but missing in new (closed)"})
            continue

        left_reasons = validate_row(pd.Series({COL_CODE: r[COL_CODE], COL_RATE: o_rate, COL_EDATE: r.get(f"{COL_EDATE}_old", pd.NaT), COL_BI: r.get(f"{COL_BI}_old", "")}))
        right_reasons = validate_row(pd.Series({COL_CODE: r[COL_CODE], COL_RATE: n_rate, COL_EDATE: n_date, COL_BI: r.get(f"{COL_BI}_new", "")}))
        invalid = bool(left_reasons or right_reasons)
        if bi_changed[i]:

            # Build a combined label. We'll append a rate label with backdated/normal flavor.
            labels = ["Billing Increments Changes"]

            eff_note = effective_note(n_date, as_of, notice_days)

            # BI decides status in this branch, same as your original logic
            if n_date < as_of:
                status = "Rejected"
                notes.append("immediate effective date")
            elif eff_note == "proper 6-day notice":
                status = "Accepted"
                notes.append("proper 6-day notice")
            else:
                status = "Rejected"
                notes.append(eff_note)

            notes.append("billing increment changed")

            # If the rate ALSO changed beyond tolerance (or exact-match mode), append a precise label:
            # Increase/Decrease vs Backdated Increase/Backdated Decrease
            if can_compare_rate[i] and pd.notna(o_rate) and pd.notna(n_rate) and (
                (rate_tol <= 0 and float(n_rate) != float(o_rate))
                or (rate_tol > 0 and not np.isclose(n_rate, o_rate, atol=rate_tol))
            ):
                delta = float(n_rate) - float(o_rate)
                if delta > (rate_tol if rate_tol > 0 else 0.0):
                    labels.append("Backdated Increase" if n_date < as_of else "Increase")
                    notes.append("rate increased")
                elif delta < (-(rate_tol if rate_tol > 0 else 0.0)):
                    labels.append("Backdated Decrease" if n_date < as_of else "Decrease")
                    notes.append("rate decreased")

            # Final combined label, e.g. "Billing Increments Changes,Backdated Increase"
            change_type = ",".join(labels)

############################
###########################
        else:
            if can_compare_rate[i]:
                delta = n_rate - o_rate
                # print(f" → Rate delta={delta}")
                eff_note = effective_note(n_date, as_of, notice_days)
                if delta > (rate_tol if rate_tol > 0 else 0.0):
                    if n_date < as_of:
                        change_type = "Backdated Increase"
                        status = "Rejected"
                        notes.append("immediate effective date")
                    elif eff_note == "proper 6-day notice":
                        change_type = "Increase"
                        status = "Accepted"
                        notes.append("proper 6-day notice")
                    else:
                        change_type = "Increase"
                        status = "Rejected"
                        notes.append(eff_note)
                elif delta < (-(rate_tol if rate_tol > 0 else 0.0)):
                    if n_date < as_of:
                        change_type = "Backdated Decrease"
                        status = "Accepted"
                        notes.append("backdated decrease")
                    else:
                        change_type = "Decrease"
                        status = "Accepted"
                        notes.append("normal decrease")
                else:
                    change_type = "Unchanged"
                    status = "Ignored"
                    notes.append("no change identified")
            else:
                change_type = "Increase"
                status = "Rejected"
                notes.append("cannot determine change due to invalid data")

        if invalid:
            status = "Rejected"
            notes.extend(left_reasons)
            notes.extend(right_reasons)

        # In the stash scenario, always require manual approval for the *replacement/new* row,
        # even when the rate is unchanged. We already emit a separate "Stashed" row for the
        # old future-dated rate.
        if stash_case:
            status = "Rejected"
            notes.append("requires manual approval (replaces future-dated rate)")

        rows.append({"Code": code, "Dst Code Name": dst_name, "Old Rate": o_rate, "New Rate": n_rate, "Old Billing Increment": bi_old, "New Billing Increment": bi_new, "Effective Date": n_date, "Status": status, "Change Type": change_type, "Notes": "; ".join(dict.fromkeys(notes))})

    out = pd.DataFrame(rows, columns=OUT_COLS)
    if not out.empty:
        out = out.sort_values(by=["Code", "Change Type"], kind="mergesort")
    print(f"\n▶ Finished. Produced {len(out)} rows after comparison")
    stats = summarize_changes(out)
    # returning out and stats
    return out, stats 
    
def write_excel(df: pd.DataFrame, path: str) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="comparison")
        ws = writer.sheets["comparison"]
        ws.autofilter(0, 0, max(0, len(df)), max(0, df.shape[1]-1))
        ws.freeze_panes(1, 0)


# --- test runner ---
if __name__ == "__main__":
    # Example usage for testing
    old_file = r"C:\Users\Tahira Sadaf\Documents\attachments\Apelby_-_HAYO_TELECOM_2025-12-22_184155_jerasoft_comparison_cleaned.xlsx"
    new_file = r"C:\Users\Tahira Sadaf\Documents\attachments\Apelby_-_HAYO_TELECOM_2025-12-22_184155_cleaned.xlsx"
    output_file = "comparison_output.xlsx"

    # Read inputs
    left_df = read_table(old_file)
    right_df = read_table(new_file)

    # Compare (tweak params as needed)
    result_df, stats = compare(
        left_df,
        right_df,
        as_of_date="2025-11-06",
        notice_days=6,
        rate_tol=0.0,
    )

    # Show summary
    print("\n=== Summary Stats ===")
    for k, v in stats.items():
        print(f"{k}: {v}")

    # Write output
    write_excel(result_df, output_file)
    print(f"\n✅ Comparison complete. Results saved to: {output_file}")