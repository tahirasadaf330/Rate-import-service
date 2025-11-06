from __future__ import annotations
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime, timezone
import pandas as pd
import traceback  # >>> ADDED

# ---- import your existing functions ----
from jerasoft import export_rates_by_query
from ratesheet_comparision_engine import read_table, compare, write_excel
from preprocess_data import load_clean_rates
from database import (
    insert_rate_upload,
    bulk_insert_rate_upload_details,
    mark_processing_stage,
)
from jerasoft import export_rates_by_query, get_table_id_by_name, fetch_active_current_future_rates, save_rates_to_excel

# Reuse constants from your code
ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}
EXPECTED_COLS = ["Code", "Old Rate", "New Rate", "Effective Date", "Status", "Change Type", "Notes"]

def load_metadata(folder: Path) -> Optional[Dict[str, Any]]:
    meta = folder / "metadata.json"
    if not meta.exists():
        return None
    try:
        with meta.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def save_metadata(folder: Path, data: Dict[str, Any]) -> None:
    # atomic write to avoid corrupting metadata.json
    path = folder / "metadata.json"
    import tempfile, os
    with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8", dir=folder) as tmp:
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.flush(); os.fsync(tmp.fileno())
        tmpname = tmp.name
    os.replace(tmpname, path)

def cleaned_out_path(p: Path) -> Path:
    """Return <stem>_cleaned<suffix> in the same folder."""
    p = Path(p)
    return p.with_name(f"{p.stem}_cleaned{p.suffix}")

def find_jerasoft_file(folder: Path) -> Optional[Path]:
    """Prefer jerasoft_comparison_all.xlsx, else first *_jerasoft_comparison.xlsx."""
    prime = folder / "jerasoft_comparison_all_cleaned.xlsx"
    if prime.exists():
        return prime
    candidates = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xls")
        and p.name.lower().endswith("_jerasoft_comparison_cleaned.xlsx")
    )
    return candidates[0] if candidates else None

def vendor_files(folder: Path) -> list[Path]:
    """
    Return vendor files to compare, preferring *_cleaned.* when both exist.
    Excludes metadata.json and any JeraSoft comparison outputs (raw or cleaned).
    """
    def is_jerasoft(p: Path) -> bool:
        n = p.name.lower()
        return (
            n == "jerasoft_comparison_all.xlsx"
            or n == "jerasoft_comparison_all_cleaned.xlsx"
            or n.endswith("_jerasoft_comparison.xlsx")
            or n.endswith("_jerasoft_comparison_cleaned.xlsx")
        )

    # collect candidates
    candidates: list[Path] = []
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.name.lower() == "metadata.json":
            continue
        if f.suffix.lower() not in ALLOWED_EXTS:
            continue
        if is_jerasoft(f):
            continue
        # if not f.stem.lower().endswith("_cleaned"):
        #     continue  # ❌ skip non-cleaned vendor files
        candidates.append(f)

    # prefer *_cleaned over raw twin
    by_base: dict[str, dict[str, Path]] = {}
    for f in candidates:
        stem = f.stem
        is_cleaned = stem.endswith("_cleaned")
        base_stem = stem[:-8] if is_cleaned else stem  # strip "_cleaned"
        key = f"{base_stem}{f.suffix.lower()}"         # base name + ext

        entry = by_base.setdefault(key, {})
        if is_cleaned:
            entry["cleaned"] = f
        else:
            entry["raw"] = f

    chosen: list[Path] = []
    for key, pair in by_base.items():
        chosen.append(pair.get("cleaned") or pair.get("raw"))

    return sorted(chosen)

def as_of_from_metadata(folder: Path) -> str:
    """Use metadata.date_utc if available, else today (UTC, YYYY-MM-DD)."""
    meta = folder / "metadata.json"
    if meta.exists():
        try:
            with meta.open("r", encoding="utf-8") as f:
                data = json.load(f)
            d = (data.get("date_utc") or "").strip()
            if len(d) == 10 and d[4] == "-" and d[7] == "-":
                return d
        except Exception:
            pass
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def read_comparison_table(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}")

    # robust column rename
    rename_map: Dict[str, str] = {}
    cols_norm = {c: " ".join(str(c).strip().split()).lower() for c in df.columns}
    for c, n in cols_norm.items():
        if n == "code": rename_map[c] = "Code"
        elif n == "old rate": rename_map[c] = "Old Rate"
        elif n == "new rate": rename_map[c] = "New Rate"
        elif n == "effective date": rename_map[c] = "Effective Date"
        elif n == "status": rename_map[c] = "Status"
        elif n == "change type": rename_map[c] = "Change Type"
        elif n == "notes": rename_map[c] = "Notes"
        elif n == "old billing increment": rename_map[c] = "Old Billing Increment"
        elif n == "new billing increment": rename_map[c] = "New Billing Increment"
        elif n == "code name": rename_map[c] = "Dst Code Name"
        elif n == "dst code name": rename_map[c] = "Dst Code Name"
    df = df.rename(columns=rename_map)

    # ensure required columns exist
    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name}: missing expected columns {missing}. Found: {list(df.columns)}")

    # Clean types but DO NOT drop optional columns
    df["Code"] = df["Code"].astype(str).str.strip()
    df = df[df["Code"].ne("")]  # drop empty codes
    df["Old Rate"] = pd.to_numeric(df["Old Rate"], errors="coerce")
    df["New Rate"] = pd.to_numeric(df["New Rate"], errors="coerce")
    df["Effective Date"] = pd.to_datetime(df["Effective Date"], errors="coerce", utc=True)
    df.dropna(how="all", inplace=True)
    return df

def df_to_detail_dicts(df: pd.DataFrame) -> List[Dict[str, Any]]:
    details: List[Dict[str, Any]] = []

    has_old_bi    = "Old Billing Increment" in df.columns
    has_new_bi    = "New Billing Increment" in df.columns
    has_code_name = "Dst Code Name" in df.columns  # <-- align with rename

    for _, r in df.iterrows():
        eff = r["Effective Date"]
        eff_py = None if pd.isna(eff) else eff.to_pydatetime()  # tz-aware UTC

        item: Dict[str, Any] = {
            "dst_code": None if pd.isna(r["Code"]) else str(r["Code"]).strip(),
            "rate_existing": None if pd.isna(r["Old Rate"]) else float(r["Old Rate"]),
            "rate_new": None if pd.isna(r["New Rate"]) else float(r["New Rate"]),
            "effective_date": eff_py,
            "change_type": None if pd.isna(r["Change Type"]) else str(r["Change Type"]).strip(),
            "status": None if pd.isna(r["Status"]) else str(r["Status"]).strip(),
            "notes": None if pd.isna(r["Notes"]) else str(r["Notes"]).strip(),
        }

        if has_old_bi:
            v = r.get("Old Billing Increment")
            item["old_billing_increment"] = None if pd.isna(v) else str(v).strip()
        if has_new_bi:
            v = r.get("New Billing Increment")
            item["new_billing_increment"] = None if pd.isna(v) else str(v).strip()
        if has_code_name:
            v = r.get("Dst Code Name")
            item["code_name"] = None if pd.isna(v) else str(v).strip()

        details.append(item)

    return details

BOUND = r"(?:(?<=^)|(?<=,))\s*{label}\s*(?:(?=,)|(?=$))"  # comma-boundary regex

def compute_upload_stats(dfs: List[pd.DataFrame]) -> Dict[str, int]:
    if not dfs:
        return {
            "total_rows": 0,
            "new": 0, "increase": 0, "decrease": 0, "unchanged": 0, "closed": 0,
            "backdated_increase": 0, "backdated_decrease": 0,
            "billing_increment_changes": 0,
        }

    df = pd.concat(dfs, ignore_index=True)

    # Membership by Change Type (supports multi-label like "Billing ... Changes,Backdated Increase")
    is_new      = _has_ct(df, "New")
    is_closed   = _has_ct(df, "Closed")
    is_unchanged= _has_ct(df, "Unchanged")

    is_back_inc = _has_ct(df, "Backdated Increase")
    is_back_dec = _has_ct(df, "Backdated Decrease")

    # Normal inc/dec exclude backdated so totals don’t double-count
    is_inc = _has_ct(df, "Increase") & ~is_back_inc
    is_dec = _has_ct(df, "Decrease") & ~is_back_dec

    # Billing increment changes: prefer ground truth from columns if present;
    # otherwise fall back to label membership.
    bic = 0
    obi = df.get("Old Billing Increment")
    nbi = df.get("New Billing Increment")
    if obi is not None and nbi is not None:
        # Compare with nulls treated as equal and types normalized
        o = pd.Series(obi, dtype="string").fillna("")
        n = pd.Series(nbi, dtype="string").fillna("")
        bic = int((o != n).sum())
    else:
        bic = int(_has_ct(df, "Billing Increments Changes").sum())

    return {
        "total_rows": int(len(df)),
        "new":               int(is_new.sum()),
        "increase":          int(is_inc.sum()),
        "decrease":          int(is_dec.sum()),
        "unchanged":         int(is_unchanged.sum()),
        "closed":            int(is_closed.sum()),
        "backdated_increase":int(is_back_inc.sum()),
        "backdated_decrease":int(is_back_dec.sum()),
        "billing_increment_changes": bic,
    }

def parse_received_at(meta: Dict[str, Any]) -> Optional[datetime]:
    raw = meta.get("receivedDateTime_raw")
    if isinstance(raw, str) and raw.strip():
        try:
            s = raw.strip().replace("Z", "+00:00")
            return datetime.fromisoformat(s).astimezone(timezone.utc)
        except Exception:
            pass
    date_s = meta.get("date_utc")
    time_s = meta.get("time_utc")
    if isinstance(date_s, str) and date_s.strip():
        try:
            ts = f"{date_s.strip()}T{(time_s or '00:00:00').strip()}"
            dt = datetime.fromisoformat(ts)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            else:
                dt = dt.astimezone(timezone.utc)
            return dt
        except Exception:
            return None
    return None

# def load_metadata(folder: Path) -> Optional[Dict[str, Any]]:
#     meta = folder / "metadata.json"
#     if not meta.exists():
#         return None
#     try:
#         with meta.open("r", encoding="utf-8") as f:
#             return json.load(f)
#     except Exception:
#         return None
    
#     def save_metadata(folder: Path, data: Dict[str, Any]) -> None:
#     # atomic write to avoid corrupting metadata.json
#      path = folder / "metadata.json"
#     import tempfile, os
#     with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8", dir=folder) as tmp:
#         json.dump(data, tmp, ensure_ascii=False, indent=2)
#         tmp.flush(); os.fsync(tmp.fileno())
#         tmpname = tmp.name
#     os.replace(tmpname, path)
    
def _has_ct(df: pd.DataFrame, label: str) -> pd.Series:
    pat = re.compile(BOUND.format(label=re.escape(label)), flags=re.IGNORECASE)
    ct = df.get("Change Type", pd.Series([], dtype="object")).astype(str)
    return ct.str.contains(pat, na=False)


# ----------------- per-folder pipeline worker -----------------

def process_one_folder(folder: Path) -> str:
    """
    Run the full pipeline for a single attachments/<folder>.
    Returns a short log string.
    """
    meta_path = folder / "metadata.json"
    meta = load_metadata(folder)
    if not meta:
        return f"[{folder.name}] skip: no/invalid metadata.json"

    if not bool(meta.get("date_verification_ingestion_status")):
        return f"[{folder.name}] skip: waiting for date verification approval"

    # -------- 1) JeraSoft export (if needed) --------
    
    if not bool(meta.get("jerasoft_preprocessed")):
        company     = (meta.get("company") or "").strip()
        subject     = (meta.get("subject") or "").strip()
        prefix      = meta.get("prefix")
        dir_path    = meta.get("directory")
        attachments = meta.get("attachments", [])
        force_table = (meta.get("force_jerasoft_table_name") or "").strip()

        if not dir_path or not attachments:
            return f"[{folder.name}] skip: missing directory/attachments info"

        # choose output filename
        if len(attachments) == 1:
            base_name = Path(attachments[0]).stem
            output_file = f"{base_name}_jerasoft_comparison.xlsx"
        else:
            output_file = "jerasoft_comparison_all.xlsx"
        output_path = str(Path(dir_path) / output_file)

        try:
            if force_table:
                # Override path: resolve table by exact name and export directly
                tid = get_table_id_by_name(force_table)
                if not tid:
                    raise RuntimeError(f"JeraSoft table not found: {force_table}")

                js_df    = fetch_active_current_future_rates(table_id=int(tid))
                saved_to = save_rates_to_excel(js_df, output_path)

                rows_js = int(js_df.shape[0]) if hasattr(js_df, "shape") else 0
                meta["best_table_name"] = force_table

            else:
                # Normal path: build a safe non-empty target_query (fallback to subject if needed)
                bad_kw = {"", "a-z", "rates", "rate", "pricing", "update", "standard", "retail"}
                target_query = company if company.lower() not in bad_kw and len(company) >= 3 else subject
                if not target_query:
                    target_query = subject

                info = export_rates_by_query(
                    target_query=target_query,
                    output_path=output_path,
                    subject=subject,
                    prefix_code=prefix,
                )
                if isinstance(info, str):
                    # export_rates_by_query returns a string on error
                    meta["keyword_error"] = info
                    save_metadata(folder, meta)
                    return f"[{folder.name}] export error: {info}"

                # read back the file for row count
                rows_js = 0
                try:
                    ext = Path(output_path).suffix.lower()
                    df_js = pd.read_excel(output_path) if ext in (".xlsx", ".xls") else pd.read_csv(output_path)
                    rows_js = int(df_js.shape[0])
                except Exception:
                    pass

                meta["best_table_name"] = info.get("best_table_name")

            # common metadata after either path
            meta["human_eval_details_jerasoft"] = {"file": Path(output_path).name, "rows": rows_js}
            meta["need_human_eval_jerasoft"] = bool(meta.get("need_human_eval_jerasoft")) or (rows_js < 100)
            meta["jerasoft_preprocessed"] = True
            save_metadata(folder, meta)

            try:
                mark_processing_stage(directory_name=folder.name, stage="jera_fetched")
            except Exception as e:
                print(f"[{folder.name}] stage warn (jera_fetched): {e}")

        except Exception as e:
            meta["keyword_error"] = str(e)
            save_metadata(folder, meta)
            return f"[{folder.name}] ✖ export failed: {e}"

    # -------- 2) Cleaning (if needed) --------
    # we consider "needed" if metadata has no 'preprocessed_results' or it's empty
    meta = load_metadata(folder) or {}
    pre_map: dict = meta.get("preprocessed_results", {}) or {}
    if not pre_map:
        any_files = False
        for file_path in sorted(p for p in folder.iterdir() if p.is_file()):
            if file_path.name.lower() == "metadata.json":
                continue
            if file_path.suffix.lower() not in ALLOWED_EXTS:
                continue
            any_files = True

            in_path = str(file_path)
            out_path = str(cleaned_out_path(file_path))

            date_fmt = (meta.get("date_format_identified") or "").strip() or None
            fname = file_path.name.lower()
            # force ISO for jerasoft outputs
            date_fmt_to_use = "YYYY-MM-DD" if ("jerasoft_comparison_all" in fname or "jerasoft_comparison" in fname) else date_fmt

            try:
                cleaned_df = load_clean_rates(in_path, out_path, 0, date_format_email=date_fmt_to_use)
                raw_name = Path(in_path).name
                clean_name = Path(out_path).name
                pre_map[raw_name] = True
                pre_map[clean_name] = True

                # small-output hint
                try:
                    row_count = int(getattr(cleaned_df, "shape", [0])[0])
                except Exception:
                    row_count = 0
                meta.setdefault("human_eval_details_pre", {})[raw_name] = {"rows": row_count}
                meta["need_human_eval_pre"] = bool(meta.get("need_human_eval_pre")) or (row_count < 100)
            except Exception as e:
                pre_map[file_path.name] = False
                print(f"[{folder.name}] clean fail {file_path.name}: {e}")

        if any_files:
            meta["preprocessed_results"] = pre_map
            save_metadata(folder, meta)

            # --- Check if any JeraSoft or Vendor file is False ---
            # Split flags into JeraSoft and Vendor based on filename
            jera_flags = [v for name, v in pre_map.items() if "jerasoft" in name.lower()]
            vendor_flags = [v for name, v in pre_map.items() if "jerasoft" not in name.lower()]

            # Flag is false if any file in either group is False
            jera_failed = any(not v for v in jera_flags)
            vendor_failed = any(not v for v in vendor_flags)

            # Final status depends on if there is any failure in either group
            final_ok = not (jera_failed or vendor_failed)

            meta["final_ok"] = final_ok
            save_metadata(folder, meta)

            try:
                mark_processing_stage(
                    directory_name=folder.name,
                    stage="file_cleaned",
                    final_status=final_ok
                )
            except Exception as e:
                print(f"[{folder.name}] stage warn (file_cleaned): {e}")

            # Optional: helpful debug logs
            if not final_ok:
                print(
                    f"[{folder.name}] file_cleaned: FALSE — "
                    f"JeraSoft failed={jera_failed} (flags={jera_flags}), "
                    f"Vendor failed={vendor_failed} (flags={vendor_flags})"
                )
            else:
                print(
                    f"[{folder.name}] file_cleaned: TRUE — "
                    f"JeraSoft ok={not jera_failed}, Vendor ok={not vendor_failed}"
                )


    # -------- 3) Comparison (if needed) --------
    meta = load_metadata(folder) or {}
    if "comparision_result" not in (meta.keys()):
        pre_map = meta.get("preprocessed_results", {}) or {}
        left_path = find_jerasoft_file(folder)
        if not left_path:
            meta["comparision_result"] = {"result": "comparison skipped: no baseline file found"}
            save_metadata(folder, meta)
            return f"[{folder.name}] skip compare: no baseline"

        baseline_ok = bool(pre_map.get(left_path.name))
        if not baseline_ok:
            meta["comparision_result"] = {"result": "comparison skipped: comparison file failed preprocessing"}
            save_metadata(folder, meta)
            return f"[{folder.name}] skip compare: baseline not preprocessed"

        vfiles = vendor_files(folder)
        if not vfiles:
            meta["comparision_result"] = {"result": "no vendor files to compare"}
            save_metadata(folder, meta)
            return f"[{folder.name}] no vendor files"

        as_of_date = as_of_from_metadata(folder)
        try:
            left_df = read_table(str(left_path), None)
        except Exception as e:
            meta["comparision_result"] = {"result": f"comparison skipped: failed to read baseline ({e})"}
            save_metadata(folder, meta)
            return f"[{folder.name}] skip compare: read baseline fail"

        comp_result: Dict[str, bool] = {}
        writes = 0
        for v in vfiles:
            vname = v.name
            if not pre_map.get(vname):
                comp_result[vname] = False
                continue
            try:
                right_df = read_table(str(v), None)
                result, stats = compare(left_df, right_df, as_of_date, 7, 0.0001)
                out_path = folder / f"{v.stem}_comparision_result.xlsx"
                write_excel(result, str(out_path))
                print(f"[{folder.name}] wrote result to {out_path}")

                writes += 1
                comp_result[vname] = True

                meta.setdefault("attachment_stats", {})
                meta["attachment_stats"][v.name] = {
                    **stats,
                    "source_attachment": v.name,
                    "result_file": out_path.name,
                    "generated_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
                }
                save_metadata(folder, meta)


            except Exception as e:
                comp_result[vname] = False
                print(f"\n\n\n\n\n[{folder.name}] compare fail {vname}: {e}\n\n\n\n")

        if comp_result:
            success_any = any(comp_result.values())
            meta["comparision_result"] = {"result": "ok" if success_any else "no comparisons succeeded", **comp_result}
        else:
            meta["comparision_result"] = {"result": "no eligible vendor files"}
        meta["processed_at_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        save_metadata(folder, meta)

        try:
            mark_processing_stage(directory_name=folder.name, stage="rate_compared")
        except Exception as e:
            print(f"[{folder.name}] stage warn (rate_compared): {e}")

    # -------- 4) DB push (per-folder) --------
    meta = load_metadata(folder) or {}
    comp = meta.get("comparision_result")
    if not (isinstance(comp, dict) and str(comp.get("result", "")).strip().lower() == "ok"):
        return f"[{folder.name}] skip DB push: result not ok"

    # discover result files
    result_files = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xls", ".csv")
        and p.stem.lower().endswith("_comparision_result")
    )
    if not result_files:
        return f"[{folder.name}] no result files to push"

    rp = meta.get("results_pushed") or {}
    to_push = [f for f in result_files if rp.get(f.name) is not True]
    if not to_push:
        return f"[{folder.name}] already pushed"

    # read DFs (cache) and aggregate stats (only non-empty)
    dfs_to_push: List[pd.DataFrame] = []
    per_file_df: Dict[str, pd.DataFrame] = {}
    for f in to_push:
        try:
            df_tmp = read_comparison_table(f)
            per_file_df[f.name] = df_tmp
            if not df_tmp.empty:
                dfs_to_push.append(df_tmp)
        except Exception as e:
            print(f"[{folder.name}] ⚠ read for stats {f.name}: {e}")
            per_file_df[f.name] = pd.DataFrame()

    if not any((not d.empty) for d in per_file_df.values()):
        for f in to_push:
            rp[f.name] = "empty file no results to push to the database"
        meta["results_pushed"] = rp
        save_metadata(folder, meta)
        return f"[{folder.name}] all empty; no upload"


    stats_totals = compute_upload_stats(dfs_to_push)

    sender = str(meta.get("sender") or "").strip() or None
    subject = (meta.get("subject") or "").strip() or None

 
    received_at = parse_received_at(meta)
    processed_at = None
    try:
        rawp = meta.get("processed_at_utc")
        processed_at = datetime.fromisoformat((rawp or "").replace("Z", "+00:00")) if rawp else None
    except Exception:
        processed_at = None

    upload_id = meta.get("rate_upload_id")
    if not upload_id:
        try:
            upload_id = insert_rate_upload(
                sender_email=sender,
                subject=subject,
                received_at=received_at,
                processed_at=processed_at,
                totals=stats_totals,
            )
            meta["rate_upload_id"] = int(upload_id)
            save_metadata(folder, meta)
        except Exception as e:
            for f in to_push:
                rp[f.name] = False
            meta["results_pushed"] = rp
            save_metadata(folder, meta)
            return f"[{folder.name}] ✖ create upload row: {e}"

    # push each result file
    pushed_any = False
    rp = meta.get("results_pushed") or {}
    for f in to_push:
        try:
            # >>> CHANGED: never use "or" with a DataFrame on the left
            df = per_file_df.get(f.name)  # may be None
            if df is None:                # >>> CHANGED
                df = read_comparison_table(f)  # >>> CHANGED

            # Safe emptiness check
            if df is None or getattr(df, "empty", True):  # >>> CHANGED
                rp[f.name] = "empty file no results to push to the database"
                print(f"[{folder.name}]    ⚠ empty comparison file; nothing to push")  # >>> ADDED
                continue

            try:  # >>> ADDED (optional logging)
                print(f"[{folder.name}]    info: {f.name} df shape = {tuple(df.shape)}")
            except Exception:
                pass

            details = df_to_detail_dicts(df)
            if not details:  # list truthiness is fine
                rp[f.name] = "no rows extracted to push"
                print(f"[{folder.name}]    ⚠ no detail rows extracted; nothing to push")  # >>> ADDED
                continue

            inserted = bulk_insert_rate_upload_details(int(upload_id), details)
            rp[f.name] = True
            pushed_any = pushed_any or (inserted > 0)
            print(f"[{folder.name}]    ✔ inserted {inserted} rows from {f.name}")  # >>> ADDED

        except Exception as e:
            rp[f.name] = False
            print(f"[{folder.name}] ✖ push {f.name}: {e}")
            print(traceback.format_exc())  # >>> ADDED

    meta["results_pushed"] = rp
    save_metadata(folder, meta)

    if pushed_any:
        try:
            mark_processing_stage(directory_name=folder.name, stage="rate_uploaded", final_status=True)
        except Exception as e:
            print(f"[{folder.name}] stage warn (rate_uploaded): {e}")

    return f"[{folder.name}] done (pushed={sum(1 for v in rp.values() if v is True)}/{len(rp)})"


# ----------------- pool runner -----------------

def run_pipeline_mt(attachments_base: str = "attachments", max_workers: int = 4) -> None:
    base = Path(attachments_base)
    if not base.exists():
        print(f"[PIPELINE] attachments base not found: {base}")
        return

    # choose folders to process (those with metadata.json)
    folders = []
    for d in base.iterdir():
        if not d.is_dir():
            continue
        if (d / "metadata.json").exists():
            folders.append(d)

    if not folders:
        print("[PIPELINE] no folders to process.")
        return

    print(f"[PIPELINE] starting multithread pipeline: {len(folders)} folder(s), workers={max_workers}")
    with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="pipe") as ex:
        futs = {ex.submit(process_one_folder, d): d for d in folders}
        for fut in as_completed(futs):
            try:
                msg = fut.result()
            except Exception as e:
                msg = f"[{futs[fut].name}] ✖ pipeline error: {e}\n{traceback.format_exc()}"  # >>> CHANGED
            print(msg)
