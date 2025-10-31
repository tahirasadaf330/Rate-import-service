# _______________ Jerasoft Script _____________

# def process_all_directories(attachments_base="attachments"):
#     """
#     Walk through all directories inside `attachments/`,
#     check metadata.json, update jerasoft_preprocessed flag,
#     and call export_rates_by_query with correct parameters.
#     """
#     for subdir in Path(attachments_base).iterdir():
#         if not subdir.is_dir():
#             continue

#         meta_file = subdir / "metadata.json"
#         if not meta_file.exists():
#             print(f"[SKIP] No metadata.json in {subdir}")
#             continue

#         # Load metadata
#         try:
#             with open(meta_file, "r", encoding="utf-8") as f:
#                 meta = json.load(f)
#         except Exception as e:
#             print(f"[ERROR] Failed to read {meta_file}: {e}")
#             continue

#         # Skip if already jerasoft_preprocessed
#         if meta.get("jerasoft_preprocessed") is True:
#             print(f"[SKIP] Already jerasoft_preprocessed: {subdir}")
#             continue

#         # NEW: require manual date verification approval
#         if not bool(meta.get("date_verification_ingestion_status")):
#             print(f"[SKIP] Awaiting date verification approval: {subdir}")
#             continue

#         # Extract subject
#         company = meta.get("company")
#         subject = meta.get("subject")
#         prefix  = meta.get("prefix")  # may be int or str

#         missing_company = not str(company or "").strip()
#         has_subject     = bool(str(subject or "").strip())
#         has_prefix      = prefix is not None and str(prefix).strip() != ""

#         if missing_company and has_subject and has_prefix:
#             print(f"[WARN] No company found in {meta_file}, skipping...")
#             continue

#         print('DEBUG: Running JeraSoft export for:', company, '| subject:', subject, '| prefix:', prefix)

#         # Extract directory and attachments
#         dir_path = meta.get("directory")
#         attachments = meta.get("attachments", [])

#         if not dir_path or not attachments:
#             print(f"[WARN] Missing directory/attachments info in {meta_file}, skipping...")
#             continue

#         # Decide output filename
#         if len(attachments) == 1:
#             base_name = Path(attachments[0]).stem
#             output_file = f"{base_name}_jerasoft_comparison.xlsx"
#         else:
#             output_file = "jerasoft_comparison_all.xlsx"

#         output_path = str(Path(dir_path) / output_file)

#         info = None
#         try:
#             print("DEBUG: Calling export_rates_by_query...")
#             info = export_rates_by_query(company, output_path, subject, prefix_code=prefix)
#             print(info)

#             if isinstance(info, str):
#                 # Export failed with a message returned by export_rates_by_query
#                 print(f"[ERROR] Export failed for {company}: {info}")
#                 meta["keyword_error"] = info
#                 with open(meta_file, "w", encoding="utf-8") as f:
#                     json.dump(meta, f, indent=2)
#                 print(f"[INFO] Updated metadata with keyword_error: {meta_file}")

#             else:
#                 # Export succeeded; update metadata and also check row count of the saved file
#                 print(f"[INFO] Export succeeded for {company}, updating metadata.")

#                 if info:
#                     meta["best_table_name"] = info.get("best_table_name")

#                 # --- Count rows in the saved JeraSoft file and set human-eval flags ---
#                 rows_js = 0
#                 try:
#                     out_ext = Path(output_path).suffix.lower()
#                     if out_ext in (".xlsx", ".xls"):
#                         df_js = pd.read_excel(output_path)
#                         rows_js = int(df_js.shape[0])
#                     elif out_ext == ".csv":
#                         df_js = pd.read_csv(output_path)
#                         rows_js = int(df_js.shape[0])
#                     else:
#                         print(f"[WARN] Unrecognized JeraSoft output extension: {out_ext}")
#                 except Exception as e2:
#                     print(f"[WARN] Could not read exported JeraSoft file for row count: {e2}")

#                 # Record details and sticky flag
#                 meta["human_eval_details_jerasoft"] = {
#                     "file": Path(output_path).name,
#                     "rows": rows_js,
#                 }
#                 meta["need_human_eval_jerasoft"] = bool(meta.get("need_human_eval_jerasoft")) or (rows_js < 100)
#                 # ---------------------------------------------------------------------

#                 # Only mark true if the export call didn’t blow up
#                 meta["jerasoft_preprocessed"] = True
#                 with open(meta_file, "w", encoding="utf-8") as f:
#                     json.dump(meta, f, indent=2)

#                 try:
#                     # DB column: is_jera_fetched
#                     mark_processing_stage(directory_name=subdir.name, stage="jera_fetched")
#                 except Exception as e:
#                     print(f"[STATUS][WARN] failed to mark jera_fetched for {subdir.name}: {e}")

#                 print(f"[INFO] Updated metadata with jerasoft_preprocessed flag/best_table_name: {meta_file}")
#                 print(f"[SUCCESS] Exported for {company} -> {output_path}")

#         except Exception as e:
#             print(f"[ERROR] Export failed for {company}: {e}")

#         if info is None:
#             print(f"[SKIP] {company} not exported; moving on.")


# def clean_preprocessed_folders(attachments_dir: str | Path):
#     """
#     For each jerasoft_preprocessed folder:
#       - run load_clean_rates(file, file, 0) for every allowed file inside.
#     Updates metadata.json with success/failure for each file.
#     Returns (folders_processed, files_cleaned).
#     """
#     root = Path(attachments_dir).expanduser().resolve()
#     if not root.exists():
#         raise FileNotFoundError(f"Attachments directory not found: {root}")

#     folders_done = 0
#     files_done = 0

#     print("\n=== Cleaning pass over jerasoft_preprocessed folders ===")
#     for folder in iter_preprocessed_dirs_(root):
#         folders_done += 1
#         print(f"\n[FOLDER] {folder}")

#         metadata_path = folder / "metadata.json"
#         try:
#             with metadata_path.open("r", encoding="utf-8") as f:
#                 metadata = json.load(f)
#         except Exception as e:
#             print(f"  ✖ Failed to load metadata: {e}")
#             continue

#         if "preprocessed_results" not in metadata:
#             metadata["preprocessed_results"] = {}
#         any_files = False

#         for file_path in files_to_clean(folder):
#             any_files = True
#             try:
#                 in_path = str(file_path)
#                 out_path = str(cleaned_out_path(file_path))
#                 print(f"  - Cleaning: {file_path}")
#                 print(f"  -> Output: {out_path}")

#                 date_fmt = (metadata.get("date_format_identified") or "").strip() or None
#                 fname = file_path.name.lower()

#                 # Force ISO dates for JeraSoft outputs
#                 if "jerasoft_comparison_all" in fname or "jerasoft_comparison" in fname:
#                     date_fmt_to_use = "YYYY-MM-DD"
#                 else:
#                     date_fmt_to_use = date_fmt

#                 print(f"DEBUG: Date format for this file: {date_fmt_to_use!r}")

#                 cleaned_df = load_clean_rates(in_path, out_path, 0, date_format_email=date_fmt_to_use)

#                 files_done += 1
#                 print(f"    ✔ cleaned -> {file_path}")

#                 raw_name = Path(in_path).name
#                 clean_name = Path(out_path).name
#                 metadata["preprocessed_results"][raw_name] = True
#                 metadata["preprocessed_results"][clean_name] = True

#                 # Optional: capture small-output hint for human eval
#                 try:
#                     row_count = int(getattr(cleaned_df, "shape", [0])[0])
#                 except Exception:
#                     row_count = 0
#                 metadata.setdefault("human_eval_details_pre", {})[raw_name] = {"rows": row_count}
#                 metadata["need_human_eval_pre"] = bool(metadata.get("need_human_eval_pre")) or (row_count < 100)

#             except Exception as e:
#                 print(f"    ✖ failed cleaning {file_path.name}: {e}")
#                 metadata["preprocessed_results"][file_path.name] = False

#         # mark stage if any file cleaned successfully in this folder
#         try:
#             if any(v is True for v in (metadata.get("preprocessed_results") or {}).values()):
#                 mark_processing_stage(directory_name=folder.name, stage="file_cleaned")
#         except Exception as e:
#             print(f"[STATUS][WARN] failed to mark file_cleaned for {folder.name}: {e}")

#         # Save the updated metadata once per folder
#         try:
#             with metadata_path.open("w", encoding="utf-8") as f:
#                 json.dump(metadata, f, ensure_ascii=False, indent=4)
#         except Exception as e:
#             print(f"  ✖ Failed to update metadata: {e}")

#         if not any_files:
#             print("  (no CSV/Excel files found)")

#     print("\n=== Cleaning summary ===")
#     print(f"Folders processed: {folders_done}")
#     print(f"Files cleaned:     {files_done}")
#     return folders_done, files_done


# def vendor_files(folder: Path) -> list[Path]:
#     """
#     Return vendor files to compare, preferring *_cleaned.* when both exist
#     (even if the extensions differ, e.g., test.csv vs test_cleaned.xlsx).
#     Excludes metadata.json, any JeraSoft comparison outputs (raw or cleaned),
#     and any previously generated *_comparision_result.* files.
#     """
#     def is_jerasoft(p: Path) -> bool:
#         n = p.name.lower()
#         return (
#             n == "jerasoft_comparison_all.xlsx"
#             or n == "jerasoft_comparison_all_cleaned.xlsx"
#             or n.endswith("_jerasoft_comparison.xlsx")
#             or n.endswith("_jerasoft_comparison_cleaned.xlsx")
#         )

#     def is_result_file(p: Path) -> bool:
#         return p.stem.lower().endswith("_comparision_result")

#     # collect candidates (non-jerasoft, non-result, allowed exts)
#     candidates: list[Path] = []
#     for f in sorted(folder.iterdir()):
#         if not f.is_file():
#             continue
#         if f.name.lower() == "metadata.json":
#             continue
#         if f.suffix.lower() not in ALLOWED_EXTS:
#             continue
#         if is_jerasoft(f) or is_result_file(f):
#             continue
#         candidates.append(f)

#     # group by base stem ignoring "_cleaned" and ignoring extension
#     # e.g., "test_csv.csv" and "test_csv_cleaned.xlsx" collapse to "test_csv"
#     groups: dict[str, dict[str, Path]] = {}
#     for f in candidates:
#         stem = f.stem
#         cleaned = stem.endswith("_cleaned")
#         base = stem[:-8] if cleaned else stem  # remove "_cleaned" if present

#         entry = groups.setdefault(base, {})
#         if cleaned:
#             entry["cleaned"] = f
#         else:
#             # only keep the first raw we see; cleaned will override anyway
#             entry.setdefault("raw", f)

#     # pick cleaned if available, else raw
#     chosen = [(v.get("cleaned") or v.get("raw")) for v in groups.values()]
#     return sorted([p for p in chosen if p is not None])


# def compare_preprocessed_folders(
#     attachments_dir: str | Path,
#     *,
#     notice_days: int = 7,
#     rate_tol: float = 0.0,
#     sheet_left=None,
#     sheet_right=None,
# ) -> Tuple[int, int]:
#     """
#     Folder-level logic:
#       1) Require 'jerasoft_preprocessed' True and absence of 'comparision_result'.
#       2) Confirm comparison/baseline file exists AND was jerasoft_preprocessed True.
#          - If its jerasoft_preprocessed flag is False (or missing), skip folder and write:
#            comparision_result: {"result": "...skipped because comparison file failed..."}
#       3) Otherwise, for each vendor file:
#          - If vendor file's jerasoft_preprocessed flag is False/missing -> comparision_result[filename] = False
#          - Else run compare; success -> True, failure -> False
#       4) Write comparision_result to metadata once per folder.
#     """
#     root = Path(attachments_dir).expanduser().resolve()
#     if not root.exists():
#         raise FileNotFoundError(f"Attachments directory not found: {root}")

#     folders_done = 0
#     writes = 0

#     print("\n=== Comparison pass over jerasoft_preprocessed folders ===")
#     for folder in iter_preprocessed_dirs(root):
#         folders_done += 1
#         print(f"\n[FOLDER] {folder}")

#         # load metadata once
#         try:
#             meta = _read_metadata(folder)
#         except Exception as e:
#             print(f"  ✖ cannot read metadata.json: {e}")
#             # can't record anything; just skip this folder
#             continue

#         preproc_map: dict = meta.get("preprocessed_results", {}) or {}

#         # 1) find baseline file
#         left_path = find_jerasoft_file(folder)
#         if not left_path:
#             print("  ✖ No JeraSoft baseline found (jerasoft_comparison_all.xlsx or *_jerasoft_comparison.xlsx). Skipping.")
#             # record reason
#             meta["comparision_result"] = {"result": "comparison skipped: no baseline file found"}
#             _write_metadata(folder, meta)
#             continue

#         # 2) check baseline jerasoft_preprocessed flag
#         baseline_name = left_path.name
#         print(f"  - Baseline file: {baseline_name}")
#         baseline_ok = bool(preproc_map.get(baseline_name))
#         if not baseline_ok:
#             print("  ✖ Baseline comparison file exists but was not successfully jerasoft_preprocessed. Skipping folder.")
#             meta["comparision_result"] = {
#                 "result": "comparison skipped: comparison file failed preprocessing"
#             }
#             _write_metadata(folder, meta)
#             continue

#         # 3) gather vendor files
#         vfiles = vendor_files(folder)
#         if not vfiles:
#             print("  (no vendor files to compare)")
#             # still write an empty result so this folder won't be processed again
#             meta["comparision_result"] = {"result": "no vendor files to compare"}
#             _write_metadata(folder, meta)
#             continue

#         as_of_date = as_of_from_metadata(folder)
#         print(f"  AS_OF_DATE: {as_of_date} | NOTICE_DAYS: {notice_days} | RATE_TOL: {rate_tol}")

#         # 4) read baseline table
#         try:
#             left_df = read_table(str(left_path), sheet_left)
#         except Exception as e:
#             print(f"  ✖ Failed reading LEFT (JeraSoft) {left_path.name}: {e}")
#             meta["comparision_result"] = {"result": f"comparison skipped: failed to read baseline ({e})"}
#             _write_metadata(folder, meta)
#             continue

#         # 5) compare each vendor; build comparision_result as requested
#         comp_result: dict[str, bool] = {}
#         for idx, v in enumerate(vfiles):
#             vname = v.name
#             print(f"  - Compare against vendor: {vname}")

#             # gate on vendor jerasoft_preprocessed flag
#             if not preproc_map.get(vname):
#                 print("    ↳ skipped: vendor file not successfully jerasoft_preprocessed")
#                 comp_result[vname] = False
#                 continue

#             try:
            
#                 right_df = read_table(str(v), sheet_right)
#                 result, stats = compare(left_df, right_df, as_of_date, notice_days, rate_tol)

#                 out_path = folder / f"{v.stem}_comparision_result.xlsx"
#                 write_excel(result, str(out_path))
#                 writes += 1
#                 comp_result[vname] = True
#                 print(f"    ✔ wrote {out_path} ({len(result)} rows)")

#                 # ⬇️ persist per-attachment stats into metadata
#                 try:
#                     meta.setdefault("attachment_stats", {})
#                     stats_key = f"{v.name}"  # e.g. VendorA.xlsx_stats
#                     meta["attachment_stats"][stats_key] = {
#                         **stats,
#                         "source_attachment": v.name,                    # the vendor file we compared
#                         "result_file": out_path.name,                   # the Excel we just wrote
#                         "generated_at_utc": datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
#                     }
#                     _write_metadata(folder, meta)
#                 except Exception as e:
#                     print(f"    ⚠ failed to update attachment_stats in metadata: {e}")

#             except Exception as e:
#                 comp_result[vname] = False
#                 print(f"    ✖ failed comparison for {vname}: {e}")

#         # 6) persist comparision_result
#         # if at least one vendor processed, include a simple outcome line
#         # 6) persist comparision_result
#         if comp_result:
#             success_any = any(comp_result.values())
#             if success_any:
#                 meta["comparision_result"] = {"result": "ok", **comp_result}
#             else:
#                 meta["comparision_result"] = {"result": "no comparisons succeeded", **comp_result}
#             meta["processed_at_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
#         else:
#             meta["comparision_result"] = {"result": "no eligible vendor files"}
#             meta["processed_at_utc"] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

#         try:
#             _write_metadata(folder, meta)
#         except Exception as e:
#             print(f"  ⚠ failed updating metadata: {e}")

#         # mark stage once comparison phase finished for this folder
#         try:
#             mark_processing_stage(directory_name=folder.name, stage="rate_compared")
#         except Exception as e:
#             print(f"[STATUS][WARN] failed to mark rate_compared for {folder.name}: {e}")

#     print("\n=== Comparison summary ===")
#     print(f"Folders processed:     {folders_done}")
#     print(f"Comparison files made: {writes}")
#     return folders_done, writes

##############################################

# def push_rejections_from_metadata(attachments_root: str | Path) -> tuple[int, int]:
#     """
#     Scan each attachments/<folder>/metadata.json and, if not already_pushed,
#     insert one row into rejected_emails based on the first applicable condition:
#       1) jerasoft_preprocessed == False  -> 'jerasoft_error'
#       2) need_human_eval_jerasoft == True -> 'need_human_eval_jerasoft'
#       3) any preprocessed_results entry == False -> 'preprocessing_error'
#       4) need_human_eval_pre == True -> 'need_human_eval_preprocessing'

#     After a successful insert, mark metadata['already_pushed'] = True.

#     Returns: (folders_scanned, rows_inserted)
#     """
#     print("\n=== Pushing rejected emails from metadata.json files ===")
#     root = Path(attachments_root).expanduser().resolve()
#     if not root.exists():
#         raise FileNotFoundError(f"Attachments directory not found: {root}")

#     scanned = 0
#     inserted = 0

#     for child in sorted(root.iterdir()):
#         if not child.is_dir():
#             continue
#         meta_path = child / "metadata.json"
#         if not meta_path.exists():
#             continue

#         scanned += 1

#         try:
#             with meta_path.open("r", encoding="utf-8") as f:
#                 meta = json.load(f)
#         except Exception as e:
#             print(f"(warn) cannot read {meta_path}: {e}")
#             continue

#         # skip if already pushed
#         if bool(meta.get("already_pushed")) is True:
#             continue

#         sender = (meta.get("sender") or "").strip() or None
#         subject = (meta.get("subject") or "").strip() or None
#         received_at = _parse_iso_utc_safe(meta.get("receivedDateTime_raw"))
#         processed_at = _parse_iso_utc_safe(meta.get("processed_at_utc"))

#         # Decide category & notes (first match wins)
#         category = None
#         notes = None

#         # 1) JeraSoft export failed/not done
#         jp = meta.get("jerasoft_preprocessed")
#         if jp is False or jp is None:
#             category = "jerasoft_error"
#             ke = meta.get("keyword_error")
#             best = meta.get("best_table_name")
#             bits = ["JeraSoft export did not complete (jerasoft_preprocessed is false/missing)."]
#             if ke:
#                 bits.append(f"keyword_error: {ke}")
#             if best:
#                 bits.append(f"best_table_name (last known): {best}")
#             notes = " ".join(bits)

#         # 2) JeraSoft needs human review: row count < 100
#         if not category and bool(meta.get("need_human_eval_jerasoft")):
#             category = "need_human_eval_jerasoft"
#             det = meta.get("human_eval_details_jerasoft") or {}
#             file_name = det.get("file")
#             rows = det.get("rows")
#             notes = f"JeraSoft export appears too small (<100 rows). File={file_name or 'unknown'}, rows={rows if rows is not None else 'unknown'}."

#         # 3) Preprocessing had failures
#         if not category:
#             pr = meta.get("preprocessed_results") or {}
#             failed = [name for name, ok in pr.items() if ok is False]
#             if failed:
#                 category = "preprocessing_error"
#                 notes = "Preprocessing failed for: " + ", ".join(failed)
        
#         # 3.5) Comparison step failed or skipped (accepts either 'comparision_result' or 'comparison_result')
#         if not category:
#             comp = meta.get("comparision_result") or meta.get("comparison_result")
#             if comp is not None:
#                 if isinstance(comp, dict):
#                     result_val = (comp.get("result") or "").strip()
#                     extra = comp.get("message") or comp.get("detail") or None
#                 else:
#                     result_val = str(comp).strip()
#                     extra = None

#                 # treat anything other than "ok" (case-insensitive) as a failure
#                 if result_val.lower() != "ok":
#                     category = "comparison_failed"
#                     note_parts = [f"Comparison result: {result_val or 'unknown'}."]
#                     if extra:
#                         note_parts.append(str(extra))
#                     notes = " ".join(note_parts)


#         # 4) Vendor files need human review: small outputs
#         if not category and bool(meta.get("need_human_eval_pre")):
#             category = "need_human_eval_preprocessing"
#             det = meta.get("human_eval_details_pre") or {}
#             # det looks like { "fileA.xlsx": {"rows": n}, ... }
#             parts = []
#             for k, v in det.items():
#                 try:
#                     parts.append(f"{k}={int((v or {}).get('rows', 0))} rows")
#                 except Exception:
#                     parts.append(f"{k}=unknown rows")
#             joined = ", ".join(parts) if parts else "no detail"
#             notes = f"Preprocessing produced small outputs (<100 rows): {joined}."
        
#         # 5) Handle 'keyword_error' flag for company name issues
#         if not category and meta.get("keyword_error"):
#             category = "company_name_error"
#             keyword_error_msg = meta.get("keyword_error")
#             notes = f"Company name error: {keyword_error_msg}"

#         # If nothing to report, skip
#         if not category:
#             continue

#         # Insert into DB
#         try:
#             new_id = insert_rejected_email_row(
#                 sender_email=sender,
#                 subject=subject,
#                 category=category,
#                 notes=notes,
#                 received_at=received_at,
#                 processed_at=processed_at,
#             )
#             print(f"(ok) rejected_emails id={new_id} ← {child.name} [{category}]")
#             inserted += 1
#         except Exception as e:
#             print(f"(warn) failed inserting rejected_emails for {child.name}: {e}")
#             continue

#         # Mark as already pushed and persist
#         try:
#             meta["already_pushed"] = True
#             _atomic_write_json(meta_path, meta)
#         except Exception as e:
#             print(f"(warn) failed to mark already_pushed in {meta_path}: {e}")

#     print(f"\n=== Rejections from metadata ===\nFolders scanned: {scanned}\nRows inserted: {inserted}\n")
#     return scanned, inserted


# def push_all_ok_results(attachments_root: str | Path) -> Tuple[int, int, int]:
#     """
#     Idempotent push:
#       - Only create a rate_uploads row if there is at least ONE result file that
#         has not been pushed yet AND has at least one row to insert.
#       - Process only files not previously marked as pushed in metadata.json.
#       - Mark results_pushed[filename] per file with True/False (or a string reason).

#     Returns: (folders_processed, files_processed, rows_inserted_total)
#     """
#     root = Path(attachments_root).expanduser().resolve()
#     if not root.exists():
#         raise FileNotFoundError(f"Attachments directory not found: {root}")

#     folders_done = 0
#     files_done = 0
#     rows_total = 0

#     print("\n=== DB push over OK comparison-result folders ===")
#     for child in sorted(root.iterdir()):
#         if not child.is_dir():
#             continue

#         meta = load_metadata(child)
#         if not meta or not comparison_result_ok(meta):
#             continue

#         # Discover result files in this folder
#         result_files = find_result_files(child)
#         if not result_files:
#             continue

#         # Determine which files still need to be pushed (strict idempotency gate)
#         rp = meta.get("results_pushed") or {}
#         to_push = [f for f in result_files if rp.get(f.name) is not True]
#         if not to_push:
#             # Nothing left to do in this folder
#             continue

#         # Read only the files we plan to push (stats + pre-check for empties)
#         dfs_to_push: List[pd.DataFrame] = []
#         per_file_df: Dict[str, pd.DataFrame] = {}
#         for f in to_push:
#             try:
#                 df_tmp = read_comparison_table(f)
#                 # Remember the DF even if empty so we can mark status later
#                 per_file_df[f.name] = df_tmp
#                 if not df_tmp.empty:
#                     dfs_to_push.append(df_tmp)
#             except Exception as e:
#                 print(f"    ⚠ failed reading {f.name} for stats aggregation: {e}")
#                 per_file_df[f.name] = pd.DataFrame()  # treat as empty so it won't insert

#         # If every to_push DF is empty, don't create a parent record; just mark files
#         if not any((not d.empty) for d in per_file_df.values()):
#             for f in to_push:
#                 mark_results_pushed(child, f.name, "empty file no results to push to the database")
#             # Nothing inserted, but we did meaningful work → count folder processed
#             folders_done += 1
#             print(f"\n[FOLDER] {child}")
#             print("  (All to-push files empty → no upload created)")
#             continue

#         # Aggregate stats from only the DFs that have rows
#         stats_totals = compute_upload_stats(dfs_to_push)

#         # Ready to create the parent upload row now (idempotent: only when there's work)
#         folders_done += 1
#         print(f"\n[FOLDER] {child}")

#         sender = str(meta.get("sender") or "").strip()
#         subject = (meta.get("subject") or "").strip() or None
#         received_at = parse_received_at(meta)
#         processed_at = _parse_iso_utc_safe(meta.get("processed_at_utc"))

#         try:
#             upload_id = insert_rate_upload(
#                 sender_email=sender or None,
#                 subject=subject,
#                 received_at=received_at,
#                 processed_at=processed_at,
#                 totals=stats_totals,
#             )
#         except Exception as e:
#             # Parent failed → mark all to_push files as failed so we don't spin forever
#             print(f"  ✖ Failed to create rate_upload row: {e}")
#             for f in to_push:
#                 mark_results_pushed(child, f.name, False)
#             continue

#         # Process only files that still need pushing
#         for f in to_push:
#             print(f"  - Processing {f.name}")
#             try:
#                 df = per_file_df.get(f.name)
#                 if df is None:
#                     # Safety: read again if it wasn't cached
#                     df = read_comparison_table(f)

#                 if df.empty:
#                     print("    ⚠ empty comparison file; nothing to push")
#                     mark_results_pushed(child, f.name, "empty file no results to push to the database")
#                     continue

#                 details = df_to_detail_dicts(df)

#                 # If you have a UNIQUE constraint on details, turn this into an UPSERT there.
#                 inserted = bulk_insert_rate_upload_details(upload_id, details)
#                 rows_total += inserted
#                 files_done += 1
#                 print(f"    ✔ inserted {inserted} rows")
#                 mark_results_pushed(child, f.name, True)

#             except Exception as e:
#                 print(f"    ✖ failed to insert from {f.name}: {e}")
#                 mark_results_pushed(child, f.name, False)

#     print("\n=== DB push summary ===")
#     print(f"Folders processed: {folders_done}")
#     print(f"Files processed:   {files_done}")
#     print(f"Rows inserted:     {rows_total}")
#     return folders_done, files_done, rows_total
# def push_all_ok_results(attachments_root: str | Path) -> Tuple[int, int, int]:
#     root = Path(attachments_root).expanduser().resolve()
#     if not root.exists():
#         raise FileNotFoundError(f"Attachments directory not found: {root}")

#     folders_done = 0
#     files_done = 0
#     rows_total = 0

#     print("\n=== DB push over OK comparison-result folders ===")
#     for child in sorted(root.iterdir()):
#         if not child.is_dir():
#             continue

#         meta = load_metadata(child)
#         if not meta or not comparison_result_ok(meta):
#             continue

#         result_files = find_result_files(child)
#         if not result_files:
#             continue

#         rp = meta.get("results_pushed") or {}
#         to_push = [f for f in result_files if rp.get(f.name) is not True]
#         if not to_push:
#             continue

#         dfs_to_push: List[pd.DataFrame] = []
#         per_file_df: Dict[str, pd.DataFrame] = {}
#         for f in to_push:
#             try:
#                 df_tmp = read_comparison_table(f)
#                 per_file_df[f.name] = df_tmp
#                 if not df_tmp.empty:
#                     dfs_to_push.append(df_tmp)
#             except Exception as e:
#                 print(f"    ⚠ failed reading {f.name} for stats aggregation: {e}")
#                 per_file_df[f.name] = pd.DataFrame()

#         if not any((not d.empty) for d in per_file_df.values()):
#             for f in to_push:
#                 mark_results_pushed(child, f.name, "empty file no results to push to the database")
#             folders_done += 1
#             print(f"\n[FOLDER] {child}")
#             print("  (All to-push files empty → no upload created)")
#             # Note: do NOT mark rate_uploaded here—nothing was uploaded.
#             continue

#         stats_totals = compute_upload_stats(dfs_to_push)

#         folders_done += 1
#         print(f"\n[FOLDER] {child}")

#         sender = str(meta.get("sender") or "").strip()
#         subject = (meta.get("subject") or "").strip() or None
#         received_at = parse_received_at(meta)
#         processed_at = _parse_iso_utc_safe(meta.get("processed_at_utc"))

#         upload_id = meta.get("rate_upload_id")
#         if upload_id:
#             print(f"  (reusing existing rate_upload_id={upload_id})")
#         else:
#             try:
#                 upload_id = insert_rate_upload(
#                     sender_email=sender or None,
#                     subject=subject,
#                     received_at=received_at,
#                     processed_at=processed_at,
#                     totals=stats_totals,
#                 )
#                 meta["rate_upload_id"] = int(upload_id)
#                 save_metadata(child, meta)
#                 print(f"  ✔ created rate_upload_id={upload_id} and saved to metadata.json")
#             except Exception as e:
#                 print(f"  ✖ Failed to create rate_upload row: {e}")
#                 for f in to_push:
#                     mark_results_pushed(child, f.name, False)
#                 continue

#         # push each file’s rows
#         pushed_any_this_folder = False
#         for f in to_push:
#             print(f"  - Processing {f.name}")
#             try:
#                 df = per_file_df.get(f.name) 
#                 if df is None:
#                     df = read_comparison_table(f)
#                     if df.empty:
#                      print("    ⚠ empty comparison file; nothing to push")
#                     mark_results_pushed(child, f.name, "empty file no results to push to the database")
#                     continue

#                 details = df_to_detail_dicts(df)
#                 inserted = bulk_insert_rate_upload_details(upload_id, details)
#                 rows_total += inserted
#                 files_done += 1
#                 print(f"    ✔ inserted {inserted} rows")
#                 mark_results_pushed(child, f.name, True)
#                 pushed_any_this_folder = True

#             except Exception as e:
#                 print(f"    ✖ failed to insert from {f.name}: {e}")
#                 mark_results_pushed(child, f.name, False)

#         # After file loop, mark rate_uploaded if any file was pushed True
#         if pushed_any_this_folder:
#             try:
#                  mark_processing_stage(
#             directory_name=child.name,
#             stage="rate_uploaded",
#             final_status=True
#         )
#             except Exception as e:
#                 print(f"[STATUS][WARN] failed to mark rate_uploaded for {child.name}: {e}")

#     print("\n=== DB push summary ===")
#     print(f"Folders processed: {folders_done}")
#     print(f"Files processed:   {files_done}")
#     print(f"Rows inserted:     {rows_total}")
#     return folders_done, files_done, rows_total
