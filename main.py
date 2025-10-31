""" 
Step 1: run the email verification script to fetch all the new valid files which will be stored in the attachments folder
Step 2: then run the jerasoft script to fetch all the relevant tables for comparision
Step 3: run the preprocess script on all the files 
Step 4: run the ratesheet comparision script to generate the comparision report
Step 5: run the database script to push the comparision results to the database 

"""
"""
while saving the files from email also create a meta data file for that directory that should include subject, sender, date, time
path to the directory.

then read that meta data file and create all the comparision files using jerasoft and save in the same directory.

then preprocess all the files 
"""
from multithreading import run_pipeline_mt
from email_verification import verify_fetch_emails
from pathlib import Path
from database import push_failed_emails_json_to_db, fetch_approved_unprocessed_paths_map
from datetime import date, datetime, timezone
from date_verification import  ingest_files_for_manual_date, mark_date_verification_ingestion
from database_flag import seed_processing_status_rows, finalize_processed_flags
FAILED_EMAILS_PATH = Path(__file__).with_name("failed_emails.json")


#_____________ Email Verification Script_____________

# after = "2025-09-29"              # only include emails on/after this date (YYYY-MM-DD) or None     "2025-08-29"
after = datetime.now().strftime("%Y-%m-%d")
before = None       # only include emails on/before this date (YYYY-MM-DD) or None
unread_only = False    
ATTEMPTS = 2
#____________________________________#
if __name__ == "__main__":
    # scrap all the valid emails
    verify_fetch_emails(after, before, unread_only)
    seed_processing_status_rows("attachments")

    ingest_files_for_manual_date("attachments")

    valid_paths = fetch_approved_unprocessed_paths_map()

    mark_date_verification_ingestion(valid_paths)

    push_failed_emails_json_to_db("failed_emails.json")  
    
    run_pipeline_mt("attachments", max_workers=3)

    finalize_processed_flags(valid_paths)


