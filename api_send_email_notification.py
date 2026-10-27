#karmafy/karmafy/api_send_email_notification.py

from fastapi import FastAPI, HTTPException
from typing import Optional
import base64
import csv
import html
import io
import os
import re
from datetime import datetime
from zoneinfo import ZoneInfo
import psycopg2
import requests
from psycopg2.extras import RealDictCursor
from supabase import create_client, Client

IST = ZoneInfo("Asia/Kolkata")


def today_ist() -> str:
    """Current calendar date in IST (YYYY-MM-DD)."""
    return datetime.now(IST).strftime("%Y-%m-%d")

# -----------------------
# Config - Load from environment variables
# -----------------------
# Load .env file from parent directory
from pathlib import Path
from dotenv import load_dotenv

env_path = Path(__file__).parent.parent / ".env"
load_dotenv(env_path)

# Database configuration - REQUIRED (Source DB)
DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise ValueError(
        "Missing required DATABASE_URL environment variable. "
        "Please set it in your .env file."
    )

# Supabase Configuration - REQUIRED (Destination DB)
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not all([SUPABASE_URL, SUPABASE_KEY]):
    raise ValueError(
        "Missing required Supabase credentials. Please set the following environment variables:\n"
        "  - SUPABASE_URL\n"
        "  - SUPABASE_KEY"
    )

# Azure AD / Microsoft Graph - used to send the daily report
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "support@applywizz.ai")
REPORT_RECIPIENT = "dinesh@applywizz.ai"

APP_NAME = "LinkedIn Job Postings Report"

# FastAPI app
app = FastAPI()


# -----------------------
# Database Connection
# -----------------------
def get_db_connection():
    """Get PostgreSQL database connection for Source DB"""
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        print(f"❌ Database connection error: {e}")
        raise HTTPException(status_code=500, detail="Database connection failed")


# -----------------------
# Query LinkedIn Job Postings
# -----------------------
def get_linkedin_job_postings(target_date: Optional[str] = None):
    """
    Query the karmafy_job table for LinkedIn job postings that have a poster profile
    for the given calendar date (defaults to today's date in IST).
    """
    if target_date is None:
        target_date = today_ist()

    conn = get_db_connection()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            query = """
                SELECT DISTINCT ON (kj.posted_by_profile)
                    kj.company,
                    kj.url,
                    kj.company_url,
                    kj.poster_full_name,
                    kj.posted_by_profile,
                    kj.source,
                    kj.title,
                    kjr.name AS job_role,
                    kj.location,
                    DATE(kj."ingestedAt") AS ingested_date
                FROM public.karmafy_job kj
                LEFT JOIN public.karmafy_jobrole kjr
                    ON kj."roleId"::bigint = kjr.id
                WHERE kj.source = 'LINKEDIN'
                    AND DATE(kj."ingestedAt") = %s
                    AND kj.posted_by_profile IS NOT NULL
                    AND kj.posted_by_profile != ''
                ORDER BY kj.posted_by_profile, kj."ingestedAt" DESC NULLS LAST, kj.company, kj.title
            """
            cursor.execute(query, (target_date,))
            results = cursor.fetchall()

            rows = []
            for row in results:
                item = dict(row)
                ingested = item.get("ingested_date")
                if ingested is not None:
                    item["ingested_date"] = ingested.strftime("%Y-%m-%d")
                rows.append(item)
            return rows
    
    except psycopg2.Error as e:
        print(f"❌ Database query error: {e}")
        raise HTTPException(status_code=500, detail="Database query failed")
    finally:
        conn.close()


# -----------------------
# Report records + Supabase insert
# -----------------------
REPORT_COLUMNS = [
    ("company", "Company"),
    ("job_title", "Job title"),
    ("job_role", "Job role"),
    ("poster_full_name", "Poster"),
    ("poster_profile_url", "Poster profile"),
    ("job_url", "Job URL"),
    ("company_url", "Company URL"),
    ("source", "Source"),
    ("report_date", "Report date"),
]


def build_report_records(jobs_data: list, target_date: str) -> list:
    """Map source DB rows into the report / Supabase payload shape."""
    records = []
    for job in jobs_data:
        records.append({
            "company": job.get("company"),
            "job_title": job.get("title"),
            "job_role": job.get("job_role"),
            "poster_full_name": job.get("poster_full_name"),
            "poster_profile_url": job.get("posted_by_profile"),
            "job_url": job.get("url"),
            "company_url": job.get("company_url"),
            "source": job.get("source"),
            "report_date": job.get("ingested_date") or target_date,
        })
    return records


def insert_jobs_to_supabase(records_to_insert: list) -> bool:
    """
    Insert job postings into Supabase.
    Returns True on success. On failure logs the error and returns False
    (does not raise — email should still be sent).
    """
    if not records_to_insert:
        print("ℹ️ No records to insert.")
        return False

    try:
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        # Batch to avoid oversized payloads
        batch_size = 500
        for i in range(0, len(records_to_insert), batch_size):
            batch = records_to_insert[i : i + batch_size]
            supabase.table("daily_linkedin_jobs_report").insert(batch).execute()
            print(f"✅ Inserted Supabase batch {i // batch_size + 1} ({len(batch)} rows)")
        print(f"✅ Successfully inserted {len(records_to_insert)} records into Supabase.")
        return True
    except Exception as e:
        print(f"⚠️ Supabase insertion failed (continuing to email): {e}")
        return False


# -----------------------
# Email the daily report via Microsoft Graph
# -----------------------
def _get_graph_access_token() -> str:
    """Acquire an app-only Microsoft Graph token with the Azure AD client credentials."""
    if not all([AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
        raise ValueError(
            "Missing Azure credentials. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET."
        )

    token_url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    response = requests.post(
        token_url,
        data={
            "client_id": AZURE_CLIENT_ID,
            "client_secret": AZURE_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    response.raise_for_status()
    token = response.json().get("access_token")
    if not token:
        raise RuntimeError("Azure token response did not include an access_token.")
    return token


def _safe_filename_date(target_date: str) -> str:
    return re.sub(r"[^0-9A-Za-z._-]+", "_", target_date).strip("_") or "report"


def _report_html(records: list, target_date: str, preview_limit: int = 50) -> str:
    header = "".join(f"<th>{html.escape(label)}</th>" for _, label in REPORT_COLUMNS)
    preview = records[:preview_limit]
    rows = []
    for record in preview:
        cells = "".join(
            f"<td>{html.escape('' if record.get(key) is None else str(record.get(key)))}</td>"
            for key, _ in REPORT_COLUMNS
        )
        rows.append(f"<tr>{cells}</tr>")

    extra = ""
    if len(records) > preview_limit:
        extra = (
            f"<p>Showing first {preview_limit} of {len(records)} rows. "
            "Full data is in the CSV attachment.</p>"
        )

    return f"""
    <p>LinkedIn job postings for <strong>{html.escape(target_date)}</strong>.</p>
    <p>{len(records)} posting(s) in this report. Full list is attached as CSV.</p>
    {extra}
    <table border="1" cellpadding="6" cellspacing="0">
        <thead><tr>{header}</tr></thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    """


def _report_csv_bytes(records: list) -> bytes:
    buffer = io.StringIO()
    fieldnames = [key for key, _ in REPORT_COLUMNS]
    writer = csv.DictWriter(buffer, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    for record in records:
        writer.writerow({key: ("" if record.get(key) is None else record.get(key)) for key in fieldnames})
    return buffer.getvalue().encode("utf-8-sig")


def send_daily_report_email(records: list, target_date: str) -> None:
    """
    Email report rows as CSV.
    From support@applywizz.ai (or SENDER_EMAIL) to dinesh@applywizz.ai.
    Runs independently of Supabase success/failure.
    """
    if not records:
        print("ℹ️ No records to email.")
        return

    token = _get_graph_access_token()
    attachment = base64.b64encode(_report_csv_bytes(records)).decode("ascii")
    file_stem = _safe_filename_date(target_date)
    payload = {
        "message": {
            "subject": f"{APP_NAME}: {len(records)} LinkedIn job posting(s) - {target_date}",
            "body": {"contentType": "HTML", "content": _report_html(records, target_date)},
            "toRecipients": [{"emailAddress": {"address": REPORT_RECIPIENT}}],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": f"linkedin_jobs_{file_stem}.csv",
                    "contentType": "text/csv",
                    "contentBytes": attachment,
                }
            ],
        },
        "saveToSentItems": True,
    }
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=120,
    )
    if response.status_code not in (200, 202):
        raise RuntimeError(
            f"Microsoft Graph sendMail failed ({response.status_code}): {response.text}"
        )
    print(f"✅ Daily report emailed to {REPORT_RECIPIENT} from {SENDER_EMAIL}.")


# -----------------------
# ENDPOINT: Process LinkedIn Job Postings
# -----------------------
@app.post("/process-linkedin-jobs")
def process_linkedin_jobs(target_date: Optional[str] = None):
    """
    Get LinkedIn job postings with poster profiles for the current IST date
    (or an optional target_date), try Supabase insert, and always email the CSV.
    """
    if target_date is None:
        target_date = today_ist()

    print(f"🔍 Fetching LinkedIn jobs for date: {target_date}")

    job_postings = get_linkedin_job_postings(target_date)
    
    if not job_postings:
        return {
            "success": True,
            "message": f"No LinkedIn job postings found for {target_date}!",
            "jobs_count": 0,
            "inserted": False,
            "emailed": False,
            "date": target_date,
        }

    records = build_report_records(job_postings, target_date)

    # Supabase is best-effort — email still goes out if this fails
    inserted = insert_jobs_to_supabase(records)

    try:
        send_daily_report_email(records, target_date)
        emailed = True
    except Exception as e:
        print(f"❌ Error sending daily report email: {e}")
        raise HTTPException(status_code=500, detail=f"Daily report email failed: {str(e)}")
    
    print(f"✅ Processing completed for {len(job_postings)} jobs (inserted={inserted}, emailed={emailed})")
   
    return {
        "success": True,
        "message": (
            f"Found {len(job_postings)} LinkedIn job posting(s) for {target_date}; "
            f"emailed {REPORT_RECIPIENT}"
            + ("; stored in Supabase." if inserted else "; Supabase insert skipped/failed.")
        ),
        "jobs_count": len(job_postings),
        "inserted": inserted,
        "emailed": emailed,
        "email_to": REPORT_RECIPIENT,
        "email_from": SENDER_EMAIL,
        "date": target_date,
    }


# Health check endpoint
@app.get("/")
def health_check():
    return {"status": "ok", "service": "LinkedIn Job Postings Processor API"}

# -----------------------
# Main execution when run as script (for cron jobs)
# -----------------------
if __name__ == "__main__":
    print("🚀 Starting LinkedIn job postings check...")
    try:
        # You can optionally parse args here if you want to pass a date via CLI
        result = process_linkedin_jobs()
        print(f"✅ Execution completed successfully!")
        print(f"Result: {result}")
    except Exception as e:
        print(f"❌ Error occurred: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
