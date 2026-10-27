#karmafy/karmafy/api_send_email_notification.py

from fastapi import FastAPI, HTTPException
from typing import Optional
import base64
import html
import io
import os
from datetime import datetime
import pandas as pd
import psycopg2
import requests
from psycopg2.extras import RealDictCursor
from supabase import create_client, Client

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
    Query the karmafy_job table for LinkedIn job postings that have a poster profile.

    For now (USE_ANY_AVAILABLE_DATES): no date filter — returns whatever dates have data.
    """
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
                    DATE(kj."ingestedAt") AS ingested_date
                FROM public.karmafy_job kj
                LEFT JOIN public.karmafy_jobrole kjr
                    ON kj."roleId"::bigint = kjr.id
                WHERE kj.source = 'LINKEDIN'
                    AND kj.posted_by_profile IS NOT NULL
                    AND kj.posted_by_profile != ''
                ORDER BY kj.posted_by_profile, kj."ingestedAt" DESC NULLS LAST, kj.company, kj.title
            """
            cursor.execute(query)
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
# Insert into Supabase
# -----------------------
def insert_jobs_to_supabase(jobs_data: list, target_date: str) -> None:
    """
    Insert job postings data into Supabase 'daily_linkedin_jobs_report' table.
    
    Args:
        jobs_data: List of job posting dictionaries
        target_date: Fallback date string YYYY-MM-DD when a row has no ingested_date
    """
    try:
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        
        # Prepare data for insertion
        records_to_insert = []
        for job in jobs_data:
            records_to_insert.append({
                'company': job.get('company'),
                'job_title': job.get('title'),
                'job_role': job.get('job_role'),
                'poster_full_name': job.get('poster_full_name'),
                'poster_profile_url': job.get('posted_by_profile'),
                'job_url': job.get('url'),
                'company_url': job.get('company_url'),
                'source': job.get('source'),
                'report_date': job.get('ingested_date') or target_date
            })
            
        if not records_to_insert:
            print("ℹ️ No records to insert.")
            return

        # Perform insertion
        # Note: If handling large datasets, consider batching (e.g., chunks of 100)
        data = supabase.table("daily_linkedin_jobs_report").insert(records_to_insert).execute()
        
        print(f"✅ Successfully inserted {len(records_to_insert)} records into Supabase.")
        
    except Exception as e:
        print(f"❌ Error inserting into Supabase: {e}")
        raise HTTPException(status_code=500, detail=f"Supabase insertion failed: {str(e)}")

    try:
        send_daily_report_email(records_to_insert, target_date)
    except Exception as e:
        print(f"❌ Error sending daily report email: {e}")
        raise HTTPException(status_code=500, detail=f"Daily report email failed: {str(e)}")


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


def _report_html(records: list, target_date: str) -> str:
    columns = [
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
    header = "".join(f"<th>{html.escape(label)}</th>" for _, label in columns)
    rows = []
    for record in records:
        cells = "".join(
            f"<td>{html.escape('' if record.get(key) is None else str(record.get(key)))}</td>"
            for key, _ in columns
        )
        rows.append(f"<tr>{cells}</tr>")

    return f"""
    <p>Daily LinkedIn job postings for <strong>{html.escape(target_date)}</strong>.</p>
    <p>{len(records)} posting(s) were stored in the daily report.</p>
    <table border="1" cellpadding="6" cellspacing="0">
        <thead><tr>{header}</tr></thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    """


def _report_excel_bytes(records: list) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(records).to_excel(buffer, index=False, sheet_name="LinkedIn jobs")
    return buffer.getvalue()


def send_daily_report_email(records: list, target_date: str) -> None:
    """
    Email the same daily rows written to Supabase.
    From support@applywizz.ai (or SENDER_EMAIL) to dinesh@applywizz.ai.
    """
    if not records:
        print("ℹ️ No records to email.")
        return

    token = _get_graph_access_token()
    attachment = base64.b64encode(_report_excel_bytes(records)).decode("ascii")
    payload = {
        "message": {
            "subject": f"{APP_NAME}: {len(records)} LinkedIn job posting(s) - {target_date}",
            "body": {"contentType": "HTML", "content": _report_html(records, target_date)},
            "toRecipients": [{"emailAddress": {"address": REPORT_RECIPIENT}}],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": f"linkedin_jobs_{target_date}.xlsx",
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        timeout=60,
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
    Get LinkedIn job postings with poster profiles from any available dates,
    store them in Supabase, and email the report.
    """
    run_date = datetime.now().strftime("%Y-%m-%d")

    print("🔍 Fetching LinkedIn jobs for any available dates (no date filter)...")

    # Query database for LinkedIn job postings
    job_postings = get_linkedin_job_postings()
    
    if not job_postings:
        return {
            "success": True,
            "message": "No LinkedIn job postings with poster profiles found!",
            "jobs_count": 0,
            "inserted": False,
            "emailed": False
        }

    available_dates = sorted({
        job.get("ingested_date")
        for job in job_postings
        if job.get("ingested_date")
    })
    report_label = (
        f"{available_dates[0]} to {available_dates[-1]}"
        if len(available_dates) > 1
        else (available_dates[0] if available_dates else run_date)
    )
    print(f"📅 Available ingested dates in result: {available_dates}")
    
    # Insert jobs into Supabase (report_date per row uses each job's ingested_date)
    insert_jobs_to_supabase(job_postings, report_label)
    
    print(f"✅ Processing completed for {len(job_postings)} jobs")
   
    return {
        "success": True,
        "message": f"Found and stored {len(job_postings)} LinkedIn job posting(s) into Supabase and emailed {REPORT_RECIPIENT}.",
        "jobs_count": len(job_postings),
        "inserted": True,
        "emailed": True,
        "email_to": REPORT_RECIPIENT,
        "email_from": SENDER_EMAIL,
        "date": report_label,
        "available_dates": available_dates,
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
