#karmafy/karmafy/api_send_email_notification.py

from fastapi import FastAPI, HTTPException
from typing import Optional
import os
from datetime import datetime
import psycopg2
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
    Query the karmafy_job table to find all LinkedIn job postings for a specific date
    that have a poster profile.
    
    Args:
        target_date: Date in 'YYYY-MM-DD' format. Defaults to today's date if not provided.
    
    Returns list of dicts with job posting information.
    """
    conn = get_db_connection()
    try:
        # If no date provided, use today's date
        if target_date is None:
            target_date = datetime.now().strftime('%Y-%m-%d')
        
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
                    kjr.name AS job_role
                FROM public.karmafy_job kj
                LEFT JOIN public.karmafy_jobrole kjr
                    ON kj."roleId"::bigint = kjr.id
                WHERE kj.source = 'LINKEDIN'
                    AND DATE(kj."ingestedAt") = %s
                    AND kj.posted_by_profile IS NOT NULL
                    AND kj.posted_by_profile != ''
                ORDER BY kj.posted_by_profile, kj.company, kj.title
            """
            cursor.execute(query, (target_date,))
            results = cursor.fetchall()
            
            # Convert to list of dicts
            return [dict(row) for row in results]
    
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
        target_date: Date string YYYY-MM-DD
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
                'report_date': target_date
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


# -----------------------
# ENDPOINT: Process LinkedIn Job Postings
# -----------------------
@app.post("/process-linkedin-jobs")
def process_linkedin_jobs(target_date: Optional[str] = None):
    """
    Get all LinkedIn job postings from today (or target_date) that have poster profiles
    and store them in Supabase.
    """
    if target_date is None:
        target_date = datetime.now().strftime('%Y-%m-%d')

    print(f"🔍 Fetching LinkedIn jobs for date: {target_date}")

    # Query database for LinkedIn job postings
    job_postings = get_linkedin_job_postings(target_date)
    
    if not job_postings:
        return {
            "success": True,
            "message": f"No LinkedIn job postings found for {target_date}!",
            "jobs_count": 0,
            "inserted": False
        }
    
    # Insert jobs into Supabase
    insert_jobs_to_supabase(job_postings, target_date)
    
    print(f"✅ Processing completed for {len(job_postings)} jobs")
   
    return {
        "success": True,
        "message": f"Found and stored {len(job_postings)} LinkedIn job posting(s) into Supabase.",
        "jobs_count": len(job_postings),
        "inserted": True,
        "date": target_date
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
