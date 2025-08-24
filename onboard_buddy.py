import os
import json
import smtplib
from email.mime.text import MIMEText
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import List

import requests
from dotenv import load_dotenv

# Google OAuth & API client libs
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request


# ---------- ENV ----------
load_dotenv()

SHEET_ID = os.getenv("SHEET_ID", "")
SHEET_TAB = os.getenv("SHEET_TAB", "SHEET1")  # <-- your default tab name
WINDOW_DAYS = int(os.getenv("WINDOW_DAYS", "7"))
CALENDAR_ID = os.getenv("CALENDAR_ID", "primary")
TIMEZONE = os.getenv("TIMEZONE", "Asia/Kolkata")

GMAIL_USER = os.getenv("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
SLACK_WEBHOOK = os.getenv("SLACK_WEBHOOK", "")


# ---------- AUTH HELPERS ----------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/calendar",
]

def get_credentials(scopes):
    """
    Create/refresh OAuth creds and cache them in token.json.
    If the existing token doesn't include all required scopes,
    we force a new consent flow.
    """
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", scopes)

        # 🔴 force re-consent if scopes don’t match
        have = set((creds.scopes or []))
        need = set(scopes)
        if not need.issubset(have):
            creds = None  # trigger new flow below

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return creds

def get_sheets_service():
    creds = get_credentials(SCOPES)     # ✅ use combined scopes
    return build("sheets", "v4", credentials=creds)

def get_calendar_service():
    creds = get_credentials(SCOPES)     # ✅ use combined scopes
    return build("calendar", "v3", credentials=creds)





# ---------- DATA MODEL ----------
@dataclass
class Employee:
    name: str
    email: str
    department: str
    start_date: str  # YYYY-MM-DD
    manager: str

    def start_date_obj(self):
        return datetime.strptime(self.start_date, "%Y-%m-%d").date()






# ---------- SHEETS PARSING ----------
def parse_upcoming_employees(values, window_days: int):
    if not values:
        return []

    header = [h.strip().lower() for h in values[0]]

    def col(name, default):
        try:
            return header.index(name)
        except ValueError:
            return default

    i_name = col("name", 0)
    i_email = col("email", 1)
    i_dept = col("department", 2)
    i_start = col("startdate", 3)
    i_mgr = col("manager", 4)

    today = datetime.now(timezone.utc).date()
    within = today + timedelta(days=window_days)

    hires = []
    for row in values[1:]:
        if len(row) < 5:
            continue
        try:
            d = datetime.strptime(row[i_start].strip(), "%Y-%m-%d").date()
        except Exception:
            continue
        if today <= d <= within:
            hires.append(
                Employee(
                    name=row[i_name].strip(),
                    email=row[i_email].strip(),
                    department=row[i_dept].strip(),
                    start_date=row[i_start].strip(),
                    manager=row[i_mgr].strip(),
                )
            )
    return hires


# ---------- CALENDAR ----------
def create_day1_event(cal_service, calendar_id: str, tz: str, emp: Employee):
    """Create a 30-minute 'Day 1 Orientation' event and invite the employee."""
    start_dt = datetime.strptime(emp.start_date, "%Y-%m-%d").replace(hour=10, minute=0)
    end_dt = start_dt + timedelta(minutes=30)

    event = {
        "summary": f"Day 1 Orientation: {emp.name}",
        "description": (
            f"Welcome {emp.name} to {emp.department}!\n"
            f"Manager: {emp.manager}\n"
            "Agenda: HR paperwork, accounts, laptop, and intro."
        ),
        "start": {"dateTime": start_dt.isoformat(), "timeZone": tz},
        "end": {"dateTime": end_dt.isoformat(), "timeZone": tz},
        "attendees": [{"email": emp.email}],
        # To auto-create a Google Meet link, uncomment below and add conferenceDataVersion=1 in insert()
        # "conferenceData": {"createRequest": {"requestId": str(uuid.uuid4())}},
    }

    created = cal_service.events().insert(
        calendarId=calendar_id,
        body=event,
        # conferenceDataVersion=1,
        sendUpdates="all",
    ).execute()

    return created.get("htmlLink")


# ---------- EMAIL ----------
def build_welcome_email(emp: Employee) -> str:
    return (
        f"Hi {emp.name},\n\n"
        f"Welcome to the {emp.department} team! Your start date is {emp.start_date}.\n"
        f"You'll report to {emp.manager}. Before Day 1 you'll receive account setup, laptop, and schedule details.\n\n"
        f"See you soon,\nHR Team"
    )


def send_email_if_configured(emp: Employee) -> bool:
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("✉️  Gmail creds not set — skipping email.")
        return True  # skip but don't fail

    try:
        msg = MIMEText(build_welcome_email(emp), _charset="utf-8")
        msg["Subject"] = f"Welcome to the team, {emp.name}!"
        msg["From"] = GMAIL_USER
        msg["To"] = emp.email

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            smtp.send_message(msg)

        print(f"✅ Email sent to {emp.email}")
        return True
    except Exception as e:
        print("❌ Email error:", e)
        return False


# ---------- SLACK ----------
def post_slack_if_configured(emp: Employee) -> bool:
    if not SLACK_WEBHOOK:
        print("💬 Slack webhook not set — skipping Slack.")
        return True  # skip but don't fail

    payload = {
        "text": f"🎉 New Team Member: {emp.name} starts {emp.start_date} in {emp.department} (Mgr: {emp.manager})"
    }
    try:
        r = requests.post(SLACK_WEBHOOK, json=payload, timeout=10)
        r.raise_for_status()
        print(f"✅ Slack posted for {emp.name}")
        return True
    except Exception as e:
        print("❌ Slack error:", e)
        return False


# ---------- MAIN ----------
def main():
    if not SHEET_ID:
        print("❗ Please set SHEET_ID in your .env")
        return

    # 1) Sheets: read A:E
    sheets = get_sheets_service()
    resp = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"{SHEET_TAB}!A:E",
    ).execute()
    values = resp.get("values", [])
    print(f"✅ Got {len(values)} rows from {SHEET_TAB}!A:E")

    # 2) Parse windowed hires
    hires = parse_upcoming_employees(values, WINDOW_DAYS)
    print(f"🔎 Upcoming hires in next {WINDOW_DAYS} days: {len(hires)}")
    if not hires:
        print("Nothing to process.")
        return

    # 3) Calendar service once
    cal = get_calendar_service()

    # 4) Process each hire
    for emp in hires:
        ok_email = send_email_if_configured(emp)
        ok_slack = post_slack_if_configured(emp)
        link = create_day1_event(cal, CALENDAR_ID, TIMEZONE, emp)
        print(f"📅 Event created for {emp.name}: {link} | email={ok_email} slack={ok_slack}")


if __name__ == "__main__":
    # If you just added Calendar scope for the first time, delete token.json once and rerun.
    main()


