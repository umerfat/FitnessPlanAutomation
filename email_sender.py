"""
Send emails with plan attachments via Gmail API (HTTPS, no SMTP needed).
"""

import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import config


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]
TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token.json")


def _write_secrets_from_env():
    """In CI, write credentials/token files from environment variables."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    if creds_json and not os.path.exists(config.GOOGLE_CREDENTIALS_FILE):
        with open(config.GOOGLE_CREDENTIALS_FILE, "w") as f:
            f.write(creds_json)
    if token_json and not os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "w") as f:
            f.write(token_json)


def _get_gmail_service():
    """Get authenticated Gmail API service."""
    _write_secrets_from_env()
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if os.environ.get("CI"):
                raise RuntimeError(
                    "Token expired and cannot re-authenticate in CI. "
                    "Run locally first to refresh token.json, then update the GOOGLE_TOKEN_JSON secret."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                config.GOOGLE_CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def send_plan_to_coach(client_name: str, client_email: str, filepath: str, metrics: dict):
    """Send the generated plan to the coach for review."""
    subject = f"[Review] New Client Plan — {client_name}"

    bmi = metrics["bmi"]
    cal = metrics["calories"]

    body = f"""New client plan generated and ready for review.

Client: {client_name}
Email: {client_email}
BMI: {bmi['bmi']} ({bmi['category']})
Target Calories: {cal['target_calories']} kcal
Goal: {cal['strategy']}

The plan document is attached. Review and edit if needed, then run:

    python send_to_client.py "{client_name}"

This will send the (potentially edited) plan to the client at {client_email}.
"""

    _send_email(
        to_email=config.COACH_EMAIL,
        subject=subject,
        body=body,
        attachment_path=filepath,
    )


def send_plan_to_client(client_name: str, client_email: str, filepath: str):
    """Send the reviewed plan to the client."""
    subject = f"Your Transformation Protocol — {config.BRAND_NAME}"

    body = f"""Hi {client_name},

Thank you for trusting {config.BRAND_NAME} with your fitness journey!

Your personalized Transformation Protocol is attached. This plan has been crafted based on your goals, body metrics, dietary preferences, and lifestyle.

Please review it carefully and reach out if you have any questions.

Let's get started!

Best regards,
{config.BRAND_NAME}
"""

    _send_email(
        to_email=client_email,
        subject=subject,
        body=body,
        attachment_path=filepath,
    )


def _send_email(to_email: str, subject: str, body: str, attachment_path: str = None):
    """Send an email via Gmail API with optional attachment."""
    msg = MIMEMultipart()
    msg["From"] = config.SMTP_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachment_path:
        path = Path(attachment_path)
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={path.name}")
        msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service = _get_gmail_service()
    result = service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()
    print(f"  Gmail API response: messageId={result.get('id')}, labelIds={result.get('labelIds')}")
