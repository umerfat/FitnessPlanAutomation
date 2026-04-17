import os
from dotenv import load_dotenv

load_dotenv()

# Google Sheets
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Form_Responses")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

# Google Gemini
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# Email
SMTP_EMAIL = os.getenv("SMTP_EMAIL")
SMTP_APP_PASSWORD = os.getenv("SMTP_APP_PASSWORD")
COACH_EMAIL = os.getenv("COACH_EMAIL")

# Brand
BRAND_NAME = os.getenv("BRAND_NAME", "UMER HURRAH")

# Paths
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
PROCESSED_FILE = os.path.join(os.path.dirname(__file__), "processed_clients.json")

# Defaults
DEFAULT_TRAINING_DAYS = 5  # Mon-Fri
