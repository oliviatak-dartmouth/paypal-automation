# PayPal to Google Sheets Automation - Example Configuration
#
# Copy this file to config.py and fill in your own values:
#     cp config.example.py config.py
# config.py is gitignored so your real Sheet ID is never committed.

import os

# Google Sheet ID (from the URL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)
GOOGLE_SHEET_ID = "your-sheet-id-here"

# Path to Google service account credentials (download from Google Cloud Console).
# Keep this file private and out of version control.
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")

# Folder containing PayPal CSV exports
CSV_FOLDER = "paypal_csvs/"

# Auto-create monthly sheets without prompting
AUTO_CREATE_MONTHLY_SHEETS = True

# Retry settings for API calls
MAX_RETRIES = 3
RETRY_DELAY_BASE = 1  # Base delay in seconds (exponential backoff: 1s, 2s, 4s)
