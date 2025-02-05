from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
import os

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Full access needed
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'  # Updated to include range

def initialize_sheets_service():
    """Initialize and return the Google Sheets service"""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except Exception as e:
        print(f"Error initializing sheets service: {e}")
        return None

def main():
    # Initialize the service
    service = initialize_sheets_service()
    if not service:
        print("Failed to initialize Google Sheets service")
        return

if __name__ == '__main__':
    main()
