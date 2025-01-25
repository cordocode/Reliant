from google.oauth2 import service_account
from googleapiclient.discovery import build

# Define the scope and credentials path
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'sheets.key.json'

def main():
    # Set up credentials
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    # Build the Sheets API service
    service = build('sheets', 'v4', credentials=credentials)
    
    # Spreadsheet ID from your URL
    SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
    
    # Write 'Hello' to cell A1
    range_name = 'VENDOR_MANAGEMENT!A1'  # Updated to your sheet name
    values = [
        ['Hello']
    ]
    body = {
        'values': values
    }
    
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        print("Successfully wrote to spreadsheet!")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()