from google.oauth2 import service_account
from googleapiclient.discovery import build

# Define the scope and credentials path
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'sheets_key.json'
SHEET_NAME = 'VENDORS'  # Add this constant

def get_sheet_names(service, spreadsheet_id):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    return [sheet['properties']['title'] for sheet in sheets]

def main():
    # Set up credentials
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    # Build the Sheets API service
    service = build('sheets', 'v4', credentials=credentials)
    
    # Spreadsheet ID from your URL
    SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
    
    try:
        # Use the known sheet name directly
        range_name = f"'{SHEET_NAME}'!A1"
        
        values = [
            ['Hello']
        ]
        body = {
            'values': values
        }
        
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f"Successfully wrote to sheet: {SHEET_NAME}")
            
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()