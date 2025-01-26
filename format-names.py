try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("Required Google packages not found. Please install them using:")
    print("pip3 install --user google-api-python-client google-auth-httplib2 google-auth-oauthlib")
    exit(1)

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Note: removed readonly
SERVICE_ACCOUNT_FILE = 'sheets_key.json'
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SHEET_NAME = 'VENDORS'
RANGE = 'C2:C79'

def format_name(name):
    """Format name string according to requirements."""
    if not name:
        return 'N/A'
    
    # Handle N/A cases
    if name.strip().upper() == 'N/A' or name.strip().lower() == 'n/a':
        return 'N/A'
    
    # Split name into parts and capitalize each part
    name_parts = name.strip().split()
    return ' '.join(part.upper() for part in name_parts)

def update_names():
    """Read, format, and update names in column C."""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        
        # Get existing values
        range_name = f'{SHEET_NAME}!{RANGE}'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print('No data found.')
            return
        
        # Process names
        updated_values = []
        for row in values:
            original = row[0] if row else ''
            formatted = format_name(original)
            updated_values.append([formatted])
            print(f"Original: {original:20} -> Formatted: {formatted}")
        
        # Update the sheet
        body = {
            'values': updated_values
        }
        
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        print(f"\nSuccessfully updated {len(updated_values)} names")
        
    except Exception as e:
        print(f'Error: {e}')

def main():
    print("Processing names in column C...")
    update_names()

if __name__ == '__main__':
    main()
