try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("Required Google packages not found. Please install them using:")
    print("pip3 install --user google-api-python-client google-auth-httplib2 google-auth-oauthlib")
    exit(1)

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = 'sheets_key.json'
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SHEET_NAME = 'VENDORS'
RANGE = 'D2:D79'

def clean_phone_number(phone):
    """Convert phone string to clean number format."""
    if not phone or phone.strip().upper() == 'N/A' or phone.strip().lower() == 'n/a':
        return 'N/A'
    
    # Extract only digits
    digits = ''.join(char for char in phone if char.isdigit())
    
    # Check if we have a valid 10-digit phone number
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return phone  # Return original if not 10 digits

def update_phone_numbers():
    """Read, clean, and update phone numbers in column D."""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets'])
        service = build('sheets', 'v4', credentials=credentials)
        
        # First get existing values
        range_name = f'{SHEET_NAME}!{RANGE}'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print('No data found.')
            return
        
        # Process phone numbers
        updated_values = []
        for row in values:
            original = row[0] if row else ''
            cleaned = clean_phone_number(original)
            updated_values.append([cleaned])
            print(f"Original: {original:20} -> Cleaned: {cleaned}")
        
        # Update the sheet with cleaned numbers
        body = {
            'values': updated_values
        }
        
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        print(f"\nSuccessfully updated {len(updated_values)} phone numbers")
        
    except Exception as e:
        print(f'Error: {e}')

def main():
    print("Processing phone numbers in column D...")
    update_phone_numbers()

if __name__ == '__main__':
    main()
