from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from datetime import datetime
import os

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Full access needed
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'  # Updated to include range

# Add template path and email subject constants
TEMPLATE_PATH = '/Users/cordo/Documents/RELIANT_SCRIPTS/Grease Trap Template.md'
EMAIL_SUBJECT = "Required Grease Trap Maintenance"

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

def get_email_from_entry(service, row_number):
    """Get email from Entry sheet for corresponding row"""
    try:
        # Note: row_number + 1 to shift reference by one row
        range_name = f'Entry!J{row_number + 1}'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        return values[0][0] if values and values[0] else 'N/A'
    except Exception as e:
        print(f"Error getting email from Entry sheet: {e}")
        return 'N/A'

def format_date_for_display(date_str):
    """Convert date string to readable format"""
    try:
        # Convert MMDDYY to datetime
        date_obj = datetime.strptime(date_str, '%m%d%y')
        return date_obj.strftime('%B %d, %Y')
    except ValueError:
        return date_str if date_str else 'N/A'

def read_email_template():
    """Read the email template file"""
    try:
        with open(TEMPLATE_PATH, 'r') as file:
            return file.read()
    except Exception as e:
        print(f"Error reading template file: {e}")
        return None

def format_email_content(template, tenant_name, tenant_address, last_service):
    """Format the email template with tenant's information"""
    try:
        # Format the last service date if it exists
        formatted_date = format_date_for_display(last_service)
        
        # Replace template variables
        content = template.replace("[Tenant Name]", tenant_name)
        content = content.replace("[Tenant Address]", tenant_address)
        content = content.replace("[Last Service Date]", formatted_date)
        return content
    except Exception as e:
        print(f"Error formatting email: {e}")
        return None

def get_grease_trap_data(service):
    """Fetch grease trap related data from Tenant sheet"""
    try:
        # Read email template first
        email_template = read_email_template()
        if not email_template:
            print("Failed to read email template")
            return

        # Get columns D, C, P, Q from Tenant sheet
        range_name = 'TENANT!C:Q'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print('No data found.')
            return
        
        print("\nGrease Trap Records:")
        print("===================")
        
        # Start from index 1 to skip header row
        for i, row in enumerate(values[1:], start=2):
            try:
                tenant_address = row[0] if len(row) > 0 else 'N/A'  # Column C
                tenant_name = row[1] if len(row) > 1 else 'N/A'     # Column D
                grease_required = row[13].strip().upper() if len(row) > 13 else 'N/A'  # Column P
                last_service = row[14] if len(row) > 14 else 'N/A'  # Column Q
                
                # Only process if grease trap is required
                if grease_required == 'YES':
                    # Get corresponding email from Entry sheet
                    tenant_email = get_email_from_entry(service, i)
                    
                    # Format email content
                    email_content = format_email_content(
                        email_template,
                        tenant_name,
                        tenant_address,
                        last_service
                    )
                    
                    print(f"\nRow {i}:")
                    print(f"Tenant: {tenant_name}")
                    print(f"Address: {tenant_address}")
                    print(f"Email: {tenant_email}")
                    print(f"Last Service Date: {last_service}")
                    print("\nEmail Content Preview:")
                    print("==============")
                    print(email_content)
                    print("==============\n")
                    
            except IndexError:
                print(f"Skipping row {i} - incomplete data")
                continue

    except Exception as e:
        print(f"Error reading grease trap data: {e}")

def main():
    # Initialize the service
    service = initialize_sheets_service()
    if not service:
        print("Failed to initialize Google Sheets service")
        return

    # Get and display grease trap data
    get_grease_trap_data(service)

if __name__ == '__main__':
    main()
