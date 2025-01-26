from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, date
import calendar
import os

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = 'sheets_key.json'
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SHEET_NAME = 'VENDORS'
TEMPLATE_FILE = 'email_templates.txt'

# Create a class to hold entry data
class VendorEntry:
    def __init__(self, row, code, vendor_name, email, exp_date, formatted_date):
        self.row = row
        self.code = code
        self.vendor_name = vendor_name
        self.email = email
        self.exp_date = exp_date
        self.formatted_date = formatted_date

def format_date(date_str):
    """Convert 6-digit date to formatted string and check if expired."""
    if not date_str or date_str.upper() == 'N/A':
        return None, 'N/A'
    
    try:
        # Convert string to date object (assuming format MMDDYY)
        date_obj = datetime.strptime(date_str, '%m%d%y').date()
        
        # Get day suffix (st, nd, rd, th)
        day = date_obj.day
        if 10 <= day % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        
        # Format date with full month name, day with suffix, and full year
        formatted = date_obj.strftime(f'%B %-d{suffix} %Y')
        return date_obj, formatted
        
    except ValueError:
        return None, date_str

def get_column_data(service, range_name):
    """Helper function to get column values."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name
    ).execute()
    return result.get('values', [])

def get_last_row(service):
    """Find the last row with data in column B (vendor names)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{SHEET_NAME}!B:B'
    ).execute()
    
    values = result.get('values', [])
    return len(values)  # This will give us the last row number

def get_expired_dates():
    """Read dates and identify expired ones with corresponding data."""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        
        # Get the last row dynamically
        last_row = get_last_row(service)
        
        # Define ranges dynamically
        DATE_RANGE = f'G2:G{last_row}'
        CODE_RANGE = f'A2:A{last_row}'
        EMAIL_RANGE = f'E2:E{last_row}'
        VENDOR_RANGE = f'B2:B{last_row}'
        
        # Get all needed column data
        dates = get_column_data(service, f'{SHEET_NAME}!{DATE_RANGE}')
        codes = get_column_data(service, f'{SHEET_NAME}!{CODE_RANGE}')
        emails = get_column_data(service, f'{SHEET_NAME}!{EMAIL_RANGE}')
        vendors = get_column_data(service, f'{SHEET_NAME}!{VENDOR_RANGE}')
        
        if not dates:
            print('No dates found.')
            return []
        
        today = date.today()
        print(f"\nChecking for dates before {today.strftime('%B %d, %Y')}...")
        
        expired_entries = []
        for i, (date_row, code_row, vendor_row, email_row) in enumerate(zip(dates, codes, vendors, emails), 2):
            date_str = date_row[0] if date_row else 'N/A'
            code = code_row[0] if code_row else 'N/A'
            vendor = vendor_row[0] if vendor_row else 'N/A'
            email = email_row[0] if email_row else 'N/A'
            
            date_obj, formatted_date = format_date(date_str)
            
            if date_obj and date_obj < today:
                entry = VendorEntry(i, code, vendor, email, date_obj, formatted_date)
                expired_entries.append(entry)
                print(f"\nExpired Entry Found:")
                print(f"Row {i}:")
                print(f"  Property Code: {code}")
                print(f"  Vendor Name: {vendor}")
                print(f"  Email: {email}")
                print(f"  Expiration: {formatted_date}")
        
        return expired_entries
        
    except Exception as e:
        print(f'Error: {e}')
        return []

def load_email_templates():
    """Load email templates and COI information."""
    templates = {}
    try:
        with open(TEMPLATE_FILE, 'r') as f:
            content = f.read()
            
        # Parse the template file
        sections = content.split('\n\n')
        for section in sections:
            if '=' in section:
                key, template = section.split('=', 1)
                templates[key.strip()] = template.strip()
    except Exception as e:
        print(f"Error loading templates: {e}")
    
    return templates

def get_coi_information(property_code):
    """Get the correct COI information based on property code."""
    templates = load_email_templates()
    
    if property_code.startswith(('100', '101', '102', '104')):
        return templates.get('100/101/102/104_INFORMATION', '')
    elif property_code.startswith('109'):
        return templates.get('109_COI_INFORMATION', '')
    elif property_code.startswith(('111', '113')):
        return templates.get('111/113_COI_INFORMATION', '')
    elif property_code.startswith(('105', '106', '107')):
        return templates.get('105/106/107_COI_INFORMATION', '')
    return ''

# Add EMAIL_BODY template directly in the code
EMAIL_BODY_TEMPLATE = """Dear {vendor_name},

Our records indicate that your Certificate of Insurance (COI) expired on {expiration_date}. 
Could you please send an updated COI to this email at your earliest convenience?

For your updated COI, kindly ensure the following text is included:

Insured and Additionally Insured:
Reliant Property Management P.O. BOX 1630, Arvada, Colorado as Certificate Holder
{property_code_COI_TEMPALTE}

Thanks in advance for taking care of this!

Best regards,
Reliant Property Management"""

def format_email_content(entry):
    """Format email content for a vendor."""
    try:
        templates = load_email_templates()
        coi_info = get_coi_information(entry.code)
        
        # Use the direct template instead of loading from file
        email_body = EMAIL_BODY_TEMPLATE.format(
            vendor_name=entry.vendor_name,
            expiration_date=entry.formatted_date,
            property_code_COI_TEMPALTE=coi_info
        )
        
        subject = templates.get('COI_SUBJECT_TEMPLATE', 'COI Update Required')
        
        return subject, email_body
    except Exception as e:
        print(f"Error formatting email: {e}")
        return "Error in subject", "Error in email body"

def main():
    print("Checking for expired dates and corresponding information...")
    expired_entries = get_expired_dates()
    
    if expired_entries:
        print(f"\nFound {len(expired_entries)} expired entries.")
        print("\nEmail Preview for each entry:")
        print("==============================")
        
        for entry in expired_entries:
            subject, email_body = format_email_content(entry)
            print(f"\nTo: {entry.email}")
            print(f"Subject: {subject}")
            print("Body:")
            print("-" * 50)
            print(email_body)
            print("-" * 50)
            print("\n")
    else:
        print("\nNo expired dates found")

if __name__ == '__main__':
    main()
