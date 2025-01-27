from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, date
import calendar
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# Test Mode Configuration
TEST_MODE = False  # Set to False for production
TEST_VENDOR = "Reliant Test Account"

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'  # Updated to absolute path
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SHEET_NAME = 'VENDORS'

# Load environment variables
load_dotenv()
SENDER_EMAIL = os.getenv('sender_email')
APP_PASSWORD = os.getenv('app_password')

# Email Templates and COI Information
EMAIL_SUBJECT = "Request for updated Certificate of Insurance"

EMAIL_BODY_TEMPLATE = """Dear {vendor_name},

Our records indicate that your Certificate of Insurance (COI) expired on {expiration_date}. 
Could you please send an updated COI to this email at your earliest convenience?

For your updated COI, kindly ensure the following text is included:

Insured and Additionally Insured:
Reliant Property Management P.O. BOX 1630, Arvada, 80001 Colorado as Certificate Holder
{property_code_COI_TEMPALTE}

Thanks in advance for taking care of this!

Best regards,
Reliant Property Management"""

# COI Information by Property Code
COI_TEMPLATES = {
    "100/101/102/104": """
Flocchini-Magnolia Associates LLC as Additionally Insured
Jeanine Landsinger GST Trust as Additionally Insured
4997 Longley Lane Associates as Additionally Insured""",

    "109": """
Flocchini Associates LLC as Additionally Insured""",

    "111/113": """
South Federal Park and Green Street Associates, TIC as Additionally Insured
South Federal Park Associates as Additionally Insured""",

    "105/106/107": """
Flocchini Family Holdings Orem as Additionally Insured"""
}

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
            
            # Skip non-test vendors when in test mode
            if TEST_MODE and vendor.strip() != TEST_VENDOR:
                continue
                
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

def get_coi_information(property_code):
    """Get the correct COI information based on property code."""
    if property_code.startswith(('100', '101', '102', '104')):
        return COI_TEMPLATES.get('100/101/102/104', '')
    elif property_code.startswith('109'):
        return COI_TEMPLATES.get('109', '')
    elif property_code.startswith(('111', '113')):
        return COI_TEMPLATES.get('111/113', '')
    elif property_code.startswith(('105', '106', '107')):
        return COI_TEMPLATES.get('105/106/107', '')
    return ''

def format_email_content(entry):
    """Format email content for a vendor."""
    try:
        coi_info = get_coi_information(entry.code)
        
        email_body = EMAIL_BODY_TEMPLATE.format(
            vendor_name=entry.vendor_name,
            expiration_date=entry.formatted_date,
            property_code_COI_TEMPALTE=coi_info
        )
        
        return EMAIL_SUBJECT, email_body
    except Exception as e:
        print(f"Error formatting email: {e}")
        return "Error in subject", "Error in email body"

def send_email(to_email, subject, body):
    """Send email using SMTP."""
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        # Setup SMTP server connection
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        
        # Send email
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, to_email, text)
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def print_expired_summary(entries):
    """Print a concise summary of expired entries."""
    print("\nVendors with Expired COIs:")
    print("==========================")
    for entry in entries:
        print(f"{entry.vendor_name} - Expired: {entry.formatted_date}")

def get_user_confirmation(count):
    """Get user confirmation before sending emails."""
    while True:
        response = input(f"\nWould you like to proceed with sending {count} emails? (y/n): ").lower()
        if response in ['y', 'n']:
            return response == 'y'
        print("Please enter 'y' for yes or 'n' for no.")

def get_mode_selection():
    """Get user selection for automatic or manual mode"""
    while True:
        print("\nSelect Mode:")
        print("1. Automatic (process all expired COIs)")
        print("2. Manual (select specific vendors)")
        choice = input("Enter choice (1 or 2): ").strip()
        if choice in ['1', '2']:
            return choice == '2'  # Returns True for manual mode
        print("Invalid choice. Please enter 1 or 2.")

def get_manual_emails():
    """Get list of email addresses from user input"""
    print("\nEnter email addresses (one per line)")
    print("Press Enter twice when finished")
    emails = []
    while True:
        email = input().strip().lower()
        if not email:
            if emails:  # Only break if we have at least one email
                break
            print("Please enter at least one email address.")
            continue
        emails.append(email)
    return emails

def get_vendor_entry(email, service):
    """Get vendor entry for a specific email"""
    try:
        # Get all relevant columns
        range_name = f'{SHEET_NAME}!A2:G'  # Include all needed columns
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        for i, row in enumerate(values, 2):
            if len(row) > 4 and row[4].lower().strip() == email.lower().strip():
                return VendorEntry(
                    row=i,
                    code=row[0],
                    vendor_name=row[1],
                    email=row[4],
                    exp_date=row[6] if len(row) > 6 else None,
                    formatted_date=format_date(row[6] if len(row) > 6 else None)[1]
                )
        return None
    except Exception as e:
        print(f"Error fetching vendor details for {email}: {e}")
        return None

def main():
    if TEST_MODE:
        print("\n*** TEST MODE ENABLED - Only processing test vendor ***")
        print(f"Test Vendor: {TEST_VENDOR}\n")
    
    # Get mode selection
    manual_mode = get_mode_selection()
    
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        
        if manual_mode:
            print("\nManual Mode Selected")
            emails = get_manual_emails()
            expired_entries = []
            
            for email in emails:
                entry = get_vendor_entry(email, service)
                if entry:
                    expired_entries.append(entry)
                    print(f"Found vendor: {entry.vendor_name}")
                else:
                    print(f"❌ Could not find vendor for email: {email}")
        else:
            print("\nAutomatic Mode Selected")
            expired_entries = get_expired_dates()
    
        # Initialize statistics
        total_to_process = len(expired_entries)
        emails_sent = 0
        failed_emails = []
        
        if expired_entries:
            # Print concise summary
            print_expired_summary(expired_entries)
            print(f"\nTotal COIs to process: {total_to_process}")
            
            # Get user confirmation
            if get_user_confirmation(total_to_process):
                print("\nProcessing Emails...")
                
                for entry in expired_entries:
                    try:
                        subject, email_body = format_email_content(entry)
                        if send_email(entry.email, subject, email_body):
                            emails_sent += 1
                            print(f"✓ Sent to {entry.vendor_name}")
                        else:
                            failed_emails.append(f"{entry.vendor_name} ({entry.email})")
                            print(f"✗ Failed: {entry.vendor_name}")
                    except Exception as e:
                        failed_emails.append(f"{entry.vendor_name} ({entry.email}): {str(e)}")
                        print(f"✗ Error: {entry.vendor_name}")
                
                # Print final statistics
                print("\nEmail Processing Complete")
                print("========================")
                print(f"Successfully Sent: {emails_sent}")
                if failed_emails:
                    print(f"\nFailed Emails ({len(failed_emails)}):")
                    for failure in failed_emails:
                        print(f"- {failure}")
            else:
                print("\nOperation cancelled by user.")
        else:
            print("\nNo vendors found to process.")
            
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
