from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from datetime import datetime, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Add template path to constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'

# Add email configuration and template
load_dotenv()
SENDER_EMAIL = os.getenv('sender_email')
APP_PASSWORD = os.getenv('app_password')
EMAIL_SUBJECT = "HVAC Maintenance Contract Update Required"

EMAIL_TEMPLATE = """Dear [Tenant Name],

As outlined in the lease [lease section], we're reaching out about your HVAC maintenance for the property at [Tenant Address]. We know, exciting stuff, right?

We aim to keep things running smoothly behind the scenes. Our approach is straightforward:
- We track most recent invoices to ensure someone has been maintaining the system
- We'll check in every six months (should still be serviced quarterly)
- If we don't receive the requested invoices within 15 days, our maintenance team might need to get involved
- This could include a 10% service fee

Our motivation is simple: these HVAC systems are a significant investment, and we want to help you avoid unexpected breakdowns or costly replacements.

The previous record we had on file showed your most recent invoice on [Previous Contract Expiration Date]. Please send your most recent invoice and we will update our records.

Thanks for your help,

Relaint Property Management"""

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

def convert_date_string(date_str):
    """Convert 6-digit date string (MMDDYY) to datetime object"""
    try:
        return datetime.strptime(date_str, '%m%d%y')
    except ValueError:
        return None

def validate_email(email):
    """Validate email format and check for common invalid values"""
    if not email or email == 'N/A' or '@' not in email:
        return False
    return True

def get_email_from_entry(service, row_number):
    """
    Get email from Entry sheet for corresponding row.
    TENANT row 59 corresponds to ENTRY row 61 column J
    Therefore we need to add 2 to the row number and use column J
    """
    try:
        # Add 2 to the row number to align TENANT with ENTRY sheet
        entry_row = row_number + 2
        range_name = f'Entry!J{entry_row}'  # Changed back to column J
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        email = values[0][0] if values and values[0] else 'N/A'
        
        # Validate email before returning
        if validate_email(email):
            return email.strip()
        else:
            print(f"⚠️ Invalid email found in Entry row {entry_row}, column J: {email}")
            return None
            
    except Exception as e:
        print(f"Error getting email from Entry sheet: {e}")
        return None

def get_mode_selection():
    """Get user selection for processing mode"""
    while True:
        print("\nSelect Processing Mode:")
        print("1. Local Reliant Address Only")
        print("2. Manual Tenant Search")
        print("3. Full List")
        choice = input("Enter choice (1, 2, or 3): ").strip()
        if choice in ['1', '2', '3']:
            return int(choice)
        print("Invalid choice. Please enter 1, 2, or 3.")

def get_manual_tenant_name():
    """Get tenant name from user input"""
    while True:
        tenant_name = input("\nEnter tenant name (or 'q' to quit): ").strip()
        if tenant_name.lower() == 'q':
            return None
        if tenant_name:
            return tenant_name
        print("Please enter a valid tenant name.")

def get_lease_section(property_code):
    """Determine lease section based on property code"""
    if property_code == '104':
        return '(section 8.02)'
    elif property_code in ['111', '113']:
        return '(section 7.01)'
    elif property_code == '109':
        return ''  # Remove lease section entirely
    return '[lease section]'  # Default case

def format_date_for_display(date_obj):
    """Convert datetime object to readable format"""
    try:
        return date_obj.strftime('%B %d, %Y')
    except Exception:
        return 'N/A'

def format_email_content(template, tenant_name, unit_number, tenant_address, hvac_date):
    """Format the email template with the tenant's information"""
    try:
        property_code = unit_number.split('-')[0] if '-' in unit_number else unit_number
        lease_section = get_lease_section(property_code)
        
        # Format the expiration date
        formatted_date = format_date_for_display(hvac_date)
        
        # Replace all template variables
        content = template.replace("[Tenant Name]", tenant_name)
        content = content.replace("[lease section]", lease_section)
        content = content.replace("[Tenant Address]", tenant_address)
        content = content.replace("[Previous Contract Expiration Date]", formatted_date)
        return content
    except Exception as e:
        print(f"Error formatting email: {e}")
        return None

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
        print(f"✓ Email sent successfully to {to_email}")
        return True
    except Exception as e:
        print(f"✗ Error sending email to {to_email}: {e}")
        return False

def get_user_confirmation(count):
    """Get user confirmation before sending emails."""
    while True:
        response = input(f"\nWould you like to proceed with sending {count} emails? (y/n): ").lower()
        if response in ['y', 'n']:
            return response == 'y'
        print("Please enter 'y' for yes or 'n' for no.")

def is_within_notification_window(hvac_date):
    """
    Check if the HVAC date is older than 45 days.
    Returns True if we should send a notification (date is older than 45 days)
    """
    today = datetime.now()
    notification_cutoff = today - timedelta(days=45)
    
    # Return True only if the date is OLDER than our 45-day cutoff
    return hvac_date < notification_cutoff

def get_hvac_data(service, mode):
    """Read HVAC data including tenant names, emails, and unit numbers"""
    try:
        # Replace read_email_template() with EMAIL_TEMPLATE
        email_template = EMAIL_TEMPLATE
        
        range_name = 'TENANT!B:O'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print('No data found.')
            return
            
        print("\nProcessing HVAC Service Entries:")
        print("============================")
        if mode == 1:
            print("MODE: Local Reliant Address Only")
        elif mode == 2:
            print("MODE: Manual Tenant Search")
        else:
            print("MODE: Full List")
        print("============================")
        
        today = datetime.now()
        notification_cutoff = today - timedelta(days=45)
        expired_entries = []
        
        # Get manual tenant name if in mode 2
        target_tenant = None
        if mode == 2:
            target_tenant = get_manual_tenant_name()
            if not target_tenant:
                print("Operation cancelled.")
                return

        print(f"\nLooking for HVAC contracts that expired before:")
        print(f"  {notification_cutoff.strftime('%B %d, %Y')} (45 days ago)")
        
        # Start from index 1 to skip header row
        for i, row in enumerate(values[1:], start=2):
            unit_number = row[0] if len(row) > 0 else 'N/A'
            tenant_address = row[1] if len(row) > 1 else 'N/A'
            tenant_name = row[2] if len(row) > 2 else 'N/A'
            hvac_required = row[12].strip().upper() if len(row) > 12 else ''
            hvac_date_str = row[13].strip() if len(row) > 13 else ''
            
            # Apply filtering based on mode
            should_process = False
            if mode == 1:
                should_process = tenant_name.strip() == 'Reliant Property Management'
            elif mode == 2:
                should_process = tenant_name.lower().strip() == target_tenant.lower().strip()
            else:  # mode == 3
                should_process = True
            
            if should_process and hvac_required == 'YES' and hvac_date_str:
                hvac_date = convert_date_string(hvac_date_str)
                if hvac_date and is_within_notification_window(hvac_date):
                    tenant_email = get_email_from_entry(service, i)
                    if tenant_email:
                        days_expired = (today - hvac_date).days
                        status = f"Expired {days_expired} days ago"
                        
                        expired_entries.append({
                            'email': tenant_email,
                            'content': format_email_content(
                                email_template, 
                                tenant_name, 
                                unit_number,
                                tenant_address,
                                hvac_date
                            ),
                            'info': f"Row {i}: Unit: {unit_number} | Tenant: {tenant_name} | "
                                   f"Email: {tenant_email} | HVAC Date: {hvac_date.strftime('%m/%d/%y')} ({status})"
                        })
                    else:
                        print(f"⚠️ Skipping row {i} due to invalid email")

        # Process expired entries
        if expired_entries:
            print(f"\nFound {len(expired_entries)} HVAC entries expired within the last 45 days:")
            for entry in expired_entries:
                print(entry['info'])
                print("Email Content Preview:")
                print("==============")
                print(entry['content'])
                print("==============\n")
            
            if get_user_confirmation(len(expired_entries)):
                print("\nSending emails...")
                for entry in expired_entries:
                    send_email(entry['email'], EMAIL_SUBJECT, entry['content'])
            else:
                print("\nEmail sending cancelled by user.")
        else:
            print("No expired HVAC entries found.")

    except Exception as e:
        print(f"Error reading HVAC data: {e}")

def main():
    # Get mode selection
    mode = get_mode_selection()
    
    # Initialize the service
    service = initialize_sheets_service()
    if not service:
        print("Failed to initialize Google Sheets service")
        return

    # Get and display HVAC data
    get_hvac_data(service, mode)

if __name__ == '__main__':
    main()
