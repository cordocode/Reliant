from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from datetime import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load environment variables
load_dotenv()

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Full access needed
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'  # Updated to include range

# Add template path and email subject constants
EMAIL_SUBJECT = "Required Grease Trap Maintenance"
EMAIL_TEMPLATE = """Dear [Tenant Name],

As specified in your Lease Agreement's Rules and Regulations, you are required to maintain, monitor, and empty the grease trap associated with your restaurant or food preparation business at your own expense.

We are writing to request critical documentation to verify proper maintenance. Specifically, we need you to provide a copy of your most recent quarterly grease trap service manifest. This documentation must be:

Completed by a licensed, insured, and certified contractor
Dated within the last quarter
Clearly showing service details for the premises at [Tenant Address]

Proper grease trap maintenance is crucial to prevent serious potential issues, including:

Significant plumbing blockages
Potential restaurant operations disruptions
Risk of costly flood damage

Your last recorded service was on [Last Service Date]. Please submit the updated manifest within 10 business days of receiving this notice.

Failure to provide documentation may result in the landlord taking necessary actions as outlined in your lease agreement.

Thank you for your immediate attention to this important matter.

Sincerely,

Reliant Property Management"""

# Add email configuration
SENDER_EMAIL = 'admin@reliant-pm.com'  # Direct from .env
APP_PASSWORD = 'otftekojfvvhqxra'      # Direct from .env

def send_email(recipient_email, subject, body):
    """Send email using configured SMTP server"""
    try:
        # Create message
        message = MIMEMultipart()
        message['From'] = SENDER_EMAIL
        message['To'] = recipient_email
        message['Subject'] = subject

        # Add body
        message.attach(MIMEText(body, 'plain'))

        print(f"Attempting to send email from {SENDER_EMAIL} to {recipient_email}")
        
        # Create SMTP session with Gmail
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(message)

        print(f"Email sent successfully to {recipient_email}")
        return True

    except Exception as e:
        print(f"Failed to send email: {e}")
        print(f"Using sender email: {SENDER_EMAIL}")
        print(f"App password length: {len(APP_PASSWORD) if APP_PASSWORD else 0}")
        return False

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

def get_tenant_names():
    """Get specific tenant names from user input"""
    tenant_names = []
    print("\nEnter tenant names (press Enter twice to finish):")
    while True:
        name = input().strip()
        if not name:
            if tenant_names:  # If we have at least one name, exit
                break
            else:  # If no names yet, confirm
                confirm = input("No names entered. Do you want to proceed? (y/n): ")
                if confirm.lower() == 'y':
                    break
                continue
        tenant_names.append(name.upper())  # Store names in uppercase for case-insensitive comparison
    return tenant_names

def get_grease_trap_data(service, mode=1, tenant_names=None):
    """Fetch grease trap related data from Tenant sheet"""
    try:
        # Use EMAIL_TEMPLATE constant directly
        email_template = EMAIL_TEMPLATE

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
        
        processed_count = 0
        
        # Create a list to store all email data
        emails_to_send = []
        
        # Start from index 1 to skip header row
        for i, row in enumerate(values[1:], start=2):
            try:
                tenant_address = row[0] if len(row) > 0 else 'N/A'  # Column C
                tenant_name = row[1] if len(row) > 1 else 'N/A'     # Column D
                grease_required = row[13].strip().upper() if len(row) > 13 else 'N/A'  # Column P
                last_service = row[14] if len(row) > 14 else 'N/A'  # Column Q
                
                # Only process if grease trap is required
                if grease_required == 'YES':
                    # Check if we should process this tenant based on mode
                    if mode == 2:
                        tenant_matches = any(
                            name.upper().strip() in tenant_name.upper().strip() or 
                            tenant_name.upper().strip() in name.upper().strip()
                            for name in tenant_names
                        )
                        if not tenant_matches:
                            continue

                    processed_count += 1
                    # Get corresponding email from Entry sheet
                    tenant_email = get_email_from_entry(service, i)
                    
                    # Format email content
                    email_content = format_email_content(
                        email_template,
                        tenant_name,
                        tenant_address,
                        last_service
                    )
                    
                    if tenant_email != 'N/A':
                        emails_to_send.append({
                            'tenant_name': tenant_name,
                            'tenant_email': tenant_email,
                            'subject': EMAIL_SUBJECT,
                            'content': email_content,
                            'row': i,
                            'address': tenant_address,
                            'last_service': last_service
                        })
                    else:
                        print(f"No valid email found for {tenant_name}")
                    
            except IndexError:
                print(f"Skipping row {i} - incomplete data")
                continue

        # If we have emails to send, show the batch preview
        if emails_to_send:
            print("\nEmail Batch Preview:")
            print("===================")
            for idx, email_data in enumerate(emails_to_send, 1):
                print(f"\nEmail {idx} of {len(emails_to_send)}:")
                print(f"To: {email_data['tenant_name']} <{email_data['tenant_email']}>")
                print(f"Subject: {email_data['subject']}")
                print(f"Property: {email_data['address']}")
                print(f"Last Service: {email_data['last_service']}")
                print("\nContent:")
                print("--------")
                print(email_data['content'])
                print("--------\n")
            
            # Batch confirmation
            while True:
                confirm = input(f"\nSend all {len(emails_to_send)} emails? (y/n): ").lower()
                if confirm in ['y', 'n']:
                    break
                print("Please enter 'y' or 'n'")
            
            if confirm == 'y':
                successful = 0
                failed = 0
                for email_data in emails_to_send:
                    if send_email(email_data['tenant_email'], 
                                email_data['subject'], 
                                email_data['content']):
                        successful += 1
                    else:
                        failed += 1
                
                print(f"\nEmail Batch Results:")
                print(f"Successfully sent: {successful}")
                print(f"Failed to send: {failed}")
            else:
                print("\nEmail batch cancelled")
                
        elif processed_count == 0:
            if mode == 2:
                print("\nNo matching tenants found with grease trap requirements.")
            else:
                print("\nNo tenants found with grease trap requirements.")

    except Exception as e:
        print(f"Error reading grease trap data: {e}")

def main():
    # Initialize the service
    service = initialize_sheets_service()
    if not service:
        print("Failed to initialize Google Sheets service")
        return

    # Mode selection
    while True:
        print("\nSelect mode:")
        print("1. Process all tenants")
        print("2. Process specific tenants")
        try:
            mode = int(input("Enter mode (1 or 2): "))
            if mode in [1, 2]:
                break
            print("Invalid mode. Please enter 1 or 2.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    tenant_names = None
    if mode == 2:
        tenant_names = get_tenant_names()
        
    # Get and display grease trap data
    get_grease_trap_data(service, mode, tenant_names)

if __name__ == '__main__':
    main()
