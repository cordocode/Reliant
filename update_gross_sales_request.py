from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load environment variables
load_dotenv()

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Full access needed
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'  # Updated to include range

# Add template constant
GROSS_SALES_TEMPLATE = """[Tenant Name],

We're missing [Cycles] [Frequency] gross sales reports for [Property Address].

Please send the outstanding reports, signed by an officer, to bring your account current.

Best regards,

Reliant"""

# Email Configuration
SENDER_EMAIL = os.getenv('sender_email')
APP_PASSWORD = os.getenv('app_password')

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
        # Add +2 to adjust for the offset between TENANT and ENTRY sheets
        # This ensures that row 58 on TENANT sheet corresponds to row 60 on ENTRY sheet
        range_name = f'Entry!J{row_number + 2}'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        return values[0][0] if values and values[0] else 'N/A'
    except Exception as e:
        print(f"Error getting email from Entry sheet: {e}")
        return 'N/A'

def parse_sales_requirement(requirement):
    """Parse the sales requirement string to get requirement and frequency"""
    if not requirement:
        return False, None
        
    # Convert to uppercase for comparison but keep original for frequency
    requirement_upper = requirement.upper()
    if requirement_upper == 'NO':
        return False, None
    
    # Handle both YES and yes cases
    if 'YES' in requirement_upper:
        # Extract frequency by removing 'yes' and any parentheses
        frequency = requirement.replace('yes', '').replace('Yes', '').replace('YES', '')
        frequency = frequency.strip('() \t\n\r')
        if frequency:
            return True, frequency
            
    return False, None

def format_date_for_display(date_str):
    """Convert date string to readable format"""
    try:
        # Convert MMDDYY to datetime
        date_obj = datetime.strptime(date_str, '%m%d%y')
        return date_obj.strftime('%B %d, %Y')
    except ValueError:
        return date_str if date_str else 'N/A'

def format_frequency(frequency, for_subject=False):
    """Format frequency text to be more natural
    Args:
        frequency: Raw frequency text like "(annual)"
        for_subject: If True, capitalize first letter and remove parentheses
    """
    try:
        # Remove parentheses and whitespace
        clean_freq = frequency.strip('() \t\n\r')
        
        if for_subject:
            # Capitalize first letter for subject line
            return clean_freq.capitalize()
        else:
            # Lowercase for email body
            return clean_freq.lower()
    except Exception:
        return frequency

def format_email_content(tenant_name, property_address, cycles, frequency):
    """Format the email template with tenant's information"""
    try:
        # Ensure all parameters are strings to prevent NoneType errors
        tenant_name = str(tenant_name) if tenant_name else "Tenant"
        property_address = str(property_address) if property_address else "your property"
        cycles = str(cycles) if cycles else "missing"
        frequency = str(frequency) if frequency else "regular"
        
        # Use the constant template directly instead of reading from file
        content = GROSS_SALES_TEMPLATE.replace("[Tenant Name]", tenant_name)
        content = content.replace("[Property Address]", property_address)
        content = content.replace("[Cycles]", cycles)
        content = content.replace("[Frequency]", frequency)
        return content
    except Exception as e:
        print(f"Error formatting email: {e}")
        return "We're missing gross sales reports for your property. Please send the outstanding reports to bring your account current."

def calculate_cycles(frequency, start_date, last_record_date, today_date):
    """Calculate number of missed reporting cycles
    Args:
        frequency: String indicating reporting frequency (monthly, quarterly, etc.)
        start_date: Initial date of reporting requirement (datetime)
        last_record_date: Date of last received record (datetime)
        today_date: Current date (datetime)
    Returns:
        Integer representing number of missed cycles
    """
    try:
        # If no last record, calculate from start date
        reference_date = last_record_date if last_record_date else start_date
        
        # Define period lengths for different frequencies
        frequency_periods = {
            'monthly': relativedelta(months=1),
            'quarterly': relativedelta(months=3),
            'bi-annual': relativedelta(months=6),
            'annual': relativedelta(years=1)
        }
        
        # Clean frequency string and get corresponding period
        clean_freq = frequency.strip('() \t\n\r').lower()
        period = frequency_periods.get(clean_freq)
        
        if not period:
            print(f"Unknown frequency format: {frequency}")
            return 0
            
        # Calculate next due date from reference date
        next_due = reference_date + period
        
        # If next due date is in the future, no cycles missed
        if next_due > today_date:
            return 0
            
        # Calculate number of periods between next due and today
        cycles = 0
        while next_due <= today_date:
            cycles += 1
            next_due += period
            
        return cycles
        
    except Exception as e:
        print(f"Error calculating cycles: {e}")
        return 0

def convert_date_string(date_str):
    """Convert 6-digit date string (MMDDYY) to datetime object"""
    try:
        return datetime.strptime(date_str, '%m%d%y').date()
    except ValueError:
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

def send_email(recipient_email, subject, body):
    """Send email using configured SMTP server"""
    try:
        # Validate email content
        if not recipient_email:
            print("Missing recipient email")
            return False
        
        if not body:
            print("Missing email body")
            return False
            
        # Ensure subject is a string
        if not subject:
            subject = "Gross Sales Reports"
        
        # Create message
        message = MIMEMultipart()
        message['From'] = SENDER_EMAIL
        message['To'] = recipient_email
        message['Subject'] = subject

        # Add body
        message.attach(MIMEText(body, 'plain'))

        # Create SMTP session with Gmail
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(message)

        print(f"Email sent successfully to {recipient_email}")
        return True

    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

def is_report_due(frequency, last_record_date, today_date):
    """
    Determine if a report is currently due based on frequency and last record date
    Returns: (bool, datetime) - (is_due, next_due_date)
    """
    try:
        if not last_record_date:
            return True, today_date
            
        # Define period lengths and grace period (5 days)
        frequency_periods = {
            'monthly': relativedelta(months=1),
            'quarterly': relativedelta(months=3),
            'bi-annual': relativedelta(months=6),
            'annual': relativedelta(years=1)
        }
        
        # Clean frequency string and get period
        clean_freq = frequency.strip('() \t\n\r').lower()
        period = frequency_periods.get(clean_freq)
        
        if not period:
            print(f"Unknown frequency format: {frequency}")
            return False, None
            
        # Calculate next due date from last record
        next_due = last_record_date + period
        
        # Add 5 days grace period
        grace_period = next_due + relativedelta(days=5)
        
        # Report is due if we're past the grace period
        return today_date > grace_period, next_due
        
    except Exception as e:
        print(f"Error checking if report is due: {e}")
        return False, None

def get_gross_sales_data(service, mode=1, tenant_names=None):
    """Fetch gross sales related data from Tenant sheet based on mode"""
    try:
        range_name = 'TENANT!C1:M999'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print('No data found.')
            return
        
        print("\nGross Sales Reporting Records:")
        print("============================")
        
        today = date.today()
        processed_count = 0
        
        # Create a list to store all email data
        emails_to_send = []
        
        # Start from index 1 to skip header row
        for i, row in enumerate(values[1:], start=2):
            try:
                if len(row) < 11:  # Check if row has enough columns
                    continue
                    
                tenant_address = row[0] if len(row) > 0 else 'N/A'
                tenant_name = row[1] if len(row) > 1 else 'N/A'
                sales_requirement = row[8] if len(row) > 8 else 'N/A'
                start_date_str = row[9] if len(row) > 9 else 'N/A'
                last_record_str = row[10] if len(row) > 10 else 'N/A'
                
                # First check if gross sales reporting is required
                is_required, frequency = parse_sales_requirement(sales_requirement)
                
                if not is_required:
                    continue
                
                # Then check if we should process this tenant based on mode
                if mode == 2:
                    tenant_matches = any(
                        name.upper().strip() in tenant_name.upper().strip() or 
                        tenant_name.upper().strip() in name.upper().strip()
                        for name in tenant_names
                    )
                    if not tenant_matches:
                        continue

                processed_count += 1
                # Convert dates
                start_date = convert_date_string(start_date_str)
                last_record_date = convert_date_string(last_record_str)
                
                # Check if report is actually due and has missed cycles
                is_due, next_due = is_report_due(frequency, last_record_date, today)
                cycles = calculate_cycles(frequency, start_date, last_record_date, today)
                
                # Only proceed if due and has missed cycles
                if not is_due or cycles == 0:
                    print(f"Skipping {tenant_name} - report not due yet (Next due: {next_due.strftime('%B %d, %Y')})")
                    continue
                
                tenant_email = get_email_from_entry(service, i)
                
                # Validate cycles before formatting
                cycles_str = str(cycles) if cycles else "1"
                
                # Format email content with validation
                email_content = format_email_content(
                    tenant_name,
                    tenant_address,
                    cycles_str,
                    frequency
                )
                
                # Format email subject - always past due since we only send when cycles > 0
                subject = f"Past Due {format_frequency(frequency, for_subject=True)} Gross Sales Reports"
                
                # Instead of sending immediately, store the email data
                if tenant_email != 'N/A':
                    # Only add to emails list if we have valid content
                    if email_content:
                        emails_to_send.append({
                            'tenant_name': tenant_name,
                            'tenant_email': tenant_email,
                            'subject': subject,
                            'content': email_content,
                            'row': i,
                            'address': tenant_address,
                            'frequency': frequency,
                            'start_date': start_date_str,
                            'last_date': last_record_str,
                            'cycles': cycles,
                            'is_past_due': True  # Always true now
                        })
                    else:
                        print(f"Failed to generate email content for {tenant_name}")
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
                print(f"Frequency: {email_data['frequency']}")
                print(f"Last Report: {email_data['last_date']}")
                print(f"Status: {'Past Due' if email_data['is_past_due'] else 'Standard'}")
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
                print("\nNo matching tenants found with gross sales reporting requirements.")
            else:
                print("\nNo tenants found with gross sales reporting requirements.")

    except Exception as e:
        print(f"Error reading gross sales data: {e}")

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
        
    # Get and display gross sales data
    get_gross_sales_data(service, mode, tenant_names)

if __name__ == '__main__':
    main()
