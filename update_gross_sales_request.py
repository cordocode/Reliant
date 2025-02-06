from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import os

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Full access needed
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
SHEET_NAME = 'TENANT!A1:Z'  # Updated to include range
# Update template path to correct file
TEMPLATE_PATH = '/Users/cordo/Documents/RELIANT_SCRIPTS/gross_sales_template.txt'
TEMPLATE_PATH_PAST_DUE = '/Users/cordo/Documents/RELIANT_SCRIPTS/gross_sales_template_past_due.txt'

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

def parse_sales_requirement(requirement):
    """Parse the sales requirement string to get requirement and frequency"""
    if not requirement or requirement.upper() == 'NO':
        return False, None
    
    parts = requirement.split()
    if len(parts) >= 2 and parts[0].upper() == 'YES':
        return True, ' '.join(parts[1:])  # Join remaining parts as frequency
    return False, None

def read_email_template(is_past_due=False):
    """Read the appropriate email template file"""
    try:
        template_path = TEMPLATE_PATH_PAST_DUE if is_past_due else TEMPLATE_PATH
        with open(template_path, 'r') as file:
            return file.read()
    except Exception as e:
        print(f"Error reading template file: {e}")
        return None

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

def format_email_content(template, tenant_name, tenant_address, frequency, last_report, cycles):
    """Format the email template with tenant's information"""
    try:
        formatted_date = format_date_for_display(last_report)
        formatted_freq = format_frequency(frequency)  # For email body
        
        # Replace template variables
        content = template.replace("[Tenant Name]", tenant_name)
        content = content.replace("[Frequency]", formatted_freq)
        content = content.replace("[Property Address]", tenant_address)
        content = content.replace("[Last Submitted Date]", formatted_date)  # Fixed variable name
        content = content.replace("[Cycles]", str(cycles))
        
        return content
    except Exception as e:
        print(f"Error formatting email: {e}")
        return None

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

def get_gross_sales_data(service):
    """Fetch gross sales related data from Tenant sheet"""
    try:
        # Get columns C, D, K, L, M from Tenant sheet
        range_name = 'TENANT!C:M'
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
        
        # Start from index 1 to skip header row
        for i, row in enumerate(values[1:], start=2):
            try:
                tenant_address = row[0] if len(row) > 0 else 'N/A'     # Column C
                tenant_name = row[1] if len(row) > 1 else 'N/A'        # Column D
                sales_requirement = row[8] if len(row) > 8 else 'N/A'  # Column K
                start_date_str = row[9] if len(row) > 9 else 'N/A'     # Column L
                last_record_str = row[10] if len(row) > 10 else 'N/A'  # Column M
                
                # Parse sales requirement
                is_required, frequency = parse_sales_requirement(sales_requirement)
                
                if is_required:
                    # Convert dates
                    start_date = convert_date_string(start_date_str)
                    last_record_date = convert_date_string(last_record_str)
                    
                    # Calculate missed cycles
                    cycles = calculate_cycles(frequency, start_date, last_record_date, today)
                    
                    # Get appropriate template based on cycles
                    email_template = read_email_template(is_past_due=(cycles > 0))
                    if not email_template:
                        print("Failed to read email template")
                        continue
                    
                    tenant_email = get_email_from_entry(service, i)
                    
                    # Format email content
                    email_content = format_email_content(
                        email_template,
                        tenant_name,
                        tenant_address,
                        frequency,
                        last_record_str,
                        cycles  # Add cycles to output
                    )
                    
                    print(f"\nRow {i}:")
                    print(f"Tenant: {tenant_name}")
                    print(f"Address: {tenant_address}")
                    print(f"Email: {tenant_email}")
                    print(f"Reporting Frequency: {frequency}")
                    print(f"Start Date: {start_date_str}")
                    print(f"Last Record Date: {last_record_str}")
                    print(f"Missed Cycles: {cycles}")
                    print(f"Template Used: {'Past Due' if cycles > 0 else 'Standard'}")
                    print("\nEmail Content Preview:")
                    print("==============")
                    print(email_content)
                    print("==============\n")
                    
            except IndexError:
                print(f"Skipping row {i} - incomplete data")
                continue

    except Exception as e:
        print(f"Error reading gross sales data: {e}")

def main():
    # Initialize the service
    service = initialize_sheets_service()
    if not service:
        print("Failed to initialize Google Sheets service")
        return

    # Get and display gross sales data
    get_gross_sales_data(service)

if __name__ == '__main__':
    main()
