import imaplib
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Load environment variables
load_dotenv()
EMAIL = os.getenv('sender_email')
PASSWORD = os.getenv('app_password')

# Google Sheets Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'

# Add this near the top with other constants
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser('~'), 'Downloads')

def get_vendor_details(email_address):
    """Get vendor name and expiration date from spreadsheet using email"""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        
        # Get all relevant columns
        range_name = 'VENDORS!B2:G'  # Includes vendor name (B) and expiration date (G)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        # Find row with matching email (column E, index 3)
        for row in values:
            if len(row) > 4 and row[3].lower().strip() == email_address.lower().strip():
                vendor_name = row[0]  # Column B
                exp_date = row[5] if len(row) > 5 else None  # Column G
                return vendor_name, exp_date
        return None, None
    except Exception as e:
        print(f"Error fetching vendor details: {e}")
        return None, None

def format_vendor_name(name):
    """Convert vendor name to uppercase with underscores"""
    return name.strip().upper().replace(' ', '_')

def get_next_year_expiration(exp_date):
    """Convert MMDDYY to next year's date"""
    try:
        # Parse the date
        date_obj = datetime.strptime(exp_date, '%m%d%y')
        # Add one year
        next_year = date_obj.year + 1
        # Format back to MMDDYY, ensuring year is two digits
        return date_obj.strftime(f'%m%d{str(next_year)[-2:]}')
    except Exception as e:
        print(f"Error processing expiration date: {e}")
        return exp_date

def save_pdf_attachment(part, sender_email):
    """Save PDF attachment with new naming convention"""
    try:
        filename = part.get_filename()
        if not filename:
            return None
            
        print(f"Found attachment: {filename}")
        if not filename.lower().endswith('.pdf'):
            return None
            
        # Get vendor details
        vendor_name, exp_date = get_vendor_details(sender_email)
        if not vendor_name or not exp_date:
            print(f"Could not find vendor details for {sender_email}")
            return None
            
        # Create filename
        formatted_name = format_vendor_name(vendor_name)
        next_exp = get_next_year_expiration(exp_date)
        new_filename = f"COI_{formatted_name}_{next_exp}.pdf"
        filepath = os.path.join(DOWNLOADS_FOLDER, new_filename)
        
        # Save file
        content = part.get_payload(decode=True)
        if content:
            with open(filepath, 'wb') as f:
                f.write(content)
            print(f"✓ Saved PDF: {new_filename}")
            return filepath
            
        return None
    except Exception as e:
        print(f"Error saving PDF: {e}")
        return None

def get_vendor_emails():
    """Fetch all vendor emails from Google Sheets"""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        
        # Get emails from column E
        range_name = 'VENDORS!E2:E'
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        return [email[0].lower() for email in result.get('values', []) if email]
    except Exception as e:
        print(f"Error fetching vendor emails: {e}")
        return []

def get_thread_emails(msg):
    """Extract all email addresses from message headers"""
    email_addresses = set()
    
    # Get all email headers that might contain addresses
    headers_to_check = ['From', 'To', 'Cc', 'Bcc', 'Reply-To']
    for header in headers_to_check:
        addresses = msg.get(header, '')
        if addresses:
            # Extract email addresses from the header
            parts = addresses.split(',')
            for part in parts:
                if '<' in part:
                    email_addr = part[part.find('<')+1:part.find('>')]
                    email_addresses.add(email_addr.lower())
                else:
                    email_addresses.add(part.strip().lower())
    
    return email_addresses

def process_email(msg, email_id, imap):
    """Process single email for PDFs and sort if necessary"""
    try:
        has_pdf = False
        saved_files = []
        
        # Get sender email with better error handling
        sender = msg['from']
        sender_email = None
        try:
            if '<' in sender:
                sender_email = sender[sender.find('<')+1:sender.find('>')]
            else:
                sender_email = sender.strip()
            print(f"Processing email from: {sender_email}")
        except Exception as e:
            print(f"Error extracting sender email: {e}")
            return False
        
        # More thorough attachment checking
        if msg.is_multipart():
            for part in msg.walk():
                # Debug content type
                print(f"Part content type: {part.get_content_type()}")
                if part.get_content_type().startswith('application/') or \
                   part.get_content_maintype() == 'application':
                    saved_file = save_pdf_attachment(part, sender_email)
                    if saved_file:
                        has_pdf = True
                        saved_files.append(saved_file)
        
        if has_pdf:
            try:
                # Try to create label if it doesn't exist
                try:
                    imap.create('SORTED_COI')
                except:
                    pass  # Label might already exist
                
                print("Moving email to SORTED_COI label...")
                imap.store(email_id, '+X-GM-LABELS', 'SORTED_COI')
                print("Removing from inbox...")
                imap.store(email_id, '-X-GM-LABELS', '\\Inbox')
                print(f"\nProcessed email from {sender}:")
                print(f"- Moved to SORTED_COI")
                print("- Saved PDFs:")
                for file in saved_files:
                    print(f"  * {os.path.basename(file)}")
                return True
            except Exception as e:
                print(f"Error moving email: {str(e)}")
                import traceback
                print(traceback.format_exc())
        else:
            print("No PDFs were found or successfully processed in this email")
        
        return has_pdf
        
    except Exception as e:
        print(f"Error in process_email: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return False

def scan_inbox():
    try:
        vendor_emails = get_vendor_emails()
        print(f"Loaded {len(vendor_emails)} vendor emails from spreadsheet")

        # Connect to Gmail
        imap_server = "imap.gmail.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(EMAIL, PASSWORD)
        imap.select("INBOX")
        
        _, messages = imap.search(None, 'SUBJECT "Request for updated Certificate of Insurance"')
        email_ids = messages[0].split()
        
        processed_count = 0
        total_emails = len(email_ids)
        
        print(f"\nProcessing {total_emails} matching email threads...")
        
        for email_id in email_ids:
            _, msg_data = imap.fetch(email_id, '(RFC822)')
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Get thread participants and find vendor match
            thread_emails = get_thread_emails(msg)
            matches = thread_emails.intersection(set(vendor_emails))
            
            print(f"\nAnalyzing Thread: {msg.get('Subject')}")
            print("Thread participants:", ", ".join(thread_emails))
            
            if matches:
                matching_email = next(iter(matches))
                print(f"\nFound matching vendor: {matching_email}")
                pdfs_found = False
                
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == 'application/pdf' or (
                            part.get_content_maintype() == 'application' and 
                            part.get_filename() and 
                            part.get_filename().lower().endswith('.pdf')
                        ):
                            # Use matching_email instead of sender's email
                            saved_file = save_pdf_attachment(part, matching_email)
                            if saved_file:
                                pdfs_found = True
                                processed_count += 1
                
                    if pdfs_found:
                        try:
                            # Handle Gmail labels
                            email_id_str = email_id.decode('utf-8') if isinstance(email_id, bytes) else email_id
                            imap.store(email_id_str, '+FLAGS', '\\Seen')
                            imap.store(email_id_str, '+X-GM-LABELS', 'SORTED_COIS')
                            imap.store(email_id_str, '-FLAGS', '\\Inbox')
                            imap.expunge()
                            print("✓ Email moved to SORTED_COIS")
                        except Exception as e:
                            print(f"✗ Error moving email: {e}")
            
            print("-" * 50)
        
        print(f"\nProcessing complete:")
        print(f"- Total threads checked: {total_emails}")
        print(f"- Threads with PDFs processed: {processed_count}")
        
        imap.close()
        imap.logout()
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    scan_inbox()
