import os
import sys
from dotenv import load_dotenv
from openai import OpenAI
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import html2text

# Load environment variables
load_dotenv()
EMAIL = os.getenv('sender_email')
PASSWORD = os.getenv('app_password')

# Define email categories
EMAIL_CATEGORIES = [
    "Maintenance Request",
    "Updated Certificate of Insurance",
    "Updated Grease Trap Record",
    "Gross Sales Report",
    "Invoice to be Paid",
    "Tenant HVAC Records"
]

# Add new constant for COI downloads
COIS_FOLDER = os.path.join(os.path.expanduser('~'), 'Downloads', 'COIS')

# Add new constant for Gmail labels
GMAIL_LABELS = {
    "Updated Certificate of Insurance": "COIS",
    "Maintenance Request": "Processed/Maintenance",
    "Gross Sales Report": "Processed/Sales",
    "Invoice to be Paid": "Processed/Invoices",
    "Tenant HVAC Records": "Processed/HVAC",
    "Other": "Processed/Other"
}

def save_pdf_attachment(msg, sender):
    """Save PDF attachments from email with original filename"""
    saved_files = []
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'application/pdf':
                filename = part.get_filename()
                if filename:
                    filepath = os.path.join(COIS_FOLDER, filename)
                    os.makedirs(COIS_FOLDER, exist_ok=True)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    saved_files.append(filepath)
    
    return saved_files

def move_email_to_label(imap, email_id, label):
    """Move email to specified Gmail label"""
    try:
        try:
            imap.create(label)
        except:
            pass
        imap.copy(email_id, label)
        imap.store(email_id, '+FLAGS', '\\Deleted')
        imap.expunge()
        return True
    except Exception as e:
        return False

def handle_coi_email(msg, subject, body, imap, email_id):
    """Handle Certificate of Insurance related email"""
    print("\nCOI Email Detected - Processing...")
    
    files_saved = False
    if msg.is_multipart():
        saved_files = save_pdf_attachment(msg, msg['from'])
        if saved_files:
            files_saved = True
            print("\nCOI Processing Results:")
            print("- Type: New COI Submission")
            print("- Action: Saved PDF(s)")
            for file in saved_files:
                print(f"  • {os.path.basename(file)}")
        else:
            print("\nCOI Processing Results:")
            print("- Type: COI Related Communication")
            print("- Action: No PDFs found to save")
            print("- Content Summary:")
            print(f"  • Subject: {subject}")
            print(f"  • Message: {body[:200]}..." if len(body) > 200 else f"  • Message: {body}")
    
    # Move email to appropriate label
    if files_saved:
        move_email_to_label(imap, email_id, GMAIL_LABELS["Updated Certificate of Insurance"])

def get_openai_client():
    """Initialize and return OpenAI client"""
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY not found in environment variables")
    return OpenAI(api_key=api_key)

def connect_to_gmail():
    """Connect to Gmail IMAP server"""
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(EMAIL, PASSWORD)
        return imap
    except Exception as e:
        print(f"Error connecting to Gmail: {e}")
        return None

def decode_email_subject(subject):
    """Decode email subject"""
    if subject is None:
        return ""
    decoded_parts = decode_header(subject)
    decoded_subject = ""
    
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                decoded_subject += part.decode(encoding if encoding else 'utf-8')
            except:
                decoded_subject += part.decode('utf-8', 'ignore')
        else:
            decoded_subject += str(part)
    
    return decoded_subject

def extract_email_content(msg):
    """Extract the body from email message"""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                try:
                    return part.get_payload(decode=True).decode()
                except:
                    return part.get_payload(decode=True).decode('utf-8', 'ignore')
            elif content_type == "text/html":
                try:
                    html_content = part.get_payload(decode=True).decode()
                    # Convert HTML to plain text
                    soup = BeautifulSoup(html_content, 'html.parser')
                    return soup.get_text()
                except:
                    return "Could not decode email content"
    else:
        try:
            return msg.get_payload(decode=True).decode()
        except:
            return msg.get_payload(decode=True).decode('utf-8', 'ignore')

def classify_email_content(subject, body):
    """Classify email content using OpenAI"""
    try:
        client = get_openai_client()
        combined_text = f"Subject: {subject}\n\nBody: {body}"
        prompt = f"""Analyze this email and classify it into exactly one of these categories:
        {', '.join(EMAIL_CATEGORIES)}
        
        Respond ONLY with the category name. If none match, respond with 'Other'.
        
        Email content:
        {combined_text[:1000]}"""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a precise email classifier for property management. Respond only with the exact category name."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=20
        )
        
        return response.choices[0].message.content.strip()
    except Exception as e:
        return "Classification Error"

def process_recent_emails(num_emails=1):
    """Process all emails in inbox"""
    imap = connect_to_gmail()
    if not imap:
        return
    
    try:
        imap.select("INBOX")
        _, messages = imap.search(None, "ALL")
        email_ids = messages[0].split()
        
        if not email_ids:
            print("No emails found in inbox")
            return None
        
        # Initialize counters
        total_emails = len(email_ids)
        files_downloaded = 0
        label_counts = {label: 0 for label in set(GMAIL_LABELS.values())}
        processed_emails = []
        
        print(f"\nProcessing {total_emails} emails...")
        
        for email_id in email_ids:
            try:
                _, msg_data = imap.fetch(email_id, "(RFC822)")
                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)
                
                subject = decode_email_subject(msg["subject"])
                sender = msg["from"]
                body = extract_email_content(msg)
                classification = classify_email_content(subject, body)
                
                if classification == "Updated Certificate of Insurance":
                    saved_files = save_pdf_attachment(msg, msg['from'])
                    files_downloaded += len(saved_files)
                    if saved_files:
                        if move_email_to_label(imap, email_id, GMAIL_LABELS[classification]):
                            label_counts[GMAIL_LABELS[classification]] += 1
                else:
                    if classification in GMAIL_LABELS:
                        label_name = GMAIL_LABELS[classification]
                    else:
                        label_name = GMAIL_LABELS["Other"]
                    
                    if move_email_to_label(imap, email_id, label_name):
                        label_counts[label_name] += 1
                
                processed_emails.append({
                    "sender": sender,
                    "subject": subject,
                    "body": body,
                    "category": classification
                })
                
            except Exception as e:
                continue
        
        # Print final statistics
        print("\nProcessing Complete!")
        print(f"Total emails processed: {total_emails}")
        print(f"Total PDFs downloaded: {files_downloaded}")
        print("\nEmails moved to labels:")
        for label, count in label_counts.items():
            if count > 0:
                print(f"- {label}: {count}")
        
        imap.close()
        imap.logout()
        return processed_emails
        
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    process_recent_emails()
