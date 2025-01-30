import os
import sys
import logging
import time
from dotenv import load_dotenv
from openai import OpenAI
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import html2text

# Set up logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
    attachment_info = []
    
    try:
        if msg.is_multipart():
            parts = [part for part in msg.walk()]
            logger.info(f"Total message parts: {len(parts)}")
            
            for part in parts:
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                logger.info(f"Found part: {content_type} - Disposition: {content_disposition}")
                
                # More aggressive PDF detection
                is_pdf = (content_type == 'application/pdf' or 
                         content_type == 'application/x-pdf' or 
                         (content_disposition and '.pdf' in content_disposition.lower()))
                
                if is_pdf:
                    filename = part.get_filename()
                    if not filename and '.pdf' in content_disposition.lower():
                        filename = f"attachment_{len(saved_files)}.pdf"
                        
                    if filename:
                        try:
                            payload = part.get_payload(decode=True)
                            if payload is None:
                                logger.error(f"Payload is None for file: {filename}")
                                attachment_info.append(f"Failed - {filename}: No payload (possible Gmail block)")
                                continue
                                
                            filepath = os.path.join(COIS_FOLDER, filename)
                            os.makedirs(COIS_FOLDER, exist_ok=True)
                            
                            with open(filepath, 'wb') as f:
                                f.write(payload)
                            saved_files.append(filepath)
                            attachment_info.append(f"Success - {filename}")
                            logger.info(f"Successfully saved PDF: {filename}")
                            
                        except Exception as e:
                            logger.error(f"Failed to save {filename}: {str(e)}")
                            attachment_info.append(f"Failed - {filename}: {str(e)}")
                    else:
                        logger.warning("PDF attachment found but no filename")
                        attachment_info.append("Failed - PDF without filename")
        
        logger.info("Attachment processing summary:")
        for info in attachment_info:
            logger.info(f"  • {info}")
            
    except Exception as e:
        logger.error(f"Error processing attachments: {str(e)}")
    
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

def fetch_entire_thread(imap, email_id):
    """
    Fetch the entire thread by using the References or In-Reply-To headers.
    Returns a list of message bodies (email.message.Message objects) for the thread.
    """
    _, raw_data = imap.fetch(email_id, "(RFC822)")
    if not raw_data or not raw_data[0]:
        logger.error(f"Could not fetch base email {email_id}")
        return []
    base_msg = email.message_from_bytes(raw_data[0][1])
    
    ref_ids = []
    references = base_msg.get("References")
    in_reply_to = base_msg.get("In-Reply-To")
    if references:
        # Typically a space-separated list of message IDs
        ref_ids = references.split()
    elif in_reply_to:
        ref_ids = [in_reply_to]

    # Always include our base message as part of the thread
    messages_in_thread = [(email_id, base_msg)]
    
    # For each reference, fetch that message if it still exists
    for ref_id in ref_ids:
        ref_id = ref_id.strip("<>")
        result, search_data = imap.search(None, f'HEADER Message-ID "{ref_id}"')
        if result == "OK" and search_data[0]:
            for mid in search_data[0].split():
                _, ref_msg_data = imap.fetch(mid, "(RFC822)")
                if ref_msg_data and ref_msg_data[0]:
                    ref_msg = email.message_from_bytes(ref_msg_data[0][1])
                    messages_in_thread.append((mid, ref_msg))

    return messages_in_thread

def process_thread(imap, email_id):
    """
    Process an entire thread as one combined unit.
    """
    thread_messages = fetch_entire_thread(imap, email_id)
    if not thread_messages:
        return None
    
    # Concatenate bodies, gather attachments
    combined_body = []
    has_coi_attachments = []
    
    for (mid, msg) in thread_messages:
        try:
            body_part = extract_email_content(msg) or ""
            combined_body.append(body_part)
        except Exception as e:
            logger.error(f"Body extraction error in thread: {e}")

    full_body = "\n\n".join(combined_body)
    # Classify once using the base message subject
    _, base_msg = thread_messages[0]
    subject = decode_email_subject(base_msg["subject"]) or ""
    classification = None

    # Attempt classification
    for attempt in range(3):
        try:
            classification = classify_email_content(subject, full_body)
            logger.info(f"Thread classification: {classification}")
            break
        except Exception as e:
            logger.error(f"Thread classification attempt {attempt + 1} failed: {e}")
            time.sleep(1)
            if attempt == 2:
                return None

    # If COI, gather PDFs from all messages
    if classification == "Updated Certificate of Insurance":
        for (mid, msg) in thread_messages:
            saved_files = save_pdf_attachment(msg, msg.get("From") or "")
            if saved_files:
                has_coi_attachments.extend(saved_files)
        if has_coi_attachments:
            logger.info(f"Thread PDFs saved: {has_coi_attachments}")

    return classification

def process_single_email(imap, email_id):
    """
    Now replaced by thread-level processing, but we keep minimal logic to unify calls.
    """
    logger.info(f"\nProcessing entire thread for email ID: {email_id}")
    classification = process_thread(imap, email_id)
    return classification

def process_recent_emails(num_emails=None):  # Changed default to None
    """Process emails with improved error handling"""
    imap = connect_to_gmail()
    if not imap:
        logger.error("Failed to connect to Gmail")
        return
    
    try:
        imap.select("INBOX")
        _, messages = imap.search(None, "ALL")
        email_ids = messages[0].split()
        
        # Only slice if num_emails is specified
        if num_emails:
            email_ids = email_ids[-num_emails:]
        email_ids.reverse()  # Process newest to oldest
        
        logger.info(f"Found {len(email_ids)} emails to process")
        
        success_count = 0
        failure_count = 0
        
        processed_results = []

        for email_id in email_ids:
            try:
                # Verify email still exists before processing
                imap.select("INBOX")
                _, check = imap.fetch(email_id, "(FLAGS)")
                if check[0] is None:
                    logger.warning(f"Email {email_id} no longer exists, skipping")
                    continue
                
                classification = process_single_email(imap, email_id)
                if classification:
                    processed_results.append((email_id, classification))
                    success_count += 1
                else:
                    failure_count += 1
                time.sleep(0.5)  # Add small delay between processing
                
            except imaplib.IMAP4.error as e:
                logger.error(f"IMAP error for email {email_id}: {e}")
                failure_count += 1
                continue

        # After we finish processing attachments/body/classification, move emails
        for email_id, classification in processed_results:
            label = GMAIL_LABELS.get(classification, GMAIL_LABELS["Other"])
            if move_email_to_label(imap, email_id, label):
                logger.info(f"Moved email {email_id} to label: {label}")
            else:
                logger.error(f"Failed to move email {email_id} to label: {label}")

        logger.info(f"Processing complete. Successes: {success_count}, Failures: {failure_count}")
        
    except Exception as e:
        logger.error(f"Error in main processing loop: {e}")
    finally:
        try:
            imap.close()
            imap.logout()
        except:
            pass

if __name__ == "__main__":
    # Only use num_emails if specifically provided as argument
    num_emails = int(sys.argv[1]) if len(sys.argv) > 1 else None
    process_recent_emails(num_emails)
