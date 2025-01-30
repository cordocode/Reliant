import os
import pickle
import logging
import time
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from dotenv import load_dotenv
from openai import OpenAI

# Set up logging without timestamps or level
logging.basicConfig(level=logging.INFO,
                   format='%(message)s')
logger = logging.getLogger(__name__)

# Email Categories
EMAIL_CATEGORIES = {
    'COI_RECORDS': {
        'name': 'COI_RECORDS',
        'label': 'COI_RECORDS'
    },
    'GREASE_TRAP_RECORDS': {
        'name': 'GREASE_TRAP_RECORDS',
        'label': 'GREASE_TRAP_RECORDS'
    },
    'GROSS_SALES_REPORT': {
        'name': 'GROSS_SALES_REPORT',
        'label': 'GROSS_SALES_REPORT'
    },
    'HVAC_RECORDS': {
        'name': 'HVAC_RECORDS',
        'label': 'HVAC_RECORDS'
    },
    'INVOICES': {
        'name': 'INVOICES',
        'label': 'INVOICES'
    }
}

# Add Thread Status constants
THREAD_STATUS = {
    'COMPLETED': 'Thread contains final response or confirmation',
    'URGENT_ACTION': 'Thread requires immediate attention or action',
    'ACTION': 'Thread requires follow-up or response'
}

# If modifying these scopes, delete the token.pickle file.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Set up credentials directory
CREDS_DIR = os.path.join(os.path.dirname(__file__), 'credentials')
TOKEN_PATH = os.path.join(CREDS_DIR, 'token.json')
os.makedirs(CREDS_DIR, exist_ok=True)

def get_openai_client():
    """Initialize and return OpenAI client"""
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY not found in environment variables")
    return OpenAI(api_key=api_key)

def get_gmail_service():
    """Get or create Gmail API service object."""
    creds = None
    token_path = os.path.join(CREDS_DIR, 'token.json')  # Changed from pickle to json
    
    # Load existing credentials if they exist
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        
    # If no valid credentials available, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.getenv('GMAIL_OAUTH_CREDENTIALS_PATH'), SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for future runs
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    try:
        # Disable cache_discovery to avoid file_cache warning
        return build('gmail', 'v1', credentials=creds, cache_discovery=False)
    except Exception as e:
        logger.error(f"Error building Gmail service: {e}")
        return None

def get_thread_ids(service, query="in:inbox"):
    """Fetch all unique thread IDs from inbox"""
    try:
        threads = service.users().threads().list(userId='me', q=query).execute()
        thread_ids = []
        
        if 'threads' in threads:
            thread_ids.extend(thread['id'] for thread in threads['threads'])
            
        while 'nextPageToken' in threads:
            page_token = threads['nextPageToken']
            threads = service.users().threads().list(
                userId='me', q=query, pageToken=page_token).execute()
            thread_ids.extend(thread['id'] for thread in threads['threads'])
            
        return thread_ids
    except Exception as e:
        logger.error(f"Error fetching thread IDs: {e}")
        return []

def get_thread_messages(service, thread_id):
    """Get all messages in a thread"""
    try:
        thread = service.users().threads().get(userId='me', id=thread_id).execute()
        messages = thread['messages']
        
        # Sort messages by internal date (oldest first)
        messages.sort(key=lambda msg: int(msg['internalDate']))
        
        logger.info(f"Found {len(messages)} messages in thread")
        return messages
    except Exception as e:
        logger.error(f"Error fetching thread {thread_id}: {e}")
        return []

def classify_thread_content(subject, full_body):
    """Classify thread content using OpenAI with weighted subject"""
    try:
        client = get_openai_client()
        # Repeat subject 3 times to give it more weight
        weighted_text = f"Subject: {subject}\nSubject: {subject}\nSubject: {subject}\n\nBody Content:\n{full_body}"
        
        prompt = f"""Based on the following text, which category matches most closely to this email thread?
        Available categories:
        - COI_RECORDS
        - GREASE_TRAP_RECORDS
        - GROSS_SALES_REPORT
        - HVAC_RECORDS
        - INVOICES

        Respond ONLY with the exact category name that matches best.

        Email thread content:
        {weighted_text[:4000]}"""  # Increased context length

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "developer", "content": "You are a precise email classifier who works in and has a deep understanding of the commercial property management world. Respond only with the exact category name from the provided list."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,  # Lower temperature for more consistent results
            max_tokens=10
        )
        
        category = response.choices[0].message.content.strip()
        if category in EMAIL_CATEGORIES:
            return category
        return None
        
    except Exception as e:
        logger.error(f"Classification error: {e}")
        return None

def determine_thread_status(subject, full_body):
    """Determine the status/urgency of the thread using OpenAI"""
    try:
        client = get_openai_client()
        
        messages = [
            {
                "role": "developer",
                "content": """You are an expert email analyzer for a property management company. 
                You understand email context and can determine when:
                - A thread is COMPLETED (all questions answered, final confirmations received)
                - A thread needs ACTION (someone is waiting for a response or follow-up)
                - A thread needs URGENT_ACTION (immediate attention required, time-sensitive requests)
                Respond only with one of these three status codes."""
            },
            {
                "role": "assistant",
                "content": """I will analyze the email thread and respond with:
                COMPLETED - when all matters are resolved
                ACTION - when a response or follow-up is needed
                URGENT_ACTION - when immediate attention is required"""
            },
            {
                "role": "user",
                "content": f"""Analyze this email thread and determine its status:
                Subject: {subject}
                
                Thread Content:
                {full_body[:4000]}"""
            }
        ]

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.3,
            max_tokens=10
        )
        
        status = response.choices[0].message.content.strip()
        return status if status in THREAD_STATUS else 'ACTION'
        
    except Exception as e:
        logger.error(f"Status determination error: {e}")
        return 'ACTION'  # Default to requiring action if there's an error

def get_email_from_headers(headers):
    """Extract email address from message headers"""
    from_header = next((h['value'] for h in headers if h['name'].lower() == 'from'), '')
    # Extract email address from "Name <email@domain.com>" format
    if '<' in from_header and '>' in from_header:
        return from_header.split('<')[1].split('>')[0]
    return from_header

# Change COIS_FOLDER to point directly to Downloads
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser('~'), 'Downloads')

def get_unique_filepath(filepath):
    """Generate a unique filepath by adding (1), (2), etc. if file exists"""
    if not os.path.exists(filepath):
        return filepath
    
    base, ext = os.path.splitext(filepath)
    counter = 1
    while True:
        new_filepath = f"{base} ({counter}){ext}"
        if not os.path.exists(new_filepath):
            return new_filepath
        counter += 1

def save_pdf_attachment(message, sender, service):
    """Download attachments exactly as they appear in Gmail"""
    saved_files = []
    
    try:
        for part in message['payload'].get('parts', []):
            if part.get('filename') and part.get('body', {}).get('attachmentId'):
                attachment = service.users().messages().attachments().get(
                    userId='me',
                    messageId=message['id'],
                    id=part['body']['attachmentId']
                ).execute()
                
                file_data = attachment['data']
                import base64
                file_bytes = base64.urlsafe_b64decode(file_data)
                
                # Get unique filepath for this download
                initial_filepath = os.path.join(DOWNLOADS_FOLDER, part['filename'])
                filepath = get_unique_filepath(initial_filepath)
                
                with open(filepath, 'wb') as f:
                    f.write(file_bytes)
                saved_files.append(filepath)
                logger.info(f"Downloaded: {os.path.basename(filepath)}")
    
    except Exception as e:
        logger.error(f"Error downloading attachment: {e}")
    
    return saved_files

def get_or_create_label(service, label_name):
    """Get existing label ID or create new label and return its ID"""
    try:
        # First try to find existing label
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        
        # Look for existing label
        for label in labels:
            if label['name'] == label_name:
                return label['id']
        
        # If not found, create new label
        label_object = {
            'name': label_name,
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show'
        }
        created_label = service.users().labels().create(
            userId='me',
            body=label_object
        ).execute()
        return created_label['id']
        
    except Exception as e:
        logger.error(f"Error managing label {label_name}: {e}")
        return None

def move_email_to_label(service, thread_id, label_name):
    """Remove email from inbox only"""
    try:
        # Get INBOX label ID
        labels_response = service.users().labels().list(userId='me').execute()
        inbox_id = next(label['id'] for label in labels_response['labels'] if label['name'] == 'INBOX')

        # Only remove from inbox
        service.users().threads().modify(
            userId='me',
            id=thread_id,
            body={
                'removeLabelIds': [inbox_id]
            }
        ).execute()
        
        logger.info(f"Removed from inbox")
        return True
        
    except Exception as e:
        logger.error(f"Error removing from inbox: {e}")
        return False

def process_thread(service, thread_id):
    """Process an entire email thread as one unit"""
    try:
        messages = get_thread_messages(service, thread_id)
        if not messages:
            return False
            
        # Get thread subject and sender info
        first_message = messages[0]
        last_message = messages[-1]
        headers = first_message['payload']['headers']
        last_headers = last_message['payload']['headers']
        
        subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
        last_sender = get_email_from_headers(last_headers)
        
        logger.info(f"\nProcessing thread: {subject}")
        logger.info(f"Most recent sender: {last_sender}")
        logger.info(f"Thread ID: {thread_id}")
        logger.info(f"Number of messages: {len(messages)}")
        
        # Combine all message bodies
        thread_text = []
        for message in messages:
            msg_body = get_message_body(message)
            if msg_body:
                thread_text.append(msg_body)
        
        combined_text = "\n---\n".join(thread_text)
        
        # First determine category and status
        category = classify_thread_content(subject, combined_text)
        if not category:
            return False
            
        status = determine_thread_status(subject, combined_text)
        logger.info(f"Thread classified as: {category}")
        logger.info(f"Thread status: {status}")
        
        # Update success/failure logic
        if status == 'URGENT_ACTION':
            logger.info("⚠️  URGENT ACTION REQUIRED  ⚠️")
            return True  # Successfully identified as urgent
        
        if status == 'ACTION':
            logger.info("Thread requires action - Keeping in inbox")
            return True  # Successfully identified as needing action
            
        # Only proceed with file download and label moving if status is COMPLETED
        if status == 'COMPLETED':
            logger.info("Thread is complete - Processing for archival")
            
            success = True
            # Download PDFs if present
            pdf_files = []
            for message in messages:
                if 'payload' in message:
                    downloaded = save_pdf_attachment(
                        message, 
                        get_email_from_headers(message['payload']['headers']),
                        service  # Pass the service object
                    )
                    pdf_files.extend(downloaded)
            
            if pdf_files:
                logger.info(f"Downloaded {len(pdf_files)} PDF files")
            
            # Move to appropriate label
            label = EMAIL_CATEGORIES[category]['label']
            if not move_email_to_label(service, thread_id, label):  # Pass the service object
                logger.error("Failed to move thread to appropriate label")
                success = False
            else:
                logger.info(f"Thread archived under: {label}")
            
            return success
                
        return True  # Default success for other cases
        
    except Exception as e:
        logger.error(f"Error processing thread {thread_id}: {e}")
        return False

def get_message_body(message):
    """Extract message body from Gmail API message format"""
    try:
        if 'payload' not in message:
            return ""
            
        parts = []
        if 'body' in message['payload']:
            if 'data' in message['payload']['body']:
                parts.append(message['payload']['body']['data'])
        
        if 'parts' in message['payload']:
            for part in message['payload']['parts']:
                if 'data' in part.get('body', {}):
                    parts.append(part['body']['data'])
        
        return "\n".join(parts)
        
    except Exception as e:
        logger.error(f"Error extracting message body: {e}")
        return ""

def get_user_preference():
    """Get user preference for test mode"""
    while True:
        response = input("\nRun in test mode? (process only first thread) [y/n]: ").lower()
        if response in ['y', 'yes']:
            return True
        elif response in ['n', 'no']:
            return False
        print("Please enter 'y' or 'n'")

def process_inbox(test_mode=False):
    """Process all threads in inbox or just first thread in test mode"""
    service = get_gmail_service()
    if not service:
        logger.error("Failed to connect to Gmail API")
        return
    
    try:
        thread_ids = get_thread_ids(service)
        total_threads = len(thread_ids)
        
        if test_mode:
            logger.info("TEST MODE: Processing only first thread")
            thread_ids = thread_ids[:1]
        
        logger.info(f"Found {total_threads} total threads")
        logger.info(f"Will process {len(thread_ids)} threads")
        
        success_count = 0
        failure_count = 0
        
        for thread_id in thread_ids:
            if process_thread(service, thread_id):
                success_count += 1
            else:
                failure_count += 1
            time.sleep(0.5)
            
        logger.info(f"\nProcessing complete:")
        logger.info(f"Threads processed successfully: {success_count}")
        logger.info(f"Threads failed: {failure_count}")
        if test_mode:
            logger.info(f"Remaining unprocessed threads: {total_threads - 1}")
        
    except Exception as e:
        logger.error(f"Error in main processing loop: {e}")

if __name__ == '__main__':
    load_dotenv()
    test_mode = get_user_preference()
    process_inbox(test_mode)

