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

# Set up logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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

def process_thread(service, thread_id):
    """Process an entire email thread as one unit"""
    try:
        messages = get_thread_messages(service, thread_id)
        if not messages:
            return False
            
        # Get thread subject from the first message
        first_message = messages[0]
        headers = first_message['payload']['headers']
        subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
        
        logger.info(f"\nProcessing thread: {subject}")
        logger.info(f"Thread ID: {thread_id}")
        logger.info(f"Number of messages: {len(messages)}")
        
        return True
    except Exception as e:
        logger.error(f"Error processing thread {thread_id}: {e}")
        return False

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

