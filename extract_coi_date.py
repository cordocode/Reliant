import os
from google.cloud import vision
from google.cloud.vision_v1 import types
import io
from pdf2image import convert_from_path
import tempfile
from PIL import Image
import re
from datetime import datetime, timedelta
import signal
from functools import partial
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, TimeoutError
import sys
import threading
from openai import OpenAI
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Load environment variables at the start
load_dotenv()

# Constants
COIS_FOLDER = '/Users/cordo/Downloads/COIS'
# Update to use new vision key file
VISION_CREDENTIALS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vision_key_new.json')
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sheets_key.json')  # Add this line

# Add timeout constants
PDF_CONVERSION_TIMEOUT = 30  # seconds
VISION_API_TIMEOUT = 20  # seconds

# Define specific regions for text extraction
TEXT_REGIONS = {
    'date': {
        'vertical_range': (35, 45),
        'horizontal_range': (50, 69)
    },
    'vendor': {
        'vertical_range': (20, 32),
        'horizontal_range': (0, 50)
    }
}

def init_vision_client():
    """Initialize Vision client with explicit credentials"""
    try:
        if not os.path.exists(VISION_CREDENTIALS):
            raise FileNotFoundError(f"Credentials file not found at {VISION_CREDENTIALS}")
            
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = VISION_CREDENTIALS
        client = vision.ImageAnnotatorClient()
        return client
    except Exception as e:
        print(f"✗ Error initializing Vision client: {e}")
        raise

def get_pdf_files():
    """Get all PDF files from COIS folder"""
    try:
        if not os.path.exists(COIS_FOLDER):
            print(f"Creating COIS folder at {COIS_FOLDER}")
            os.makedirs(COIS_FOLDER)
        
        pdf_files = [f for f in os.listdir(COIS_FOLDER) if f.lower().endswith('.pdf')]
        print(f"Found {len(pdf_files)} PDF files")
        return pdf_files
    except Exception as e:
        print(f"Error accessing COIS folder: {e}")
        return []

def pdf_to_images(pdf_path):
    """Convert PDF to images with timeout handling"""
    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(convert_from_path, pdf_path, first_page=1, last_page=1)
            try:
                return future.result(timeout=PDF_CONVERSION_TIMEOUT)
            except TimeoutError:
                print(f"PDF conversion timed out for {pdf_path}")
                return None
    except Exception as e:
        print(f"Error converting PDF: {e}")
        return None

def get_intersection_bounds(image, vertical_range, horizontal_range):
    """Calculate the intersection of vertical and horizontal ranges"""
    width, height = image.size
    
    # Convert percentages to pixels for the intersection area only
    top = int(height * (vertical_range[0] / 100))
    bottom = int(height * (vertical_range[1] / 100))
    left = int(width * (horizontal_range[0] / 100))
    right = int(width * (horizontal_range[1] / 100))
    
    return {
        'left': left,
        'top': top,
        'right': right,
        'bottom': bottom
    }

def extract_text_from_region(image, region_bounds):
    """Extract text from specific region of the image"""
    try:
        region = image.crop((
            region_bounds['left'],
            region_bounds['top'],
            region_bounds['right'],
            region_bounds['bottom']
        ))
        
        # Get client for each request
        client = init_vision_client()
        
        with io.BytesIO() as buffer:
            region.save(buffer, format='PNG')
            content = buffer.getvalue()
        
        vision_image = types.Image(content=content)
        response = client.text_detection(image=vision_image)
        texts = response.text_annotations
        
        if texts:
            return texts[0].description
        return ""
    except Exception as e:
        print(f"Error extracting text from region: {e}")
        return ""

def extract_dates_from_text(text):
    """Extract and parse dates from text, return the latest future date"""
    # Date patterns to match various formats
    date_patterns = [
        r'(\d{1,2})/(\d{1,2})/(\d{2,4})',  # MM/DD/YY or MM/DD/YYYY
        r'(\d{1,2})-(\d{1,2})-(\d{2,4})',  # MM-DD-YY or MM-DD-YYYY
        r'(\d{1,2}\s(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s\d{2,4})'  # 15 January 2024
    ]
    
    latest_date = None
    today = datetime.now()
    
    try:
        for pattern in date_patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                try:
                    if len(match.groups()) == 3:  # MM/DD/YY format
                        month, day, year = match.groups()
                        # Convert 2-digit year to 4-digit
                        if len(year) == 2:
                            year = '20' + year
                        date_str = f"{month}/{day}/{year}"
                        date_obj = datetime.strptime(date_str, '%m/%d/%Y')
                    else:  # Other formats
                        date_str = match.group()
                        # Try multiple date formats
                        for fmt in ['%d %B %Y', '%d %b %Y', '%m/%d/%Y', '%m-%d-%Y']:
                            try:
                                date_obj = datetime.strptime(date_str, fmt)
                                break
                            except ValueError:
                                continue
                    
                    # Only consider future dates
                    if date_obj > today:
                        if latest_date is None or date_obj > latest_date:
                            latest_date = date_obj
                except ValueError:
                    continue
    except Exception as e:
        print(f"Error parsing dates: {e}")
    
    return latest_date

def rename_file_with_date(pdf_file, result):
    """Rename the PDF file to include the extracted date and vendor in uppercase"""
    try:
        old_path = os.path.join(COIS_FOLDER, pdf_file)
        
        # Format date for filename (MMDDYY)
        date_str = datetime.strptime(result['date'], '%m/%d/%Y').strftime('%m%d%y')
        
        # Format vendor name: uppercase and replace spaces with underscores
        vendor_name = re.sub(r'[^a-zA-Z0-9\s]', '', result['vendor'])  # Keep spaces, remove other special chars
        vendor_name = vendor_name.strip().upper().replace(' ', '_')
        
        # Create new filename in required format
        new_filename = f"COI_{vendor_name}_{date_str}.pdf"
        new_path = os.path.join(COIS_FOLDER, new_filename)
        
        # Perform the rename
        print(f"{pdf_file} -> {new_filename}")
        os.rename(old_path, new_path)
        
        return new_filename
        
    except Exception as e:
        return pdf_file

def update_vendor_coi_date(vendor_name, new_date):
    """Update the COI expiration date for a vendor in the spreadsheet"""
    try:
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        
        service = build('sheets', 'v4', credentials=creds)
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='VENDORS!B:B'
        ).execute()
        
        rows = result.get('values', [])
        vendor_row = None
        for i, row in enumerate(rows):
            if row and row[0].strip() == vendor_name:
                vendor_row = i + 1  # Convert to 1-based index
                break
        
        if not vendor_row:
            return False
            
        date_str = datetime.strptime(new_date, '%m/%d/%Y').strftime('%m%d%y')
        
        range_name = f'VENDORS!G{vendor_row}'
        body = {
            'values': [[date_str]]
        }
        
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        return True
        
    except Exception as e:
        return False

def process_single_pdf(pdf_file):
    """Process a single PDF"""
    try:
        pdf_path = os.path.join(COIS_FOLDER, pdf_file)
        
        images = pdf_to_images(pdf_path)
        if not images:
            return None
            
        first_page = images[0]
        extracted_texts = {}
        
        for region_name, bounds in TEXT_REGIONS.items():
            region_bounds = get_intersection_bounds(
                first_page,
                vertical_range=bounds['vertical_range'],
                horizontal_range=bounds['horizontal_range']
            )
            
            with ThreadPoolExecutor() as executor:
                future = executor.submit(
                    extract_text_from_region,
                    first_page,
                    region_bounds
                )
                try:
                    extracted_texts[region_name] = future.result(timeout=VISION_API_TIMEOUT)
                except TimeoutError:
                    return None
        
        updated_coi_date = extract_dates_from_text(extracted_texts['date'].strip())
        date_str = updated_coi_date.strftime('%m/%d/%Y') if updated_coi_date else None
        
        if not date_str:
            return None
            
        vendor_list = get_active_vendor_list()
        
        vendor_name = extract_vendor_name(extracted_texts['vendor'], vendor_list)
        
        result = {
            'date': date_str,
            'vendor': vendor_name or 'UNKNOWN'
        }
        
        if vendor_name and vendor_name != 'UNKNOWN':
            update_vendor_coi_date(vendor_name, date_str)
        
        return result
                
    except Exception as e:
        print(f"Error: {e}")
        return None

def process_pdfs():
    """Process PDFs with minimal logging"""
    pdf_files = get_pdf_files()
    if not pdf_files:
        return
    
    for pdf_file in pdf_files:
        try:
            result = process_single_pdf(pdf_file)
            if result:
                rename_file_with_date(pdf_file, result)
        except KeyboardInterrupt:
            sys.exit(0)
        except Exception as e:
            print(f"Error processing {pdf_file}")

def main():
    try:
        init_vision_client()
        process_pdfs()
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print(f"Error: {e}")

def get_openai_client():
    """Initialize and return OpenAI client"""
    try:
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            print("Warning: OPENAI_API_KEY not found in environment variables")
            print(f"Current env variables: {dict(os.environ)}")  # Temporary debug line
            raise ValueError("OPENAI_API_KEY not found in environment variables")
        
        client = OpenAI(
            api_key=api_key,
            timeout=30.0  # Add timeout for API calls
        )
        return client
    except Exception as e:
        print(f"Error initializing OpenAI client: {str(e)}")
        raise

def extract_vendor_name(text, vendor_list=None):
    """Extract vendor name from text using OpenAI, comparing against known vendor list"""
    try:
        client = get_openai_client()
        if not client:
            return None
            
        vendor_context = "No vendor list available"
        if vendor_list:
            vendor_context = "\n".join(f"{i+1}. {v}" for i, v in enumerate(vendor_list))
            
        prompt = f"""Using the following list of active vendors we have on file:

{vendor_context}

And examining this Certificate of Insurance text:
{text[:1000]}

Which vendor from our list do you think is associated with this document?
If none match closely enough, return 'UNKNOWN'.
Respond ONLY with the exact vendor name from the list or 'UNKNOWN'."""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a precise COI analyzer. Match the document to our known vendor list."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=50
        )
        
        vendor_name = response.choices[0].message.content.strip()
        return vendor_name if vendor_name != "UNKNOWN" else None
    except Exception as e:
        print(f"Error extracting vendor name: {e}")
        return None

def get_active_vendor_list():
    """Retrieve a list of active vendor names from the spreadsheet"""
    try:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(f"Service account file not found: {SERVICE_ACCOUNT_FILE}")
            
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='VENDORS!B2:B'
        ).execute()
        
        rows = result.get('values', [])
        return [row[0] for row in rows if row]
        
    except Exception as e:
        print(f"Error loading vendor list: {e}")
        return []

if __name__ == "__main__":
    main()
