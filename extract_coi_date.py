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

# Constants
COIS_FOLDER = '/Users/cordo/Downloads/COIS'
# Update to use new vision key file
VISION_CREDENTIALS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vision_key_new.json')

# Add timeout constants
PDF_CONVERSION_TIMEOUT = 30  # seconds
VISION_API_TIMEOUT = 20  # seconds

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

def rename_file_with_date(pdf_file, extracted_date):
    """Rename the PDF file to include the extracted date"""
    try:
        old_path = os.path.join(COIS_FOLDER, pdf_file)
        
        # Get the current name without extension
        base_name = os.path.splitext(pdf_file)[0]
        
        # Format date for filename (MMDDYY)
        date_str = datetime.strptime(extracted_date, '%m/%d/%Y').strftime('%m%d%y')
        
        # Only add date if filename starts with COI_
        if base_name.startswith('COI_'):
            new_filename = f"{base_name}_{date_str}.pdf"
            new_path = os.path.join(COIS_FOLDER, new_filename)
            
            # Rename file
            os.rename(old_path, new_path)
            print(f"✓ Renamed file to: {new_filename}")
            return new_filename
        return pdf_file
    except Exception as e:
        print(f"Error renaming file: {e}")
        return pdf_file

def process_single_pdf(pdf_file):
    """Process a single PDF"""
    try:
        pdf_path = os.path.join(COIS_FOLDER, pdf_file)
        
        # Convert PDF to image
        images = pdf_to_images(pdf_path)
        if not images:
            return None
            
        first_page = images[0]
        intersection_bounds = get_intersection_bounds(
            first_page,
            vertical_range=(35, 45),
            horizontal_range=(50, 69)
        )
        
        # Extract text with timeout
        with ThreadPoolExecutor() as executor:
            future = executor.submit(
                extract_text_from_region,
                first_page,
                intersection_bounds
            )
            try:
                extracted_text = future.result(timeout=VISION_API_TIMEOUT)
                updated_coi_date = extract_dates_from_text(extracted_text.strip())
                
                if updated_coi_date:
                    date_str = updated_coi_date.strftime('%m/%d/%Y')
                    rename_file_with_date(pdf_file, date_str)
                    return date_str
                return None
            except TimeoutError:
                return None
                
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return None

def process_pdfs():
    """Process PDFs with improved error handling"""
    pdf_files = get_pdf_files()
    if not pdf_files:
        return
    
    results = []
    # Process files one at a time with timeouts
    for pdf_file in pdf_files:
        try:
            updated_coi_date = process_single_pdf(pdf_file)
            if updated_coi_date:
                results.append((pdf_file, updated_coi_date))
                print(f"Processed {pdf_file}: {updated_coi_date}")
        except KeyboardInterrupt:
            sys.exit(0)
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")
            continue
    
    print("\nProcessing Summary:")
    for filename, date in results:
        print(f"- {filename}: {date}")

def main():
    print("Starting COI text extraction...")
    try:
        init_vision_client()
        process_pdfs()
    except KeyboardInterrupt:
        print("\nProgram interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n✗ Error: {e}")
    finally:
        print("\nProcessing complete!")

if __name__ == "__main__":
    main()
