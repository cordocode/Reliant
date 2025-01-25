import os
import io
from google.cloud import vision
import pdf2image  # We still need this just for PDF to image conversion
from pathlib import Path
from google.oauth2 import service_account
import shutil
import tempfile
import json

# Update credentials configuration
CREDENTIALS_FILE = 'excel-extractor-448918-aced0e7d9e55.json'

def validate_credentials(cred_file):
    """Validate that the credentials file has all required fields."""
    required_fields = ['client_email', 'token_uri', 'private_key']
    try:
        with open(cred_file, 'r') as f:
            creds = json.load(f)
        
        missing_fields = [field for field in required_fields if field not in creds]
        if missing_fields:
            raise ValueError(f"Credentials file missing required fields: {missing_fields}")
        return True
    except (json.JSONDecodeError, FileNotFoundError) as e:
        raise ValueError(f"Credentials error: {str(e)}")

def convert_pdf_to_images(pdf_path):
    """Convert all pages of PDF to images."""
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_pdf = os.path.join(temp_dir, 'temp.pdf')
            shutil.copy2(pdf_path, temp_pdf)
            
            # Convert all pages
            images = pdf2image.convert_from_path(temp_pdf)
            
            if images:
                return images
            return None
            
    except Exception as e:
        print(f"Error converting PDF to image: {e}")
        return None

def process_page(image):
    """Process a single page and return its document text detection."""
    try:
        validate_credentials(CREDENTIALS_FILE)
        credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
        client = vision.ImageAnnotatorClient(credentials=credentials)
        
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='PNG')
        content = img_byte_arr.getvalue()
        
        image = vision.Image(content=content)
        response = client.document_text_detection(image=image)
        return response.full_text_annotation
        
    except Exception as e:
        raise Exception(f"Vision API error: {str(e)}")

def extract_names(document):
    """Extract names from percentage-based column position."""
    names = []
    
    # Get page dimensions from first page
    if not document.pages:
        return names
    
    # Calculate full page width from first page's vertices
    first_page = document.pages[0]
    all_x_coords = [v.x for block in first_page.blocks for p in block.paragraphs for v in p.bounding_box.vertices]
    page_width = max(all_x_coords)
    
    # Define column boundaries as percentages
    LEFT_PCT = 0.035   # Column starts at 3.5% from left edge
    RIGHT_PCT = 0.325  # Column ends at 32.5% from left edge
    
    # Calculate actual x-coordinates
    target_left = page_width * LEFT_PCT
    target_right = page_width * RIGHT_PCT
    
    print(f"\nPage width detected: {page_width}")
    print(f"Looking for text between:")
    print(f"  {target_left:.1f}px ({LEFT_PCT*100}% from left)")
    print(f"  {target_right:.1f}px ({RIGHT_PCT*100}% from left)")
    
    for page in document.pages:
        for block in page.blocks:
            for paragraph in block.paragraphs:
                vertices = paragraph.bounding_box.vertices
                text_left = min(v.x for v in vertices)
                text_right = max(v.x for v in vertices)
                
                # Calculate text position as percentage of page width
                text_left_pct = text_left / page_width
                text_right_pct = text_right / page_width
                
                # Check if text falls within our target column percentages
                if (LEFT_PCT <= text_left_pct <= RIGHT_PCT and 
                    LEFT_PCT <= text_right_pct <= RIGHT_PCT):
                    
                    text = ''.join([''.join([s.text for s in w.symbols]) for w in paragraph.words]).strip()
                    
                    if (text and 
                        not text.upper() == 'NAME' and
                        not text.isdigit() and
                        len(text.strip()) > 1):
                        
                        print(f"Found text at {text_left_pct*100:.1f}%: {text}")
                        names.append(text)
    
    print(f"\nFound {len(names)} names in the column")
    return names

def main():
    # PDF file path
    pdf_path = str(Path.home() / "Downloads" / "blvd.pdf")
    
    # Convert all pages to images
    images = convert_pdf_to_images(pdf_path)
    if not images:
        print("Failed to convert PDF to images")
        return
    
    all_names = []
    
    # Process each page
    for image in images:
        try:
            document = process_page(image)
            names = extract_names(document)
            all_names.extend(names)
        except Exception as e:
            print(f"Error processing image with Vision API: {e}")
    
    print("\nAll extracted names:")
    print("--------------------")
    for name in all_names:
        print(name)

if __name__ == "__main__":
    main()
