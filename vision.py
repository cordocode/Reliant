import os
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
    except json.JSONDecodeError:
        raise ValueError("Credentials file is not valid JSON")
    except FileNotFoundError:
        raise ValueError(f"Credentials file '{cred_file}' not found")

def convert_pdf_to_images(pdf_path):
    """Convert first page of PDF to image while preserving original."""
    try:
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a copy of the PDF in the temp directory
            temp_pdf = os.path.join(temp_dir, 'temp.pdf')
            shutil.copy2(pdf_path, temp_pdf)
            
            # Convert the temporary PDF copy to images
            images = pdf2image.convert_from_path(temp_pdf, first_page=1, last_page=1)
            
            # If conversion successful, save the first image to temp dir
            if images:
                temp_image_path = os.path.join(temp_dir, 'page_1.png')
                images[0].save(temp_image_path, 'PNG')
                return images[0]
            return None
            
    except Exception as e:
        print(f"Error converting PDF to image: {e}")
        return None

def detect_text(image):
    """Detect text in an image using Google Cloud Vision API."""
    try:
        validate_credentials(CREDENTIALS_FILE)
        credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
        client = vision.ImageAnnotatorClient(credentials=credentials)
        
        # Convert PIL Image to bytes
        import io
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='PNG')
        content = img_byte_arr.getvalue()
        
        image = vision.Image(content=content)
        response = client.text_detection(image=image)
        texts = response.text_annotations
        
        if texts:
            return texts[0].description
        return ""
    except ValueError as e:
        raise Exception(f"Credentials error: {str(e)}")

def main():
    # PDF file path
    pdf_path = str(Path.home() / "Downloads" / "blvd.pdf")
    
    # Convert first page to image
    image = convert_pdf_to_images(pdf_path)
    if not image:
        print("Failed to convert PDF to image")
        return
    
    # Extract text using Vision API
    try:
        extracted_text = detect_text(image)
        print("Extracted text from first page:")
        print("--------------------------------")
        print(extracted_text)
    except Exception as e:
        print(f"Error processing image with Vision API: {e}")

if __name__ == "__main__":
    main()
