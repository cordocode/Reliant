import os
from google.cloud import vision_v1
from pdf2image import convert_from_path
import tempfile
from PIL import Image
import gspread
from google.oauth2 import service_account
import re
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

# Configuration variables
INPUT_DIR = os.getenv('INPUT_DIR', "/users/cordo/downloads/NAME_IT")
OUTPUT_DIR = os.getenv('OUTPUT_DIR', "/users/cordo/downloads/NAMED")
VISION_CREDENTIALS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vision_key_new.json')
SHEETS_CREDENTIALS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sheets_key.json')
SPREADSHEET_NAME = 'VENDOR_MANAGEMENT'  # Main spreadsheet name
SHEET_CODE_NAME = 'CODE1'                # Sheet for property code data
SHEET_VENDORS_NAME = 'VENDORS'          # Sheet for vendor names & invoice numbers

# Initialize the Google Cloud Vision client
def init_vision_client():
    """Initialize Vision client with explicit credentials"""
    try:
        if not os.path.exists(VISION_CREDENTIALS):
            raise FileNotFoundError(f"Credentials file not found at {VISION_CREDENTIALS}")
            
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = VISION_CREDENTIALS
        client = vision_v1.ImageAnnotatorClient()
        return client
    except Exception as e:
        print(f"✗ Error initializing Vision client: {e}")
        raise

client = init_vision_client()

# Replace get_gspread_client to use 'sheets_key.json'
def get_gspread_client():
    """Get gspread client using service account credentials"""
    try:
        credentials = service_account.Credentials.from_service_account_file(SHEETS_CREDENTIALS, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ])
        return gspread.authorize(credentials)
    except Exception as e:
        print(f"✗ Error initializing Sheets client: {e}")
        raise

gspread_client = get_gspread_client()
spreadsheet = gspread_client.open(SPREADSHEET_NAME)
sheet_code = spreadsheet.worksheet(SHEET_CODE_NAME)      # For property code data
sheet_vendors = spreadsheet.worksheet(SHEET_VENDORS_NAME)  # For vendor name / invoice numbers

def process_directory(input_dir):
    """
    Process each file in the input directory, extract details, rename files,
    and move them if all details are identified.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)  # Ensure the output directory exists

    for filename in os.listdir(input_dir):
        if filename.lower().endswith((".pdf", ".jpeg", ".jpg", ".png")):
            print(f"Processing file: {filename}\n")
            print("----------------------------------------")
            try:
                file_path = os.path.join(input_dir, filename)

                # Extract text from the file
                if filename.lower().endswith(".pdf"):
                    text = extract_text_from_pdf(file_path)
                else:
                    text = extract_text_from_image(file_path)
                print("Extracted Text:", text)

                # Extract relevant details
                property_code = determine_property_code(text)
                vendor_name = determine_vendor_name(text)
                invoice_date = determine_invoice_date(text)
                invoice_number = determine_invoice_number(vendor_name, invoice_date, text)

                # Generate new file name
                new_file_name = generate_file_name(
                    property_code if property_code != "Unknown" else "MISSING_PROPERTY_CODE",
                    vendor_name if vendor_name != "No Name" else "MISSING_VENDOR",
                    invoice_date if invoice_date != "No Date Found" else "MISSING_DATE",
                    invoice_number if invoice_number != "No Invoice Number" else "MISSING_NUMBER",
                )
                print(f"Generated File Name: {new_file_name}")

                # Define new file path
                new_file_path = os.path.join(input_dir, f"{new_file_name}{os.path.splitext(filename)[1]}")

                # Rename the file in the input directory
                os.rename(file_path, new_file_path)
                print(f"Renamed File: {new_file_path}")

                # Move the file to the output directory if all details are complete
                if "MISSING" not in new_file_name:
                    final_output_path = os.path.join(OUTPUT_DIR, os.path.basename(new_file_path))
                    os.rename(new_file_path, final_output_path)
                    print(f"Moved File to: {final_output_path}")
                else:
                    print("File not moved due to missing elements.")

            except Exception as e:
                print(f"Failed to process {filename}: {e}")
            print("----------------------------------------\n")

def extract_text_from_pdf(pdf_path):
    with tempfile.TemporaryDirectory() as temp_dir:
        images = convert_from_path(pdf_path, dpi=100, output_folder=temp_dir, fmt='jpeg')
        full_text = ""
        for image in images[:2]:
            with tempfile.NamedTemporaryFile(suffix=".jpeg", delete=False) as temp_image:
                image.convert("RGB").save(temp_image.name, "JPEG", quality=95)
                full_text += extract_text_from_image(temp_image.name) + "\n"
        return full_text.strip()

def extract_text_from_image(image_path):
    with open(image_path, "rb") as image_file:
        content = image_file.read()
    response = client.document_text_detection(image=vision_v1.types.Image(content=content))
    return response.full_text_annotation.text.strip()

def determine_property_code(text):
    """
    Use 'CODE' sheet for property codes.
    Column C is weighted 5x more heavily than other columns.
    """
    codes = sheet_code.get_all_values()[1:]  # Skip header row
    max_matches = 0
    best_code = "Unknown"

    for row in codes:
        property_code = row[0]  # Column A
        keywords = row[1:]      # Columns B-I
        
        # Calculate matches with Column C weighted 5x
        matches = 0
        for i, keyword in enumerate(keywords):
            if keyword and keyword.lower() in text.lower():
                if i == 1:  # Column C (index 1 since we started from B)
                    matches += 5  # Weight Column C matches 5x more
                else:
                    matches += 1

        if matches > max_matches:
            max_matches = matches
            best_code = property_code

    return best_code

def determine_vendor_name(text):
    """Use 'VENDORS' sheet for vendor names."""
    names = sheet_vendors.col_values(2)[1:]  # Column B, skip the header row

    # Step 1: Exact match for the full name in the text
    for name in names:
        if name.lower() in text.lower():
            return name

    # Step 2: Search for split names across lines
    for name in names:
        split_name_parts = name.lower().split()
        if all(part in text.lower() for part in split_name_parts):
            return name

    return "XXX"

def determine_invoice_date(text):
    """Determine the invoice date by finding all dates in the text and returning the furthest future date."""
    date_patterns = [
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",  # Matches formats like xx/xx/xx or xx/xx/xxxx
        r"\b\d{1,2}-\d{1,2}-\d{2,4}\b",  # Matches formats like xx-xx-xx or xx-xx-xxxx
        r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2}, \d{4}\b",  # Matches formats like Mar 25, 2024
        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b"  # Matches formats like September 26, 2024
    ]

    potential_dates = []
    date_text_map = {}

    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            try:
                if "/" in match or "-" in match:
                    # Handle numeric formats
                    if "-" in match:
                        match = match.replace("-", "/")
                    if len(match.split("/")[-1]) == 2:  # Check for 2-digit year
                        parsed_date = datetime.strptime(match, "%m/%d/%y")
                    else:
                        parsed_date = datetime.strptime(match, "%m/%d/%Y")
                else:
                    # Handle textual formats
                    try:
                        parsed_date = datetime.strptime(match, "%b %d, %Y")
                    except ValueError:
                        parsed_date = datetime.strptime(match, "%B %d, %Y")
                potential_dates.append(parsed_date)
                date_text_map[parsed_date] = match  # Store original text
            except ValueError:
                continue

    if potential_dates:
        furthest_date = max(potential_dates)
        return date_text_map[furthest_date]  # Return as it appeared in the text
    return "No Date Found"

def determine_invoice_number(vendor_name, invoice_date, text):
    """
    Determine the invoice number based on the vendor name, invoice date, and text.
    Includes a regex pattern to match the structure of the invoice number and 
    searches the text comprehensively.
    """
    if not vendor_name or invoice_date == "No Name":
        return "No Invoice Number"

    # Find the invoice date as a substring in the text
    date_index = text.find(invoice_date)

    if date_index != -1:
        # Count the number of words from the top to the found position
        words = text.split()
        word_position = len(text[:date_index].split()) + 1  # 1-based index
        print(f"Word count position of the date: {word_position}")
    else:
        return "Invoice date not found in text."

    try:
        records = sheet_vendors.get_all_values()[1:]  # Skip header
        for row in records:
            if row[1].lower() == vendor_name.lower():  # Column B
                sheet_invoice_number = row[7]          # Column H
                print(f"Invoice number in sheet next to vendor name: {sheet_invoice_number}")

                # Convert the invoice number into a flexible regex pattern
                regex_pattern = ""
                for ch in sheet_invoice_number:
                    if ch.isdigit():
                        regex_pattern += "\\d"
                    elif ch.isalpha():
                        regex_pattern += "[A-Za-z]"
                    else:
                        regex_pattern += re.escape(ch)  # Escape special characters
                print(f"Regex pattern being used: {regex_pattern}")

                # Comprehensive search: Above and below the word position
                for offset in range(len(words)):
                    for index in [word_position + offset, word_position - offset]:
                        if 0 <= index < len(words) and re.fullmatch(regex_pattern, words[index]):
                            return words[index]  # Return the matched invoice number

                return "No Matching Invoice Number Found"
    except Exception as e:
        print(f"Error accessing Google Sheet for invoice number: {e}")

    return "No Invoice Number"

def generate_file_name(property_code, vendor_name, invoice_date, invoice_number):
    """
    Generate a new file name based on the extracted information.
    - Property Code: exact as extracted
    - Vendor Name: uppercase with underscores between words
    - Invoice Date: converted to DDMMYY format without separators
    - Invoice Number: only numbers
    """
    # Ensure property code is as-is
    formatted_property_code = property_code

    # Format vendor name to uppercase with underscores
    formatted_vendor_name = "_".join(vendor_name.upper().split())

    # Convert the invoice date to DDMMYY format
    formatted_date = "XXXXXX"  # Default value if no format matches
    date_formats = [
        "%m/%d/%Y",  # MM/DD/YYYY
        "%m/%d/%y",  # MM/DD/YY
        "%b %d, %Y",  # Abbreviated month name (e.g., "Dec 12, 2023")
        "%B %d, %Y",  # Full month name (e.g., "December 12, 2023")
        "%b %d, %y",  # Abbreviated month name with 2-digit year
        "%B %d, %y",  # Full month name with 2-digit year
        "%d-%m-%Y",   # DD-MM-YYYY
        "%d-%m-%y",   # DD-MM-YY
    ]
    
    for date_format in date_formats:
        try:
            parsed_date = datetime.strptime(invoice_date, date_format)
            formatted_date = parsed_date.strftime("%m%d%y")  # Convert to DDMMYY
            break
        except ValueError:
            continue  # Try the next format

    # Extract only digits from the invoice number
    formatted_invoice_number = "".join(filter(str.isdigit, invoice_number))

    # Combine all parts into the final file name
    new_file_name = f"{formatted_property_code}_{formatted_vendor_name}_{formatted_date}_{formatted_invoice_number}"
    return new_file_name

if __name__ == "__main__":
    process_directory(INPUT_DIR)