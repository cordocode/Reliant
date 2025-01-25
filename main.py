from google.oauth2 import service_account
from googleapiclient.discovery import build
import pdf2image
from pathlib import Path
import tempfile
import shutil
import os
from vision import convert_pdf_to_images, process_page, extract_names

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'sheets_key.json'
SHEET_NAME = 'VENDORS'
SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'

def write_to_sheet(service, names):
    """Write extracted text to Google Sheet."""
    try:
        if not names:
            print("No names to write to sheet")
            return
            
        print(f"\nWriting {len(names)} names to sheet...")
        range_name = f"'{SHEET_NAME}'!B2:B{len(names) + 1}"
        values = [[name] for name in names]
        body = {'values': values}
        
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        print("Successfully wrote names to sheet")
        print("\nNames written:")
        for i, name in enumerate(names, 1):
            print(f"{i}. {name}")
    except Exception as e:
        print(f"Error writing to sheet: {e}")

def main():
    # Set up Google Sheets credentials and service
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    sheets_service = build('sheets', 'v4', credentials=credentials)
    
    pdf_path = str(Path.home() / "Downloads" / "blvd.pdf")
    
    try:
        images = convert_pdf_to_images(pdf_path)
        if not images:
            print("Failed to convert PDF to images")
            return
            
        all_names = []
        
        # Process each page
        for image in images:
            document = process_page(image)
            page_names = extract_names(document)
            if page_names:
                all_names.extend(page_names)
                print(f"Found {len(page_names)} names")
        
        if all_names:
            write_to_sheet(sheets_service, all_names)
        else:
            print("No names found in any page")
            
    except Exception as e:
        print(f"Error in processing: {e}")

if __name__ == '__main__':
    main()