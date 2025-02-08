from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
from openai import OpenAI
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json

load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = '/Users/cordo/Documents/RELIANT_SCRIPTS/email_key.json'
TENANT_SPREADSHEET_ID = '18-a4IUWgZ27l_dlrJA7L_MmDmpDLVWEEUCTIDAsUBuo'
VENDOR_SPREADSHEET_ID = '19PKId-MCbmA1iG_DwbXqwR-LbQy2b6YBAh-yyMtLI-s'
TENANT_SHEET_NAME = 'TENANT'
VENDOR_SHEET_NAME = 'VENDORS'
SENDER_EMAIL = 'admin@reliant-pm.com'
APP_PASSWORD = 'otftekojfvvhqxra'

# Initialize OpenAI client
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

class TenantColumns:
    PROPERTY_NUMBER = 'B'
    TENANT_ADDRESS = 'C'
    TENANT_NAME = 'D'
    CONTACT_NAME = 'H'
    CONTACT_EMAIL = 'J'

class VendorColumns:
    PROPERTY = 'A'
    VENDOR_NAME = 'B'
    PRIMARY_CONTACT_NAME = 'C'
    VENDOR_EMAIL = 'E'
    VENDOR_TYPE = 'F'

def get_column_letter_to_index(letter):
    """Convert column letter to zero-based index."""
    return ord(letter.upper()) - ord('A')

def initialize_sheets_service():
    """Initialize and return the Google Sheets service."""
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)
    return service

def read_tenant_sheet(service, range_name):
    """Read data from tenant sheet and return structured list of dicts."""
    result = service.spreadsheets().values().get(
        spreadsheetId=TENANT_SPREADSHEET_ID,
        range=f'{TENANT_SHEET_NAME}!{range_name}'
    ).execute()
    
    values = result.get('values', [])
    if not values:
        return []

    # Skip header row
    structured_data = []
    for row in values[1:]:
        while len(row) < get_column_letter_to_index(TenantColumns.CONTACT_EMAIL) + 1:
            row.append('')
        tenant_info = {
            'property_number': row[get_column_letter_to_index(TenantColumns.PROPERTY_NUMBER)],
            'tenant_address': row[get_column_letter_to_index(TenantColumns.TENANT_ADDRESS)],
            'tenant_name': row[get_column_letter_to_index(TenantColumns.TENANT_NAME)],
            'contact_name': row[get_column_letter_to_index(TenantColumns.CONTACT_NAME)],
            'contact_email': row[get_column_letter_to_index(TenantColumns.CONTACT_EMAIL)]
        }
        structured_data.append(tenant_info)

    return structured_data

def read_vendor_sheet(service, range_name):
    """Read data from vendor sheet and return structured list of dicts."""
    result = service.spreadsheets().values().get(
        spreadsheetId=VENDOR_SPREADSHEET_ID,
        range=f'{VENDOR_SHEET_NAME}!{range_name}'
    ).execute()
    
    values = result.get('values', [])
    if not values:
        return []

    structured_data = []
    for row in values[1:]:
        while len(row) < get_column_letter_to_index(VendorColumns.VENDOR_TYPE) + 1:
            row.append('')
        vendor_info = {
            'property': row[get_column_letter_to_index(VendorColumns.PROPERTY)],
            'vendor_name': row[get_column_letter_to_index(VendorColumns.VENDOR_NAME)],
            'primary_contact_name': row[get_column_letter_to_index(VendorColumns.PRIMARY_CONTACT_NAME)],
            'vendor_email': row[get_column_letter_to_index(VendorColumns.VENDOR_EMAIL)],
            'vendor_type': row[get_column_letter_to_index(VendorColumns.VENDOR_TYPE)]
        }
        structured_data.append(vendor_info)

    return structured_data

def get_sheets_data():
    """Get structured data from both sheets."""
    service = initialize_sheets_service()
    tenant_data = read_tenant_sheet(service, 'A1:J')
    vendor_data = read_vendor_sheet(service, 'A1:F')
    return tenant_data, vendor_data

def send_email(recipient_email, subject, body):
    """Send email using configured SMTP server."""
    try:
        message = MIMEMultipart()
        message['From'] = SENDER_EMAIL
        message['To'] = recipient_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(message)
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

def query_openai(prompt):
    """Send a prompt to GPT (gpt-4o) and return the response as plain text."""
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.5,
            max_tokens=2000  # Increased to handle the data plus answer
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"Error querying OpenAI: {e}")
        return ""

# ---------------------------------------------
#  BUILDING THE PROMPT (DATA + USER QUERY)
# ---------------------------------------------
def format_data_for_prompt(tenant_data, vendor_data):
    """
    Convert all tenant_data, vendor_data rows to a compact CSV or list format 
    that GPT can read within one prompt.
    Try to keep lines short but enough detail to let GPT find correct matches.
    """
    lines = []

    lines.append("TENANTS (property_number, tenant_name, contact_name, contact_email):")
    for t in tenant_data:
        # Make a single line:  102, Bob's Diner, Bob Smith, bob@example.com
        line = f"{t['property_number']}, {t['tenant_name']}, {t['contact_name']}, {t['contact_email']}"
        lines.append(line)

    lines.append("")
    lines.append("VENDORS (property, vendor_name, primary_contact_name, vendor_email, vendor_type):")
    for v in vendor_data:
        # e.g.:  100, ACCO Engineering, John Doe, jdoe@acco.com, HVAC
        line = f"{v['property']}, {v['vendor_name']}, {v['primary_contact_name']}, {v['vendor_email']}, {v['vendor_type']}"
        lines.append(line)

    return "\n".join(lines)

def ask_gpt_for_emails(user_query, tenant_data, vendor_data):
    """
    We load ALL tenants/vendors into a single prompt. Then we ask GPT:
    "Return only a JSON array of emails that match the user's request."
    """
    data_string = format_data_for_prompt(tenant_data, vendor_data)

    prompt = f"""
We have the following data of tenants and vendors:

{data_string}

The user says: "{user_query}"

Your task: Identify exactly which rows match the user's request, 
and return ONLY a JSON array of their 'contact_email' or 'vendor_email' fields. 
No other text or explanation.

If there are no matches, return an empty JSON array: []
If multiple matches, return them all.

Example output:
["bob@example.com", "mary@example.com"]

Now produce only the JSON array for this request.
"""
    raw_response = query_openai(prompt)

    # Attempt to parse as JSON list
    try:
        result = json.loads(raw_response)
        # ensure it's a list of strings
        if isinstance(result, list) and all(isinstance(x, str) for x in result):
            return result
        else:
            print("GPT returned JSON but not a list of strings. Raw:", raw_response)
            return []
    except json.JSONDecodeError:
        print("GPT did not return valid JSON. Raw:", raw_response)
        return []

# ---------------------------------------------
#         MAIN TEST LOOP
# ---------------------------------------------
if __name__ == "__main__":
    tenant_data, vendor_data = get_sheets_data()

    while True:
        user_input = input("\nEnter your request (or 'quit' to exit): ")
        if user_input.lower() == 'quit':
            break

        emails = ask_gpt_for_emails(user_input, tenant_data, vendor_data)
        print("\nGPT's matching emails:")
        print(emails)
