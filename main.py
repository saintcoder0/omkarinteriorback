import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from typing import Optional

import pytz
from dotenv import load_dotenv
from fastapi import FastAPI, Request, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, EmailStr, Field, ValidationError, field_validator
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Load environment variables from a .env file for local development
load_dotenv()

# --- FastAPI App & CORS Configuration ---
app = FastAPI()

CLIENT_ORIGIN = os.getenv("CLIENT_ORIGIN")
allowed_origins = CLIENT_ORIGIN.split(',') if CLIENT_ORIGIN else []

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

# --- Pydantic Model for Input Validation ---
class ContactForm(BaseModel):
    name: str = Field(min_length=2, max_length=200)
    email: EmailStr
    message: str = Field(min_length=5, max_length=5000)
    phone: Optional[str] = Field(default=None, max_length=50)

    @field_validator("name", "message", "phone", mode="before")
    @classmethod
    def strip_whitespace(cls, value):
        if isinstance(value, str):
            return value.strip()
        return value

# --- Helper Functions ---

def format_ist_timestamp() -> str:
    """Formats the current time to a string in IST."""
    dt = datetime.now(pytz.timezone('Asia/Kolkata'))
    return dt.strftime("%d-%m-%Y %I:%M:%S %p")

async def ensure_sheet_header(sheets_service, spreadsheet_id: str):
    """
    Checks if the header row exists in the sheet. If not, it creates and formats it.
    """
    headers = ['Timestamp', 'Name', 'Email', 'Phone', 'Message', 'IP', 'User Agent']
    try:
        # 1. Read the first row to see if headers are already there
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range='Sheet1!A1:G1'
        ).execute()
        
        existing_headers = result.get('values', [[]])[0]
        
        # 2. If the row is empty or doesn't match, write the headers
        if not existing_headers or [h.lower() for h in existing_headers] != [h.lower() for h in headers]:
            print("Header not found or incorrect. Writing new header...")
            # Write the header titles
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Sheet1!A1:G1',
                valueInputOption='RAW',
                body={'values': [headers]}
            ).execute()

            # Bonus: Format the header row to be bold, have a background color, and be frozen
            formatting_requests = {
                'requests': [
                    {
                        'repeatCell': {
                            'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                            'cell': {
                                'userEnteredFormat': {
                                    'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                                    'textFormat': {'bold': True}
                                }
                            },
                            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                        }
                    },
                    {
                        'updateSheetProperties': {
                            'properties': {'sheetId': 0, 'gridProperties': {'frozenRowCount': 1}},
                            'fields': 'gridProperties.frozenRowCount'
                        }
                    }
                ]
            }
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body=formatting_requests
            ).execute()
    except Exception as e:
        print(f"Error while ensuring sheet header: {e}")


async def append_to_sheet(data: dict):
    """Appends a new row of data to the configured Google Sheet."""
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    client_email = os.getenv("GOOGLE_CLIENT_EMAIL")
    private_key = os.getenv("GOOGLE_PRIVATE_KEY", "").replace('\\n', '\n')
    project_id = os.getenv("GOOGLE_PROJECT_ID")

    if not all([sheet_id, client_email, private_key, project_id]):
        print("Google Sheets environment variables are not fully configured. Skipping.")
        return

    try:
        creds = service_account.Credentials.from_service_account_info(
            {
                "type": "service_account",
                "project_id": project_id,
                "private_key": private_key,
                "client_email": client_email,
                "token_uri": "https://oauth2.googleapis.com/token",
            },
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        service = build('sheets', 'v4', credentials=creds)

        # --- KEY CHANGE: Ensure header exists before appending ---
        await ensure_sheet_header(service, sheet_id)

        values = [[
            format_ist_timestamp(), data.get('name'), data.get('email'),
            data.get('phone', ''), data.get('message'),
            data.get('ip', ''), data.get('userAgent', '')
        ]]

        # Appends the data to the first empty row
        service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range='Sheet1!A1', # Appending from A1 finds the next empty row automatically
            valueInputOption='USER_ENTERED', 
            body={'values': values}
        ).execute()
    except Exception as e:
        print(f"Error appending to Google Sheet: {e}")


def send_email(data: dict):
    """Sends an email notification using SMTP."""
    smtp_user = os.getenv('SMTP_USER')
    smtp_pass = os.getenv('SMTP_PASS')
    mail_to = os.getenv('MAIL_TO', smtp_user)

    if not all([smtp_user, smtp_pass]):
        raise ValueError("SMTP credentials are not configured.")
    
    assert smtp_user is not None and smtp_pass is not None

    msg = EmailMessage()
    msg['Subject'] = f"New Inquiry from {data['name']}"
    msg['From'] = f"Omkar Interiors <{smtp_user}>"
    msg['To'] = mail_to
    msg['Reply-To'] = data['email']

    html_content = (
        f"<p><strong>Name:</strong> {data['name']}</p>"
        f"<p><strong>Email:</strong> {data['email']}</p>"
        f"<p><strong>Phone:</strong> {data.get('phone', 'N/A')}</p>"
        f"<p><strong>Message:</strong><br/>{data['message'].replace(chr(10), '<br>')}</p>"
        f"<hr/>"
        f"<p><small>IP: {data.get('ip')} &middot; UA: {data.get('userAgent')}</small></p>"
    )
    msg.add_alternative(html_content, subtype='html')

    with smtplib.SMTP_SSL(os.getenv('SMTP_HOST', 'smtp.gmail.com'), int(os.getenv('SMTP_PORT', 465))) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

# --- API Endpoints ---
@app.post("/api/contact")
async def handle_contact_form(request: Request):
    try:
        body = await request.json()
        form_data = ContactForm(**body)
    except ValidationError as e:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={"ok": False, "errors": e.errors()}
        )
    except Exception:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"ok": False, "error": "Invalid request body."}
        )

    data = form_data.model_dump()
    data['ip'] = request.headers.get('x-forwarded-for') or (request.client.host if request.client else 'unknown')
    data['userAgent'] = request.headers.get('user-agent', 'Unknown')

    try:
        send_email(data)
        await append_to_sheet(data)
        return {"ok": True, "message": "Message sent successfully!"}
    except Exception as e:
        print(f"Error sending message: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail={"ok": False, "error": "Failed to send message."}
        )

@app.get("/api/health")
def health_check():
    """Provides a simple health check endpoint."""
    return {"ok": True, "service": "contact-api", "status": "active"}