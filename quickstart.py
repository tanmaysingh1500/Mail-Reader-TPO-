
import os.path
import base64
import requests
import pdfplumber
import io
import pytesseract
from PIL import Image

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

LANGFLOW_URL = "http://localhost:7860/api/v1/webhook/[FLOW_ID]"


def authenticate():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def extract_pdf_text(pdf_bytes):

    text = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:

        for page in pdf.pages:

            # 1️⃣ Normal text extraction
            page_text = page.extract_text() or ""
            text += page_text + "\n"

            # 2️⃣ Extract text from images (OCR fallback)
            if page.images:

                page_image = page.to_image(resolution=300).original

                ocr_text = pytesseract.image_to_string(page_image)

                text += "\n[OCR TEXT]\n" + ocr_text

    return text


def get_email_body(payload):

    if "parts" in payload:
        for part in payload["parts"]:
            if part["mimeType"] == "text/plain":
                data = part["body"].get("data")
                if data:
                    return base64.urlsafe_b64decode(data).decode()

    if payload["body"].get("data"):
        return base64.urlsafe_b64decode(
            payload["body"]["data"]
        ).decode()

    return ""


def get_attachments(service, msg):

    attachment_text = ""

    parts = msg["payload"].get("parts", [])

    for part in parts:

        filename = part.get("filename")

        if filename:

            attachment_id = part["body"]["attachmentId"]

            attachment = service.users().messages().attachments().get(
                userId="me",
                messageId=msg["id"],
                id=attachment_id
            ).execute()

            data = base64.urlsafe_b64decode(attachment["data"])

            if filename.endswith(".pdf"):
                print("Processing attachment:", filename)

                attachment_text += extract_pdf_text(data)

    return attachment_text


def send_to_langflow(subject, body, attachment_text):

    payload = {
        "input_value": f"""
Email Subject:
{subject}

Email Body:
{body}

Attachment Content:
{attachment_text}
""",
        "output_type": "chat",
        "input_type": "chat"
    }

    response = requests.post(LANGFLOW_URL, json=payload)

    print("Langflow response:")
    print(response.json())


def process_emails(service):

    results = service.users().messages().list(
        userId="me",
        q="(from:tpo OR from:noreply_tpoerp) is:unread",
        maxResults=1
    ).execute()

    messages = results.get("messages", [])

    if not messages:
        print("No new TPO emails.")
        return

    for message in messages:

        msg = service.users().messages().get(
            userId="me",
            id=message["id"],
            format="full"
        ).execute()

        headers = msg["payload"]["headers"]

        subject = ""
        sender = ""

        for h in headers:
            if h["name"] == "Subject":
                subject = h["value"]
            if h["name"] == "From":
                sender = h["value"]

        body = get_email_body(msg["payload"])

        attachment_text = get_attachments(service, msg)

        print("\nFrom:", sender)
        print("Subject:", subject)
        print("Attachments extracted")

        send_to_langflow(subject, body, attachment_text)

        service.users().messages().modify(
            userId="me",
            id=message["id"],
            body={"removeLabelIds": ["UNREAD"]}
        ).execute()

        print("Email processed\n")


def main():

    try:
        creds = authenticate()

        service = build("gmail", "v1", credentials=creds)

        process_emails(service)

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()
