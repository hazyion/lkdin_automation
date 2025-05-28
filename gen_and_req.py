import os
import base64
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone
import time
import gspread
import requests
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from twilio.rest import Client
import smtplib
from email.mime.text import MIMEText

load_dotenv()

PLACID_TOKEN = os.getenv("PLACID_TOKEN")
PLACID_URL = os.getenv("PLACID_URL")
TEMPLATE_UUID = os.getenv("PLACID_TEMPLATE_UUID")
TWILIO_SID = os.getenv("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.getenv("TWILIO_AUTH_TOKEN")
TWILIO_WHATSAPP_NUMBER = os.getenv("TWILIO_WHATSAPP_NUMBER")
YOUR_WHATSAPP = os.getenv("YOUR_WHATSAPP")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))
G_CREDENTIAL_FILE = os.getenv("G_CREDENTIAL_FILE")
PRESENTATION_ID = os.getenv("PRESENTATION_ID")

# Set up Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive",
         "https://www.googleapis.com/auth/presentations"]
# creds = ServiceAccountCredentials.from_json_keyfile_name(
#     G_CREDENTIAL_FILE, scope)

b64_creds = os.environ["SERVICE_ACCOUNT_CREDS"]
creds_dict = json.loads(base64.b64decode(b64_creds).decode("utf-8"))
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

auditDetails = {
    "name": "Carl Jason",
    "title": "Software Developer",
    "followers": 801,
    "pb_s": 4.3,
    "cs_s": 8.1,
    "es_s": 8.5,
    "po_s": 2.5,
    "ape_s": 8.2,
    "recommendations": "",
    "ovr": ""
}


def convert_and_round_down(x):
    if not 0 <= x <= 10:
        raise ValueError("Input must be between 0 and 10.")
    return int(x) * 10


def start_process():
    try:
        # Get data from spreadsheet
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet('Template')

        auditDetails["name"] = sheet.acell("A1").value
        auditDetails["followers"] = sheet.acell("C39").value
        auditDetails["title"] = sheet.acell("C40").value
        auditDetails["recommendations"] = sheet.acell("C42").value.strip()

        sum = 0
        auditDetails["pb_s"] = convert_and_round_down(
            float(sheet.acell("D8").value))
        sum += float(sheet.acell("D8").value)
        auditDetails["cs_s"] = convert_and_round_down(
            float(sheet.acell("D23").value))
        sum += float(sheet.acell("D23").value)
        auditDetails["es_s"] = convert_and_round_down(
            float(sheet.acell("D35").value))
        sum += float(sheet.acell("D35").value)
        auditDetails["po_s"] = convert_and_round_down(
            float(sheet.acell("D30").value))
        sum += float(sheet.acell("D30").value)
        auditDetails["ape_s"] = convert_and_round_down(
            float(sheet.acell("D18").value))
        sum += float(sheet.acell("D18").value)
        auditDetails["os_s"] = convert_and_round_down(sum / 5)

        if auditDetails["os_s"] <= 40:
            auditDetails["ovr"] = "Poor"
        elif auditDetails["os_s"] <= 60:
            auditDetails["ovr"] = "Good"
        elif auditDetails["os_s"] <= 80:
            auditDetails["ovr"] = "Great"
        else:
            auditDetails["ovr"] = "Excellent"

        recipient_email = sheet.acell("C43").value

        # Create new report and enter data
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        copy_response = drive_service.files().copy(
            fileId=PRESENTATION_ID,
            body={"name": f"Temp Copy for {auditDetails['name']}"}
        ).execute()
        copy_presentation_id = copy_response["id"]

        requests_body = {
            "requests": []
        }

        for key, value in auditDetails.items():
            placeholder = key.upper()
            repValue = str(value)
            if placeholder in ["ES_S", "PB_S", "CS_S", "PO_S", "APE_S", "OS_S"]:
                repValue += "%"
            # if placeholder in ["RECOMMENDATIONS"]:
            #     continue
            requests_body["requests"].append({
                "replaceAllText": {
                    "containsText": {
                        "text": f"[[{placeholder}]]",
                        "matchCase": True
                    },
                    "replaceText": repValue
                }
            })

        requests_body["requests"].append(
            {
                "createParagraphBullets": {
                    "objectId": "p1_i140",
                    "textRange": {
                        "type": "ALL"
                    },
                    "bulletPreset": "BULLET_ARROW_DIAMOND_DISC"
                }
            })

        slides_service.presentations().batchUpdate(
            presentationId=copy_presentation_id,
            body=requests_body
        ).execute()

        # Update report permissions and name
        fileUpdates = {
            'type': 'anyone',
            'role': 'reader',
            'name': f"{auditDetails["name"].replace(" ", "_")}_profile_audit"
        }

        drive_service.permissions().create(
            fileId=copy_presentation_id,
            body=fileUpdates
        ).execute()

        # presentation = slides_service.presentations().get(
        #     presentationId=copy_presentation_id).execute()
        # for idx, slide in enumerate(presentation.get('slides')):
        #     print(f"\nSlide {idx + 1}:")
        #
        #     for element in slide.get('pageElements', []):
        #         obj_id = element.get('objectId')
        #         shape_type = element.get('shape', {}).get('shapeType', 'UNKNOWN')
        #         text_elements = element.get('shape', {}).get(
        #             'text', {}).get('textElements', [])
        #
        #         text = ""
        #         for te in text_elements:
        #             if 'textRun' in te:
        #                 text += te['textRun'].get('content', '')
        #
        #         print(
        #             f"- ID: {obj_id} | Type: {shape_type} | Text: '{text.strip()}'")

        # request = drive_service.files().export_media(
        #     fileId=copy_presentation_id, mimeType='application/pdf')
        #
        # name = auditDetails["name"].replace(" ", "_")
        # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        #
        # filename = f"{name}_{timestamp}.pdf"

        # fh = io.BytesIO()
        # downloader = MediaIoBaseDownload(fh, request)
        #
        # done = False
        # while not done:
        #     status, done = downloader.next_chunk()
        #
        # with open("your_filename.pdf", "wb") as f:
        #     f.write(fh.getbuffer())
        # pdf_file = request.execute()
        # with open(f"pdfs/{filename}", 'wb') as f:
        #     f.write(pdf_file)

        # print(f"PDF downloaded as {filename}")

        # PDF link
        pdf_link = f"https://docs.google.com/presentation/d/{
            copy_presentation_id}/export/pdf"

        sheet.update_acell("C45", pdf_link)

        # Send WhatsApp message via Twilio
        twilio_client = Client(TWILIO_SID, TWILIO_AUTH_TOKEN)
        message = twilio_client.messages.create(
            from_=f"{TWILIO_WHATSAPP_NUMBER}",
            to=f"{YOUR_WHATSAPP}",
            body=f"{auditDetails["name"]}'s report is ready: {
                pdf_link}\nReply 'yes' to approve and send to {recipient_email}",
            media_url=[pdf_link]
        )
        print("Message sent. SID:", message)

        start_time = time.time()
        approved = False
        print("Waiting for approval...")

        curDatetime = datetime.now(timezone.utc)

        # Poll Twilio for replies
        while time.time() - start_time < 300:  # 5 minutes timeout
            messages = twilio_client.messages.list(
                to=f"{TWILIO_WHATSAPP_NUMBER}", limit=5)
            print(messages)
            for msg in messages:
                print('datesenttime ', msg.date_sent)
                if msg.direction == "inbound" and msg.date_sent > curDatetime and msg.from_ == f"{YOUR_WHATSAPP}":
                    if msg.body.strip().lower() in ["yes", "approve"]:
                        approved = True
                        print("Approval received.")
                        send_email(recipient_email, pdf_link)
                        # Delete pdf
                        # drive_service.files().delete(fileId=copy_presentation_id).execute()
                        break
            if approved:
                break
            time.sleep(5)

        if not approved:
            print("Approval not received in time.")
            return

    except Exception as error:
        print(error)


def send_email(to_email, pdf_url):
    try:
        msg = MIMEText(f"Hi {auditDetails['name']},\n\nAttached is your LinkedIn profile audit. It outlines what's working, and where a few smart tweaks can amplify how you're seen by the right audience.\n\nYou already bring the credibility - this helps make it visible.\n\nLet me know what you think. Happy to walk you through it anytime.\n\nBest,\nAmbuj\n\n{pdf_url}")
        msg["Subject"] = "Your LinkedIn Audit is ready - actionable insights to upgrade your brand"
        msg["From"] = EMAIL_FROM
        msg["To"] = to_email

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, to_email, msg.as_string())

        print("Email sent to", to_email)

    except Exception as e:
        print(e)


if __name__ == "__main__":
    start_process()
    # send_email("carljason.3012@gmail.com", "https://example.com")
