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

scoreBarElements = {
    "pb_s": {
        "poor": "g36773a6a01a_0_2",
        "average": "g36773a6a01a_0_5",
        "amazing": "p1_i117"
    },
    "cs_s": {
        "poor": "g36773a6a01a_0_14",
        "average": "g36773a6a01a_0_11",
        "amazing": "p1_i121"
    },
    "es_s": {
        "poor": "g36773a6a01a_0_20",
        "average": "p1_i129",
        "amazing": "g36773a6a01a_0_17"
    },
    "po_s": {
        "poor": "p1_i133",
        "average": "g36773a6a01a_0_26",
        "amazing": "g36773a6a01a_0_23"
    },
    "ape_s": {
        "poor": "g36773a6a01a_0_29",
        "average": "g36773a6a01a_0_32",
        "amazing": "p1_i125"
    },
}


def mask_email(email):
    user, domain = email.split("@")
    return f"{user[0]}***@{domain}"


def mask_phone(phone):
    return f"{'*' * (len(phone) - 4)}{phone[-4:]}"


def get_score_description(score):
    if score < 50:
        return "Poor"
    elif score < 70:
        return "Average"
    else:
        return "Amazing"


def scale_and_round(score):
    return round(score * 10 / 5) * 5


def start_process(sheetName='29 Ambuj S'):
    try:
        # Get data from spreadsheet
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheetName)

        auditDetails["name"] = sheet.acell("A1").value
        auditDetails["followers"] = sheet.acell("C39").value
        auditDetails["title"] = sheet.acell("C40").value
        auditDetails["recommendations"] = sheet.acell("C42").value.strip()
        image_url = sheet.acell("C41").value

        sum = 0
        auditDetails["pb_s"] = scale_and_round(
            float(sheet.acell("D8").value))
        sum += float(sheet.acell("D8").value)
        auditDetails["cs_s"] = scale_and_round(
            float(sheet.acell("D23").value))
        sum += float(sheet.acell("D23").value)
        auditDetails["es_s"] = scale_and_round(
            float(sheet.acell("D35").value))
        sum += float(sheet.acell("D35").value)
        auditDetails["po_s"] = scale_and_round(
            float(sheet.acell("D30").value))
        sum += float(sheet.acell("D30").value)
        auditDetails["ape_s"] = scale_and_round(
            float(sheet.acell("D18").value))
        sum += float(sheet.acell("D18").value)
        auditDetails["os_s"] = scale_and_round(sum / 5)
        auditDetails["ovr"] = get_score_description(auditDetails["os_s"])

        recipient_email = sheet.acell("C43").value
        phone_number = sheet.acell("C42").value

        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        # Create new report
        copy_response = drive_service.files().copy(
            fileId=PRESENTATION_ID,
            body={"name": f"Temp Copy for {auditDetails['name']}"}
        ).execute()
        copy_presentation_id = copy_response["id"]

        requests_body = {
            "requests": []
        }

        # Insert scores
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

        # Insert recommendations
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

        # Insert picture
        requests_body["requests"].append(
            {
                "replaceAllShapesWithImage": {
                    "imageUrl": image_url,
                    "replaceMethod": "CENTER_INSIDE",
                    "containsText": {
                        "text": "[[IMAGE]]",
                        "matchCase": True
                    }
                }
            })

        # Hide irrelevant score bars
        for key in ["pb_s", "cs_s", "es_s", "po_s", "ape_s"]:
            desc = get_score_description(auditDetails[key])
            for scoreKey in scoreBarElements[key].keys():
                if desc.lower() != scoreKey:
                    requests_body["requests"].append({
                        "updatePageElementTransform": {
                            "objectId": scoreBarElements[key][scoreKey],
                            "applyMode": "ABSOLUTE",
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": 5000,
                                "translateY": 5000,
                                "unit": "PT"
                            }
                        }
                    })

        # Update slide
        slides_service.presentations().batchUpdate(
            presentationId=copy_presentation_id,
            body=requests_body
        ).execute()

        # Update report permissions and name
        fileUpdates = {
            'type': 'anyone',
            'role': 'writer',
            'name': f"{auditDetails["name"].replace(" ", "_")}_profile_audit"
        }

        # Create slide
        drive_service.permissions().create(
            fileId=copy_presentation_id,
            body=fileUpdates
        ).execute()

        editor_link = f"https://docs.google.com/presentation/d/{copy_presentation_id}/edit"
        # sheet.update_acell("C49", editor_link)
        return

        # # Show all elements in slide
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

        # Download PDF

        request = drive_service.files().export_media(
            fileId=copy_presentation_id, mimeType='application/pdf')
        pdf_file = request.execute()

        name = auditDetails["name"].replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{name}_{timestamp}.pdf"

        with open(f"pdfs/{filename}", 'wb') as f:
            f.write(pdf_file)

        print(f"PDF downloaded as {filename}")

        # # PDF link
        # pdf_link = f"https://docs.google.com/presentation/d/{
        #     copy_presentation_id}/export/pdf"
        #
        # sheet.update_acell("C45", pdf_link)
        return

        # Send WhatsApp message via Twilio for approval
        twilio_client = Client(TWILIO_SID, TWILIO_AUTH_TOKEN)
        message = twilio_client.messages.create(
            from_=f"{TWILIO_WHATSAPP_NUMBER}",
            to=f"{YOUR_WHATSAPP}",
            body=f"{auditDetails["name"]}'s report is ready: {
                pdf_link}\nIf you wish to make changes to the document, click this link to edit and save the document: {editor_link}\nReply 'yes' to approve and send to {recipient_email}",
            media_url=[pdf_link]
        )
        print("Approval message sent. SID:", message)

        start_time = time.time()
        approved = False
        print("Waiting for approval...")

        curDatetime = datetime.now(timezone.utc)

        # Poll Twilio for replies
        while time.time() - start_time < 300:  # 5 minutes timeout
            messages = twilio_client.messages.list(
                to=f"{TWILIO_WHATSAPP_NUMBER}", limit=5)
            for msg in messages:
                if msg.direction == "inbound" and msg.date_sent > curDatetime and msg.from_ == f"{YOUR_WHATSAPP}":
                    if msg.body.strip().lower() in ["yes", "approve"]:
                        approved = True
                        print("Approval received.")
                        send_email(recipient_email, pdf_link)
                        send_w_message(phone_number, pdf_link)
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


# Send email to client
def send_email(to_email, pdf_url):
    try:
        msg = MIMEText(f"Hi {auditDetails['name']},\n\nAttached is your LinkedIn profile audit. It outlines what's working, and where a few smart tweaks can amplify how you're seen by the right audience.\n\nYou already bring the credibility - this helps make it visible.\n\nLet me know what you think. Happy to walk you through it anytime.\n\nBest,\nAmbuj\n\n{pdf_url}")
        msg["Subject"] = "Your LinkedIn Audit is ready - actionable insights to upgrade your brand"
        msg["From"] = EMAIL_FROM
        msg["To"] = to_email

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, to_email, msg.as_string())

        print(f"Email sent to client's address at {mask_email(to_email)}")

    except Exception as e:
        print(e)


# Send whatsapp message to client
def send_w_message(phone_num, pdf_url):
    try:
        message = client.messages.create(
            from_='whatsapp:+15557957011',
            to=f'whatsapp:+91${phone_num}',
            body='Template message'
        )

        print(f"Message sent to client's whatsapp at {mask_phone(phone_num)}.")

    except Exception as e:
        print(e)


if __name__ == "__main__":
    # start_process()
    send_email("carljason.3012@gmail.com", "https://docs.google.com/presentation/d/1HEpifNk3pTArN2qERbBX75dhOCbgSWLMIjkmGXrHxxU/export/pdf")
    # send_w_message("9942727949", "https://example.com");
