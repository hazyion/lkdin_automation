import boto3
from botocore.exceptions import ClientError
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

AWS_REGION = os.getenv("AWS_REGION")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")
TWILIO_SID = os.getenv("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.getenv("TWILIO_AUTH_TOKEN")
TWILIO_WHATSAPP_NUMBER = os.getenv("TWILIO_WHATSAPP_NUMBER")
BUSINESS_WHATSAPP_NUMBER = os.getenv("BUSINESS_WHATSAPP_NUMBER")
AMBUJ_WHATSAPP = os.getenv("AMBUJ_WHATSAPP")
JASON_WHATSAPP = os.getenv("JASON_WHATSAPP")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CLIENT_MESSAGE_SID = os.getenv("CLIENT_MESSAGE_SID")
APPROVAL_MESSAGE_SID = os.getenv("APPROVAL_MESSAGE_SID")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
G_CREDENTIAL_FILE = os.getenv("G_CREDENTIAL_FILE")
PRESENTATION_ID = os.getenv("PRESENTATION_ID")

CHARSET = "UTF-8"

# Create SES client
aws_client = boto3.client(
    'ses',
    region_name=AWS_REGION,
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY
)

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
    "name": "",
    "title": "",
    "followers": 0,
    "pb_s": 0,
    "cs_s": 0,
    "es_s": 0,
    "po_s": 0,
    "ape_s": 0,
    "recommendations": "",
    "os_s": 0,
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
        print('Process started')
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
        phone_number = sheet.acell("C44").value
        print('Extracted data from google sheet')

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

        # PDF link
        editor_link = f"https://docs.google.com/presentation/d/{copy_presentation_id}/edit"
        sheet.update_acell("C46", editor_link)
        pdf_link = f"https://docs.google.com/presentation/d/{copy_presentation_id}/export/pdf"
        sheet.update_acell("C45", pdf_link)
        print('Created new report in google slides')

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

        # # Download PDF
        #
        # request = drive_service.files().export_media(
        #     fileId=copy_presentation_id, mimeType='application/pdf')
        # pdf_file = request.execute()
        #
        # name = auditDetails["name"].replace(" ", "_")
        # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # filename = f"{name}_{timestamp}.pdf"
        #
        # with open(f"pdfs/{filename}", 'wb') as f:
        #     f.write(pdf_file)
        #
        # print(f"PDF downloaded as {filename}")

        # Send WhatsApp message via Twilio for approval
        twilio_client = Client(TWILIO_SID, TWILIO_AUTH_TOKEN)
        message = twilio_client.messages.create(
            from_=f"{BUSINESS_WHATSAPP_NUMBER}",
            to=f"{JASON_WHATSAPP}",
            content_variables = json.dumps({
                "1": auditDetails["name"],
                "2": pdf_link,
                "3": editor_link,
                "4": recipient_email
                }),
            content_sid = APPROVAL_MESSAGE_SID
        )
        print("Approval message sent.")

        start_time = time.time()
        approved = False
        print("Waiting for approval...")

        curDatetime = datetime.now(timezone.utc)

        # Poll Twilio for replies
        while time.time() - start_time < 300:  # 5 minutes timeout
            messages = twilio_client.messages.list(
                to=f"{BUSINESS_WHATSAPP_NUMBER}", limit=5)
            for msg in messages:
                if msg.direction == "inbound" and msg.date_sent > curDatetime and msg.from_ == f"{JASON_WHATSAPP}":
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
    BODY_HTML = f"""
    <html>
    <head><head>
    <body>
        <h4>Hi {auditDetails['name']},</h4>
        <p>Attached is your LinkedIn profile audit. It outlines what's working, and where a few smart tweaks can amplify how you're seen by the right audience.</p>
        <p>You already bring the credibility - this helps make it visible.</p>
        <p>Let me know what you think. Happy to walk you through it anytime.</p>
        <br/>
        <h4>Best,</h4>
        <h4>Ambuj</h4>
        <br/>
        <a target="_blank" href="{pdf_url}">LinkedIn profile audit report</a>
    </body>
    </html>
    """
    BODY_TEXT = f"""
    Hi {auditDetails['name']},\n\nAttached is your LinkedIn profile audit. It outlines what's working, and where a few smart tweaks can amplify how you're seen by the right audience.\n\nYou already bring the credibility - this helps make it visible.\n\nLet me know what you think. Happy to walk you through it anytime.\n\nBest,\nAmbuj\n\n{pdf_url}
    """
    SUBJECT = "Your LinkedIn Audit is ready - actionable insights to upgrade your brand"
    try:
        response = aws_client.send_email(
            Destination={'ToAddresses': [to_email]},
            Message={
                'Body': {
                    'Html': {'Charset': CHARSET, 'Data': BODY_HTML},
                    'Text': {'Charset': CHARSET, 'Data': BODY_TEXT},
                },
                'Subject': {'Charset': CHARSET, 'Data': SUBJECT},
            },
            Source=EMAIL_SENDER
        )
    except ClientError as e:
        print(e.response['Error']['Message'])
    else:
        print(f"Email sent! Message ID: {response['MessageId']}")


# Send whatsapp message to client
def send_w_message(phone_num, pdf_link):
    try:
        client = Client(TWILIO_SID, TWILIO_AUTH_TOKEN)
        message = client.messages.create(
            from_='whatsapp:+15557957011',
            to=f'whatsapp:+91{phone_num}',
            content_variables = json.dumps({
                "1": auditDetails["name"],
                "2": pdf_link
                }),
            content_sid = CLIENT_MESSAGE_SID
        )
        print(f"Message sent to client's whatsapp at {mask_phone(phone_num)}.")

    except Exception as e:
        print(e)


if __name__ == "__main__":
    # start_process()
    # send_email("carljason.3012@gmail.com", "https://docs.google.com/presentation/d/1HEpifNk3pTArN2qERbBX75dhOCbgSWLMIjkmGXrHxxU/export/pdf")
    # send_w_message("9942727949", "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf");
    pass
