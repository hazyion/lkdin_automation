import os
from email.utils import parsedate_to_datetime
import datetime
import time
import gspread
import requests
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from twilio.rest import Client
import smtplib
from email.mime.text import MIMEText

load_dotenv()

# Load environment variables
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

# Set up Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    G_CREDENTIAL_FILE, scope)
client = gspread.authorize(creds)

sheet = client.open("lkdin_automation_test").sheet1
cell_value = sheet.acell("A1").value

# sheet = client.open_by_key(SPREADSHEET_ID).sheet1


def start_process():
    text = sheet.acell("A1").value
    name = sheet.acell("B1").value
    recipient_email = sheet.acell("C1").value

    # Step 1: Generate PDF with Placid
    headers = {"Authorization": f"Bearer {PLACID_TOKEN}",
               "Content-Type": "application/json"}
    payload = {
        "pages": [
            {
                "template_uuid": TEMPLATE_UUID,
                "layers": {
                    "type": {"text": text}
                }
            }
        ]
    }

    # Bypass pdf creation for now
    # response = requests.post(PLACID_URL, json=payload, headers=headers)
    #
    # if response.status_code != 200:
    #     raise Exception(f"Placid error: {response.status_code} {response.text}")
    #
    # data = response.json()
    # pdf_link = ""
    # while True:
    #     get_response = requests.get(f"{PLACID_URL}/{data["id"]}", headers=headers)
    #     pdf_data = get_response.json()
    #     if pdf_data["status"] != 'queued':
    #         print("Successful")
    #         pdf_link = pdf_data["pdf_url"]
    #         break
    #     else:
    #         print("Report creation pending...")
    #         time.sleep(5)
    #
    # sheet.update_acell("D1", pdf_link)

    pdf_link = sheet.acell("D1").value

    # Step 2: Send WhatsApp message via Twilio
    twilio_client = Client(TWILIO_SID, TWILIO_AUTH_TOKEN)
    message = twilio_client.messages.create(
        from_=f"{TWILIO_WHATSAPP_NUMBER}",
        to=f"{YOUR_WHATSAPP}",
        body=f"Hi {name}, your PDF is ready: {
            pdf_link}\nReply 'yes' to approve and send to {recipient_email}"
    )
    print("Message sent. SID:", message)
    sent_message_sid = message.sid

    start_time = time.time()
    approved = False
    print("Waiting for approval...")

    curDatetime = datetime.datetime.now(datetime.timezone.utc)

    # Step 3: Poll Twilio for replies
    while time.time() - start_time < 300:  # 5 minutes timeout
        messages = twilio_client.messages.list(
            to=f"{TWILIO_WHATSAPP_NUMBER}", limit=5)
        print(messages)
        for msg in messages:
            print('datesenttime ', msg.date_sent)
            # Check if message is AFTER the one we sent
            if msg.direction == "inbound" and msg.date_sent > curDatetime and msg.from_ == f"{YOUR_WHATSAPP}":
                if msg.body.strip().lower() in ["yes", "approve"]:
                    approved = True
                    print("Approval received.")
                    send_email(recipient_email, pdf_link)
                    break
        if approved:
            break
        time.sleep(5)  # poll every 5 seconds

    if not approved:
        print("Approval not received in time.")
        return

def send_email(to_email, pdf_url):
    msg = MIMEText(f"Here is your approved PDF: {pdf_url}")
    msg["Subject"] = "Approved PDF"
    msg["From"] = EMAIL_FROM
    msg["To"] = to_email

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, to_email, msg.as_string())

    print("Email sent to", to_email)


if __name__ == "__main__":
    start_process()
