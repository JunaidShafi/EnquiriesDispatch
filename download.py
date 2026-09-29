from playwright.sync_api import Playwright, sync_playwright
import mimetypes
import os
import smtplib
from email.message import EmailMessage 
from dotenv import load_dotenv
from datetime import datetime
from email.utils import formataddr
load_dotenv()


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://bomis.nascorptechnologies.com/Index")
    page.get_by_role("textbox", name="User Name").fill("junaidshafi@bomiskashmir.edu.in")
    page.get_by_role("textbox", name="Password").fill("Crap1234")
    page.get_by_role("button", name="Login").click()
    page.get_by_role("listitem").filter(has_text="Academic Setup Achievement/").get_by_role("link").click()
    page.get_by_role("link", name=" Admission Management").click()
    page.locator("a").filter(has_text="Registrations").click()
    page.get_by_role("button", name=" ").click()
    with page.expect_download() as download_info:
        with page.expect_popup() as page1_info:
            page.get_by_text("Export to excel").click()
        page1 = page1_info.value
    download = download_info.value
    download.save_as("enq.xls")

    page1.close()

    # ---------------------


with sync_playwright() as playwright:
    run(playwright)


def send_file(filepath):
    try:
        print("Started Sending Mail")
        SMTP_SERVER = 'smtp.gmail.com'
        SMTP_PORT = 465  # Use 465 for SSL or 587 for TLS
        GMAIL_USER = os.getenv("GMAIL_USER")
        GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD") # 16-character App Password

        msg = EmailMessage()
        msg['Subject'] = 'Daily Enquiry List'
        msg['From'] = formataddr(("Junaid Shafi",GMAIL_USER))
        msg['To'] = os.getenv("RECEIVER_EMAIL")
        msg.set_content(f'Hello Sleeping people here are the registration details on {datetime.now().date()}')

        file_path = filepath
        mime_type, _ = mimetypes.guess_type(file_path)
        main_type, sub_type = (mime_type or 'application/octet-stream').split('/', 1)

        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype=main_type, subtype=sub_type, filename=os.path.basename(file_path))

        # 3. Connect to Gmail and send
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.send_message(msg)
            print('Email successfully sent!')

    except smtplib.SMTPException as e:
        print(f"General Error Occured {e}")

send_file("enq.xls")