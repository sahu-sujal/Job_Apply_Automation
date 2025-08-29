import smtplib
import pandas as pd
import time
import os
import random
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import requests

# Skills template
SKILLS = """CI/CD, Jenkins, Docker, Bash Scripting, Linux, AWS, Burpsuite, OWASP, Nmap, Wireshark, Vulnerability Assessment, Penetration Testing, Web Application Security, Network Security"""

# Resume path
RESUME_PATH = "Sujal_CV.pdf"

# Google Sheet Configuration
SHEET_ID = '1dS5_pWGGrxfXY5jdIm-WjcqC76pCSXdd9RXt3gg7bZw'
SHEET_NAME = 'Sheet1'
URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'

def get_google_sheet_data():
    try:
        data = pd.read_csv(URL)
        # Ensure all required columns exist
        required_columns = ['RecipientEmail', 'JobPosition', 'CompanyName', 'Skills', 'Applied']
        if not all(col in data.columns for col in required_columns):
            raise Exception('Missing required columns in the Google Sheet.')
        return data
    except Exception as e:
        raise Exception(f'Failed to fetch data from Google Sheet: {e}')

def update_application_status(sheet_id, email):
    try:
        # The Google Sheet update URL using Google Apps Script web app
        update_url = "https://script.google.com/macros/s/AKfycbw_cQMKasHY34eFd6v5bIVyIiFXksUoI5uRC3AqQWGnqd6BGpF92KOu-EOPF7DvYUSXtg/exec"
        payload = {
            'sheetId': sheet_id,
            'email': email,
            'status': USER_ID  
        }
        
        logging.info(f"Attempting to update sheet for email: {email} with USER_ID: {USER_ID}")
        print(f"Sending update request to Google Sheet for {email}")
        
        response = requests.post(update_url, json=payload)
        logging.info(f"Update response status: {response.status_code}")
        logging.info(f"Update response content: {response.text}")
        
        if response.status_code != 200:
            raise Exception(f"Failed to update status: {response.text}")
        
        print(f"Successfully updated application status for {email}")
        logging.info(f"Updated application status for {email} by user {USER_ID}")
    except Exception as e:
        print(f"Error updating application status: {str(e)}")
        logging.error(f"Failed to update application status for {email}: {e}")

# Configure logging
logging.basicConfig(filename="email_log.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Load environment variables
load_dotenv()
EMAIL = os.getenv("EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
USER_ID = os.getenv("USER_ID")

if not all([EMAIL, APP_PASSWORD, USER_ID]):
    logging.error("Email, APP_PASSWORD, or USER_ID not found in environment variables")
    exit(1)

# Read data from Google Sheet
try:
    data = get_google_sheet_data()
except Exception as e:
    logging.error(f"Failed to fetch data from Google Sheet: {e}")
    exit(1)

# Email template
SUBJECT_TEMPLATE = "Application for {job_position} - Sujal Sahu"

# Rate limit for emails
RATE_LIMIT = 30
email_counter = 0

BODY_TEMPLATE = """Dear Hiring Manager,

I am excited to apply for the {job_position} at {company_name}. With my expertise in {skills}, I am confident in my ability to contribute to your team.

Please find my resume attached for your review. I would appreciate the opportunity to discuss how I can add value to {company_name}.

Thank you for your time and consideration.

Best regards,
Sujal Sahu
"""

# Send emails
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(EMAIL, APP_PASSWORD)

    for index, row in data.iterrows():
        try:
            # Skip if current user has already applied
            applied_status = str(row.get('Applied', '')).strip()
            if pd.notna(applied_status) and USER_ID in applied_status.split(','):
                logging.info(f"Skipping {row['CompanyName']}: Already applied by {USER_ID}")
                continue

            recipient_email = row["RecipientEmail"].strip()
            job_position = row["JobPosition"].strip()
            company_name = row["CompanyName"].strip()
            
            if not recipient_email:
                logging.warning(f"Skipping row {index}: Missing required data")
                continue

            # Personalize subject and body
            subject = SUBJECT_TEMPLATE.format(job_position=job_position)
            body = BODY_TEMPLATE.format(job_position=job_position, company_name=company_name, skills=SKILLS)
            
            # Compose the email
            msg = MIMEMultipart()
            msg["From"] = EMAIL
            msg["To"] = recipient_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))
            
            # Attach resume
            if os.path.exists(RESUME_PATH):
                with open(RESUME_PATH, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                filename = os.path.basename(RESUME_PATH)
                part.add_header("Content-Disposition", f"attachment; filename={filename}")
                msg.attach(part)
            else:
                logging.warning(f"Resume not found: {RESUME_PATH}")
                continue

            # Send email with retry mechanism
            MAX_RETRIES = 3
            success = False
            for attempt in range(MAX_RETRIES):
                try:
                    server.sendmail(EMAIL, recipient_email, msg.as_string())
                    logging.info(f"Email sent to {recipient_email}")
                    print(f"Email sent to {recipient_email}")
                    success = True
                    break
                except smtplib.SMTPException as e:
                    logging.error(f"Attempt {attempt + 1} failed for {recipient_email}: {e}")
                    time.sleep(5)  # Wait before retrying
            
            # Update application status if email was sent successfully
            if success:
                update_application_status(SHEET_ID, recipient_email)
                email_counter += 1
                
                # Check if rate limit is reached
                if email_counter >= RATE_LIMIT:
                    print(f"\nRate limit of {RATE_LIMIT} emails reached. Stopping the process.")
                    logging.info(f"Rate limit of {RATE_LIMIT} emails reached. Process stopped.")
                    break
              # Random delay to prevent spam detection (80-150 seconds)
            delay_time = random.uniform(90, 150)
            print(f"Waiting for {delay_time:.1f} seconds before sending next email...")
            time.sleep(delay_time)
        
        except Exception as e:
            logging.error(f"Error sending to {recipient_email}: {e}")

print("Email sending process completed.")
