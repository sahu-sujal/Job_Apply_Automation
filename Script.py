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

# Load environment variables
load_dotenv()
EMAIL = os.getenv("EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")

# Read email data from Excel
data = pd.read_excel("data.xlsx")

# Email template
SUBJECT_TEMPLATE = "Application for {job_position} - Sujal Sahu"

BODY_TEMPLATE = """Dear Hiring Manager,

I am excited to apply for the {job_position} at {company_name}. With my expertise in {skills}, I am confident in my ability to contribute to your team.

Please find my resume attached for your review. I would appreciate the opportunity to discuss how I can add value to {company_name}.

Thank you for your time and consideration.

Best regards,
Sujal Sahu
"""

# Configure logging
logging.basicConfig(filename="email_log.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Send emails
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(EMAIL, APP_PASSWORD)

    for index, row in data.iterrows():
        try:
            recipient_email = row.get("RecipientEmail", "").strip()
            job_position = row.get("JobPosition", "").strip()
            company_name = row.get("CompanyName", "").strip()
            skills = row.get("Skills", "").strip()
            resume_option = row.get("ResumeOption", "").strip()
            
            if not recipient_email or not resume_option:
                logging.warning(f"Skipping row {index}: Missing required data")
                continue

            # Personalize subject and body
            subject = SUBJECT_TEMPLATE.format(job_position=job_position)
            body = BODY_TEMPLATE.format(job_position=job_position, company_name=company_name, skills=skills)
            
            # Compose the email
            msg = MIMEMultipart()
            msg["From"] = EMAIL
            msg["To"] = recipient_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))
            
            # Attach the selected resume
            if os.path.exists(resume_option):
                with open(resume_option, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(resume_option)}")
                msg.attach(part)
            else:
                logging.warning(f"Attachment not found: {resume_option}")
                continue

            # Retry mechanism for sending emails
            MAX_RETRIES = 3
            for attempt in range(MAX_RETRIES):
                try:
                    server.sendmail(EMAIL, recipient_email, msg.as_string())
                    logging.info(f"Email sent to {recipient_email}")
                    print(f"Email sent to {recipient_email}")
                    break
                except smtplib.SMTPException as e:
                    logging.error(f"Attempt {attempt + 1} failed for {recipient_email}: {e}")
                    time.sleep(5)  # Wait before retrying
            
            # Random delay to prevent spam detection
            time.sleep(random.uniform(2, 5))
        
        except Exception as e:
            logging.error(f"Error sending to {recipient_email}: {e}")

print("Email sending process completed.")
