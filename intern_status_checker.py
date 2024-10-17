import imaplib
import email
from email import policy
from email.parser import BytesParser
from email.header import decode_header
import re
import pandas as pd
import os
def decode_mime_words(text):
    decoded_words = decode_header(text)
    subject_parts = []
    for word, encoding in decoded_words:
        if isinstance(word, bytes):
            subject_parts.append(word.decode(encoding if encoding else 'utf-8'))
        else:
            subject_parts.append(word)
    return ''.join(subject_parts)
def get_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if content_type == "text/plain" and "attachment" not in content_disposition:
                return part.get_payload(decode=True).decode("utf-8")
            elif content_type == "text/html":
                html_body = part.get_payload(decode=True).decode("utf-8")
                return html_body
    else:
        return msg.get_payload(decode=True).decode("utf-8")                                
def process_email(msg):
    subject = decode_mime_words(msg["Subject"])

    email_body = get_email_body(msg)
    company_name = None
    application_status = None
    
    if "Congratulations" in email_body or "offer" in email_body.lower():
        application_status = "Accepted"
        company_name_match = re.search(r"from (.*?)(?:,|\n|\.)", email_body)
        if company_name_match:
            company_name = company_name_match.group(1).strip()
    elif "regret" in email_body or "Unfortunately" in email_body:
        application_status = "Rejected"
        company_name_match = re.search(r"from (.*?)(?:,|\n|\.)", email_body)
        if company_name_match:
            company_name = company_name_match.group(1).strip()
    
    return subject, company_name, application_status                

imap_server = "outlook.office365.com"
imap_port = 993
email_user = "your_email@mavs.uta.edu.com"
email_pass = os.getenv("EMAIL_PASSWORD")

# connect to the server
mail = imaplib.IMAP4_SSL(imap_server, imap_port)
mail.login(email_user, email_pass)

# select the mailbox
mail.select("inbox")

status, messages = mail.search(None, '(SUBJECT "acceptance" OR SUBJECT "rejection")')

email_data = []

for num in messages[0].split():
    status, data = mail.fetch(num, "(RFC882)")
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1], policy=policy.default)
            subject, company_name, application_status = process_email(msg)

            email_data.append({
                "Subject": subject,
                "Company": company_name,
                "Status": application_status
            })
            
df = pd.DataFrame(email_data)
df.to_excel("internship_application_status.xlsx", index=False)

mail.logout()            


