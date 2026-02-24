import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
import pandas as pd
import time
import random
import os
import re

class EmailSender:
    def __init__(self, smtp_email, smtp_password, smtp_server="smtp.gmail.com", smtp_port=465, is_ssl=True):
        self.smtp_email = smtp_email
        self.smtp_password = smtp_password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.is_ssl = is_ssl
        self.status_log_path = "status_log.csv"
        self.is_stopped = False

    def validate_email(self, email):
        regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return re.match(regex, str(email)) is not None

    def process_and_send(self, database_path, template_text, resume_path, progress_callback=None):
        self.is_stopped = False
        try:
            if database_path.endswith('.csv'):
                df = pd.read_csv(database_path)
            else:
                df = pd.read_excel(database_path)
        except Exception as e:
            if progress_callback: progress_callback(f"Error loading database: {str(e)}")
            return

        total_emails = len(df)
        if total_emails == 0:
            if progress_callback: progress_callback("Database is empty.")
            return

        # Prepare log list
        logs = []

        try:
            if self.is_ssl:
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            else:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.starttls()
            
            server.login(self.smtp_email, self.smtp_password)
            if progress_callback: progress_callback("SMTP Login Successful.")
        except Exception as e:
            if progress_callback: progress_callback(f"SMTP Login Failed: {str(e)}")
            return

        for index, row in df.iterrows():
            if self.is_stopped:
                if progress_callback: progress_callback("Process stopped by user.")
                break
            
            target_email = str(row.get('Email', '')).strip()
            if not self.validate_email(target_email):
                logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': 'Invalid Email Format'})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Skipped Invalid Email: {target_email}")
                continue

            # Generate dynamic email body
            body = template_text
            for col in df.columns:
                placeholder = "{" + str(col) + "}"
                if placeholder in body:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    body = body.replace(placeholder, val)
            
            # Identify subject and body. Simple assumption: First line is Subject, rest is body.
            # However, usually users just provide body, or we can look for "Subject: ..."
            # Let's extract Subject if it starts with "Subject:"
            subject = "Internship Inquiry"
            lines = body.split('\n')
            if lines[0].lower().startswith("subject:"):
                subject = lines[0][8:].strip()
                body = '\n'.join(lines[1:]).strip()
            
            msg = MIMEMultipart()
            msg['From'] = formataddr((self.smtp_email.split('@')[0], self.smtp_email))
            msg['To'] = target_email
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))

            # Attach Resume
            if resume_path and os.path.exists(resume_path):
                try:
                    with open(resume_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(resume_path))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(resume_path)}"'
                    msg.attach(part)
                except Exception as e:
                    logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': f'Attachment Error: {str(e)}'})
                    continue

            # Send Email
            try:
                server.send_message(msg)
                logs.append({'Email': target_email, 'Status': 'Sent', 'Reason': ''})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Sent to {target_email}")
            except Exception as e:
                logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': f'Send Error: {str(e)}'})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Failed {target_email}: {str(e)}")

            # Random delay if not the last email
            if index < total_emails - 1 and not self.is_stopped:
                delay = random.randint(30, 90)
                if progress_callback: progress_callback(f"Waiting for {delay} seconds to prevent spam...")
                
                # Sleep in small chunks to allow stopping
                for _ in range(delay):
                    if self.is_stopped:
                        break
                    time.sleep(1)

        server.quit()

        # Save logs
        log_df = pd.DataFrame(logs)
        log_file_mode = 'a' if os.path.exists(self.status_log_path) else 'w'
        header = not os.path.exists(self.status_log_path)
        log_df.to_csv(self.status_log_path, mode=log_file_mode, header=header, index=False)
        
        if progress_callback: progress_callback("Process Finished. Logs saved to status_log.csv")

    def stop(self):
        self.is_stopped = True
