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
        self.sent_log_path = "sent_log.csv"
        self.is_stopped = False
        
        # Ensure log file exists for reading
        if not os.path.exists(self.sent_log_path):
            pd.DataFrame(columns=['Email', 'Timestamp']).to_csv(self.sent_log_path, index=False)

    def validate_email(self, email):
        regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return re.match(regex, str(email)) is not None

    def get_already_sent_emails(self):
        try:
            if os.path.exists(self.sent_log_path):
                df = pd.read_csv(self.sent_log_path)
                # Ensure the Email column exists and converting it to string, stripping handles NaN floats safely
                return set((df['Email'].astype(str).str.strip()).tolist())
        except Exception:
            return set()
        return set()

    def generate_preview(self, row, template_text):
        """Generates the substituted template for UI preview."""
        body = template_text
        
        # Match variables and fallbacks like {Research_Interest|your impressive work} or {Professor_Name}
        matches = re.finditer(r"\{([^}]+)\}", body)
        
        # We process matches from end to beginning so we don't mess up indices
        # Actually a simple replace is difficult if we have multiple identical placeholders, 
        # so let's do a replace based on groups.
        def repl(match):
            content = match.group(1)
            parts = content.split('|', 1)
            col_name = parts[0].strip()
            fallback = parts[1].strip() if len(parts) > 1 else ""
            
            val = str(row.get(col_name, ""))
            if val == 'nan' or not val.strip():
                val = fallback
            return val

        body = re.sub(r"\{([^}]+)\}", repl, body)

        # Identify subject
        subject = "Internship Inquiry"
        lines = body.split('\n')
        if lines and lines[0].lower().startswith("subject:"):
            subject = lines[0][8:].strip()
            body = '\n'.join(lines[1:]).strip()
            
        return subject, body

    def process_and_send(self, database_path, template_text, resume_path, is_html=False, progress_callback=None):
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

        # Prepare logs
        logs = []
        sent_emails_this_run = []
        already_sent = self.get_already_sent_emails()

        # Check if we have anything to send first
        emails_to_send = 0
        for index, row in df.iterrows():
            target_email = str(row.get('Email', '')).strip()
            if target_email not in already_sent and self.validate_email(target_email):
                emails_to_send += 1
                
        if emails_to_send == 0:
            if progress_callback: progress_callback("No new valid emails to send (all already sent or invalid).")
            return

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

        session_sent_count = 0

        for index, row in df.iterrows():
            if self.is_stopped:
                if progress_callback: progress_callback("Process stopped by user.")
                break
            
            target_email = str(row.get('Email', '')).strip()
            
            # Validation & Avoid Duplicates
            if not self.validate_email(target_email):
                logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': 'Invalid Email Format'})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Skipped Invalid Email.")
                continue
                
            if target_email in already_sent:
                logs.append({'Email': target_email, 'Status': 'Skipped', 'Reason': 'Already in sent_log.csv'})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Skipped Duplicate: {target_email}")
                continue

            # Generate dynamic email body using preview generator
            subject, body = self.generate_preview(row, template_text)
            
            msg = MIMEMultipart()
            msg['From'] = formataddr((self.smtp_email.split('@')[0], self.smtp_email))
            msg['To'] = target_email
            msg['Subject'] = subject
            
            if is_html:
                msg.attach(MIMEText(body, 'html'))
            else:
                msg.attach(MIMEText(body, 'plain'))

            # Attach Resume with fixed name format
            if resume_path and os.path.exists(resume_path):
                try:
                    with open(resume_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name="Resume.pdf")
                    part['Content-Disposition'] = 'attachment; filename="Resume.pdf"'
                    msg.attach(part)
                except Exception as e:
                    logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': f'Attachment Error: {str(e)}'})
                    continue

            # Send Email
            try:
                server.send_message(msg)
                logs.append({'Email': target_email, 'Status': 'Sent', 'Reason': ''})
                sent_emails_this_run.append({'Email': target_email, 'Timestamp': time.strftime("%Y-%m-%d %H:%M:%S")})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Sent to {target_email}")
                session_sent_count += 1
            except Exception as e:
                logs.append({'Email': target_email, 'Status': 'Failed', 'Reason': f'Send Error: {str(e)}'})
                if progress_callback: progress_callback(f"[{index+1}/{total_emails}] Failed {target_email}: {str(e)}")
                continue # If send fails, don't sleep the normal time

            # Manage delays if we just successfully sent and have more left
            if index < total_emails - 1 and not self.is_stopped:
                # Check for batch limit pause
                if session_sent_count > 0 and session_sent_count % 10 == 0:
                    delay = 600 # 10 minutes
                    if progress_callback: progress_callback("Batch limit reached (10 emails). Pausing for 10 minutes...")
                else:
                    delay = random.randint(45, 120)
                    if progress_callback: progress_callback(f"Waiting for {delay} seconds to prevent spam...")
                
                # Sleep in small chunks to allow stopping
                for _ in range(delay):
                    if self.is_stopped:
                        break
                    time.sleep(1)

        server.quit()

        # Update log files
        log_df = pd.DataFrame(logs)
        if not log_df.empty:
            log_file_mode = 'a' if os.path.exists(self.status_log_path) else 'w'
            header = not os.path.exists(self.status_log_path)
            log_df.to_csv(self.status_log_path, mode=log_file_mode, header=header, index=False)
            
        sent_df = pd.DataFrame(sent_emails_this_run)
        if not sent_df.empty:
            sent_df.to_csv(self.sent_log_path, mode='a', header=False, index=False)
        
        if progress_callback: progress_callback("Process Finished. Logs saved.")

    def stop(self):
        self.is_stopped = True
