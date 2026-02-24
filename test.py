import os
import unittest
from unittest.mock import patch, MagicMock
from mailer import EmailSender
import pandas as pd

class TestProEmailSender(unittest.TestCase):
    def setUp(self):
        self.sender = EmailSender(smtp_email="test@gmail.com", smtp_password="pwd")
        self.db_path = "test_db.csv"
        self.resume_path = "test_resume.pdf"
        self.template_text = "Subject: Application\n\nDear {Professor_Name},\nI like {Research_Interest|your lab}."
        
        # Reset sent logs
        if os.path.exists("sent_log.csv"):
            os.remove("sent_log.csv")
        if os.path.exists("status_log.csv"):
            os.remove("status_log.csv")
            
        # Re-init since we deleted the file
        self.sender = EmailSender(smtp_email="test@gmail.com", smtp_password="pwd")

    def test_preview_generator_with_fallback(self):
        # Dr. Amit Shah has "Embedded Systems"
        row1 = pd.Series({"Professor_Name": "Dr. Amit Shah", "Research_Interest": "Embedded Systems"})
        subj, body = self.sender.generate_preview(row1, self.template_text)
        self.assertIn("Embedded Systems", body)
        self.assertEqual(subj, "Application")
        
        # Someone with no research interest
        row2 = pd.Series({"Professor_Name": "Dr. Blank", "Research_Interest": float('nan')})
        subj2, body2 = self.sender.generate_preview(row2, self.template_text)
        self.assertIn("your lab", body2) # fallback string

    @patch("mailer.smtplib.SMTP_SSL")
    def test_processing_skips_duplicates(self, mock_smtp):
        # We manually seed the sent_log.csv so amit@example.com is skipped
        pd.DataFrame([{"Email": "amit@example.com", "Timestamp": "2024-01-01"}]).to_csv("sent_log.csv", index=False)
        
        with patch('mailer.time.sleep', return_value=None):
            mock_server = MagicMock()
            mock_smtp.return_value = mock_server
            
            logs = []
            def progress(msg):
                logs.append(msg)
            
            self.sender.process_and_send(self.db_path, self.template_text, self.resume_path, is_html=False, progress_callback=progress)
            
            # test_db has 4 rows. 1 null email, 1 invalid email. "amit" is already sent.
            # Thus, only "priya@example.in" should be sent.
            self.assertEqual(mock_server.send_message.call_count, 1)
            
            # Let's inspect status_log.csv
            status_df = pd.read_csv("status_log.csv")
            
            # We expect 3 logged rows (the invalid, the duplicate skipped, and the successful send)
            # The null email gets skipped entirely before hitting loop because dataframe iterrows but validation catches it as invalid.
            self.assertEqual(len(status_df), 4)
            
            skipped_row = status_df[status_df['Email'] == "amit@example.com"].iloc[0]
            self.assertEqual(skipped_row['Status'], "Skipped")

if __name__ == '__main__':
    unittest.main()
