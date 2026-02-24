import os
import unittest
from unittest.mock import patch, MagicMock
from mailer import EmailSender
import pandas as pd

class TestEmailSender(unittest.TestCase):
    def setUp(self):
        self.sender = EmailSender(smtp_email="test@gmail.com", smtp_password="pwd")
        self.db_path = "test_db.csv"
        self.resume_path = "test_resume.pdf"
        self.template_text = "Subject: Application\n\nDear {Professor_Name},\nI like {Research_Interest}."

    @patch("mailer.smtplib.SMTP_SSL")
    def test_processing(self, mock_smtp):
        # We will stop the sender after the first couple of items so we don't sleep forever
        # Actually, let's patch time.sleep
        with patch('mailer.time.sleep', return_value=None):
            # We want to trace messages sent
            mock_server = MagicMock()
            mock_smtp.return_value = mock_server

            logs = []
            def progress(msg):
                logs.append(msg)
            
            self.sender.process_and_send(self.db_path, self.template_text, self.resume_path, progress)

            # Assert SMTP login was called
            mock_server.login.assert_called_with("test@gmail.com", "pwd")

            # Check emails sent
            self.assertEqual(mock_server.send_message.call_count, 2)
            
            # Check the generated CSV log
            self.assertTrue(os.path.exists("status_log.csv"))
            log_df = pd.read_csv("status_log.csv")
            
            # 4 entries in db, 1 has invalid email, 1 has null email
            self.assertEqual(len(log_df), 4) # "Dr. Null Email" has NaN, let's see what validation did
            # Wait, 4 entries? Let's check test_db
            
            # The emails should have templated content in MIME
            args, kwargs = mock_server.send_message.call_args_list[0]
            msg = args[0]
            self.assertEqual(msg['Subject'], 'Application')
            self.assertEqual(msg['To'], 'amit@example.com')
            
            payloads = msg.get_payload()
            # MIMEText part
            text_part = payloads[0]
            self.assertIn("Dr. Amit Shah", text_part.get_payload())
            self.assertIn("Embedded Systems", text_part.get_payload())

if __name__ == '__main__':
    unittest.main()
