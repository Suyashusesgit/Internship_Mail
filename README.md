# Internship Outreach Suite (Pro-Level)

A comprehensive, Python-based desktop application designed to streamline and automate personalized internship inquiry emails to professors. It features a modern GUI, smart templating with fallbacks, robust anti-spam measures, and automatic duplicate prevention.

## Features

- **Modern UI**: Polished graphical interface built with `customtkinter`.
- **Smart Templating**: Use placeholders like `{Professor_Name}` or `{Research_Interest}` that automatically map to columns in your CSV/Excel file.
- **Dynamic Fallbacks**: Support for fallbacks when a CSV field is empty, e.g., `{Research_Topic|your impressive work}`.
- **Multiple Profiles**: Create, save, and load different email templates (e.g., "Embedded Systems" vs "Rail Engineering") through the UI.
- **Anti-Spam Human Touch**: 
  - Random `45` to `120` second variable delays between emails.
  - Automatic 10-minute pauses every 10 emails outbox to emulate human-like behavior.
- **Duplicate Prevention**: Keeps a local `sent_log.csv` to ensure you never accidentally double-email a professor from your list.
- **Preview Mode**: Visually verify how your variables and fallbacks map out on a sample row before executing the bulk dispatch.
- **HTML Support**: Options to send formatted rich-text emails.
- **Clean PDF Attachments**: Upload any local PDF resume, and the sender engine will automatically sanitize its attachment name to `Resume.pdf` to the recipient.

## Prerequisites

- Python 3.10+
- An email account (Gmail/Outlook) with an **App Password** configured. Do not use your standard login password.

## Setup Instructions

1. **Clone the repository or download the source code.**
2. **Navigate to the project directory** in your terminal:
   ```bash
   cd /path/to/Internship_Mail
   ```
3. **Create a virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate
   ```
4. **Install the dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start the application:
   ```bash
   python main.py
   ```
2. **Select your Database**: Load your Excel or CSV file. Ensure it contains a column named `Email`.
3. **Select your Resume**: Choose the PDF file you wish to attach.
4. **Enter SMTP Credentials**: Provide your Gmail address and the **App Password**.
5. **Configure Template**:
   - Write your email content. Use column headers wrapped in brackets like `{Department}`.
   - Use the `|` syntax for fallbacks if a cell might be empty: `{Department|your department}`.
   - Save the template profile for future use.
6. **Preview**: Click `Preview First Match` to see what the generated email will look like before sending.
7. **Start**: Click `START SEQUENCE` and let the automation run reliably in the background.

## Important Note on Spam & Limits

When using personal Google Accounts, avoid sending more than 40-50 automated emails per day. The variable timeouts baked into the software will handle the cadence, but ensure your outbox volumes remain organic to preserve your account standing.

## Logs

- `status_log.csv`: Tracks the success/failure state of every email attempted in a session.
- `sent_log.csv`: A permanent record of every email address receiving a successful delivery to prevent duplicates.
