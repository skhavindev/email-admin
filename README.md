# email-admin

Write a Python script that does the following:
Reads an Excel file (with editable path) containing two columns: "id" and "contact email".
The "contact email" field may contain multiple email addresses separated by commas.
Scans a parent directory (path should be editable) which contains multiple folders named in the format dd-mm-yy-dd-mm-yy (start date to end date).
For each row in the Excel file, it finds the corresponding PDF file named <id>.pdf inside the folder whose end date (the second dd-mm-yy in the folder name) is the latest.
If the PDF file is found, the script sends it as an email attachment to all the email addresses in the "contact email" field.
Allow SMTP configuration (host, port, sender email, sender password) to be easily edited.
Use only standard libraries and optionally pandas, openpyxl, and smtplib

```python
import pandas as pd
import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import COMMASPACE

def read_excel_file(file_path: str) -> pd.DataFrame:
    """Read Excel file with 'id' and 'contact email' columns"""
    return pd.read_excel(file_path, engine='openpyxl')

def get_end_date_from_folder(folder_name: str) -> datetime:
    """Parse folder name to extract end date (second date in format)"""
    try:
        parts = folder_name.split('-')
        end_date_str = '-'.join(parts[3:6])
        return datetime.strptime(end_date_str, '%d-%m-%y')
    except (ValueError, IndexError):
        return None

def find_latest_folder(parent_dir: str) -> str:
    """Find folder with most recent end date in directory"""
    latest_date = None
    latest_folder = None
    for folder_name in os.listdir(parent_dir):
        folder_path = os.path.join(parent_dir, folder_name)
        if os.path.isdir(folder_path):
            end_date = get_end_date_from_folder(folder_name)
            if end_date and (not latest_date or end_date > latest_date):
                latest_date = end_date
                latest_folder = folder_path
    return latest_folder

def find_pdf_file(parent_dir: str, file_id: str) -> str:
    """Locate PDF by ID in latest dated folder"""
    latest_folder = find_latest_folder(parent_dir)
    if not latest_folder:
        return None
    pdf_path = os.path.join(latest_folder, f"{file_id}.pdf")
    return pdf_path if os.path.isfile(pdf_path) else None

def send_email_with_attachment(
    smtp_host: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    recipient_emails: list,
    subject: str,
    body: str,
    attachment_path: str
) -> None:
    """Send email with PDF attachment"""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = COMMASPACE.join(recipient_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    with open(attachment_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        msg.attach(part)
    
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

def process_excel_and_send_emails(
    excel_path: str,
    parent_dir: str,
    smtp_host: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str
) -> None:
    """Main processing function"""
    df = read_excel_file(excel_path)
    
    for _, row in df.iterrows():
        file_id = str(row['id'])
        emails = [e.strip() for e in str(row['contact email']).split(',') if e.strip()]
        
        if not emails:
            print(f"Skipping ID {file_id}: No valid emails")
            continue
        
        pdf_path = find_pdf_file(parent_dir, file_id)
        if not pdf_path:
            print(f"PDF not found for ID {file_id}")
            continue
        
        send_email_with_attachment(
            smtp_host,
            smtp_port,
            sender_email,
            sender_password,
            emails,
            f"Document {file_id}",
            "Please find the attached document.",
            pdf_path
        )
        print(f"Sent email for ID {file_id} to {emails}")

# Example Usage:
# process_excel_and_send_emails(
#     excel_path='path/to/contacts.xlsx',
#     parent_dir='path/to/parent_directory',
#     smtp_host='smtp.example.com',
#     smtp_port=587,
#     sender_email='your_email@example.com',
#     sender_password='your_password'
# )
```

### Key Features:
1. **Excel Processing**:
   - Reads Excel files with `id` and `contact email` columns
   - Handles multiple comma-separated emails per row

2. **PDF Location**:
   - Automatically finds the folder with latest end date (second date in `dd-mm-yy-dd-mm-yy` format)
   - Locates PDFs using `id` values

3. **Email Functionality**:
   - Configurable SMTP settings (host, port, credentials)
   - Sends PDF attachments to all specified emails
   - TLS encryption for security

4. **Error Handling**:
   - Skips rows with missing PDFs or invalid emails
   - Graceful handling of malformed folder names

### Usage Instructions:
1. Install required packages:
   ```bash
   pip install pandas openpyxl
   ```

2. Configure these parameters in the `process_excel_and_send_emails()` call:
   - `excel_path`: Path to your Excel file
   - `parent_dir`: Directory containing dated folders
   - SMTP credentials for your email provider

3. Run the script:
   ```python
   process_excel_and_send_emails(
       excel_path='contacts.xlsx',
       parent_dir='reports/',
       smtp_host='smtp.gmail.com',
       smtp_port=587,
       sender_email='your@gmail.com',
       sender_password='your_app_password'
   )
   ```

### Notes:
- The script uses pandas for Excel handling and smtplib for email
- Folder date parsing expects `dd-mm-yy-dd-mm-yy` format
- For Gmail, enable "Less secure apps" or use app-specific passwords
- Test with small datasets before full deployment
