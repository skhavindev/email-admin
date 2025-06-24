# email-admin

```python
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import getpass
from datetime import datetime

# ===== USER CONFIGURATION =====
EXCEL_PATH = "path/to/your/excel_file.xlsx"  # Update with your Excel path
BASE_FOLDER = "path/to/parent/folder"        # Update with parent folder of dated folders
END_DATE = "31/12/25"                        # Update with end date in dd/mm/yy format
SENDER_EMAIL = "your_email@example.com"       # Your email address
# ==============================

def send_email_with_pdf(receiver_email, pdf_path, sender_email, password):
    """Sends email with PDF attachment"""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Your Document"
    
    body = "Please find the attached document."
    msg.attach(MIMEText(body, 'plain'))
    
    with open(pdf_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="pdf")
        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
        msg.attach(attach)
    
    with smtplib.SMTP('smtp.gmail.com', 587) as server:  # For Gmail
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)

def main():
    # Convert end date to folder name format
    end_date_clean = END_DATE.replace("/", "_")
    
    # Find target folder
    target_folder = None
    for folder_name in os.listdir(BASE_FOLDER):
        if folder_name.endswith(end_date_clean):
            target_folder = os.path.join(BASE_FOLDER, folder_name)
            break
    
    if not target_folder:
        print(f"Error: Folder ending with '{END_DATE}' not found in {BASE_FOLDER}")
        return
    
    print(f"Using folder: {target_folder}")
    
    # Read Excel data
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Email password
    password = getpass.getpass("Enter your email password: ")
    
    # Process records
    for index, row in df.iterrows():
        pdf_name = f"{row['id']}.pdf"
        pdf_path = os.path.join(target_folder, pdf_name)
        
        if not os.path.exists(pdf_path):
            print(f"⚠️ PDF not found for ID {row['id']}: {pdf_name}")
            continue
        
        emails = [email.strip() for email in str(row['contact email']).split(';')]
        
        for email in emails:
            try:
                send_email_with_pdf(email, pdf_path, SENDER_EMAIL, password)
                print(f"✅ Sent {pdf_name} to {email}")
            except Exception as e:
                print(f"❌ Failed to send to {email}: {str(e)}")

if __name__ == "__main__":
    main()
```

### Key Features:
1. **Easy Configuration**:
   - Set `EXCEL_PATH`, `BASE_FOLDER`, `END_DATE`, and `SENDER_EMAIL` at the top of the script
   - Password entered securely at runtime

2. **Folder Matching**:
   - Automatically finds folders named `dd_mm_yy-dd_mm_yy`
   - Uses the END_DATE to locate the correct folder

3. **Email Handling**:
   - Supports multiple emails per contact (separated by semicolons)
   - PDFs attached with clear filename
   - Real-time success/failure logging

4. **Error Handling**:
   - Missing PDFs reported with ID
   - Invalid email formats skipped
   - Excel read errors captured

### Usage Instructions:
1. **Prerequisites**:
   ```bash
   pip install pandas openpyxl
   ```
2. **Configure**:
   - Update the 4 configuration variables at the top
   - Ensure Excel columns are named exactly:
     - `id` (matches PDF filenames)
     - `contact email` (semicolon-separated emails)

3. **Run**:
   ```bash
   python script_name.py
   ```
   - Enter email password when prompted

### Notes:
- For Gmail, enable "Less Secure Apps" or use App Password
- Folder names should use underscores (e.g., `01_01_25_31_12_25`)
- Dates must be in `dd/mm/yy` format in configuration
- Excel can be `.xlsx` or `.xls`

To modify paths/date, simply update the configuration variables at the top of the script. The system will automatically locate the correct folder and match PDFs by ID.
