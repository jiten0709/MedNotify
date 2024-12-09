import os
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage
import pandas as pd

# Load environment variables from .env file
load_dotenv()

EMAIL = os.getenv('EMAIL_ID')
PASSWORD = os.getenv('EMAIL_PASSWORD')

def send_email(to_email, name, medicine, batch_no, expiry, po_header):
    subject = f"Medicine Expiry Notification: {medicine}"
    body = f"""Dear {name},

This is to inform you that medicine {medicine} having the batch number of {batch_no} is expiring on {expiry} against PO {po_header}.

You are kindly requested to take necessary action.

Thanks.

Regards,
Vipul Mishra
Mangala Healthcare Services
+91-9620300567

Powered by – Mars System & Solution"""

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL
    msg['To'] = to_email
    msg.set_content(body)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def process_excel(file_path):
    try:
        # Load all sheets in the Excel file
        excel_data = pd.ExcelFile(file_path)
        
        # Iterate through each sheet
        for sheet_name in excel_data.sheet_names:
            print(f"Processing sheet: {sheet_name}")
            po_header = sheet_name
            
            # Read the data starting from the fourth row
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
            
            for _, row in df.iterrows():
                name = row['Contact Person Name'] or row['Name']
                send_email(
                    to_email=row['Email'],
                    name=name,
                    medicine=row['Description'],
                    batch_no=row['Batch No'],
                    expiry=row['Expiry'],
                    po_header=po_header
                )
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Send emails based on Excel data.")
    parser.add_argument("file_path", help="Path to the Excel file.")
    args = parser.parse_args()

    process_excel(args.file_path)
