import os
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage
import pandas as pd

# Load environment variables from .env file
load_dotenv()

EMAIL = os.getenv('EMAIL_ID')
PASSWORD = os.getenv('EMAIL_PASSWORD')

def send_email(to_email, first_name, medicine, batch_no, expiry, po_header):
    subject = "Your Subject"
    body = f"""Dear {first_name},

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
        # Read the Excel file, assuming headers are in the fourth row (index 3)
        df = pd.read_excel(file_path, header=3)
        
        # Read the header values from the first two rows
        header_df = pd.read_excel(file_path, header=None, nrows=2)
        po_header_row1 = header_df.iloc[0, :].astype(str).str.cat(sep=' ')
        po_header_row2 = header_df.iloc[1, :].astype(str).str.cat(sep=' ')
        po_header = f"{po_header_row1} {po_header_row2}"
        
        # Print the columns to debug
        print("Columns in the Excel file:", df.columns)
        
        for _, row in df.iterrows():
            send_email(
                to_email=row['Email'],
                first_name=row['Contact Person Name'],
                medicine=row['Description'],
                batch_no=row['Batch No'],
                expiry=row['Expiry'],
                po_header=po_header  # Use the combined header value
            )
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Send emails based on Excel data.")
    parser.add_argument("file_path", help="Path to the Excel file.")
    args = parser.parse_args()
    process_excel(args.file_path)