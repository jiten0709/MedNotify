import pandas as pd
import smtplib
from email.message import EmailMessage
import argparse

def send_email(to_email, first_name, second_name, medicine, batch_no, expiry, po_header):
    subject = f"Medicine Expiry Notification: {medicine}"
    body = f"""Dear {first_name} {second_name},

This email is for informational purposes only.

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
    msg['From'] = "your_email@example.com"  # Replace with your email
    msg['To'] = to_email
    msg.set_content(body)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Replace with your SMTP server
            server.starttls()
            server.login("your_email@example.com", "your_password")  # Replace with your credentials
            server.send_message(msg)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def process_excel(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        for _, row in df.iterrows():
            first_name, second_name = row['Contact Person Name'].split(" ", 1)
            send_email(
                to_email=row['Email'],
                first_name=first_name,
                second_name=second_name,
                medicine=row['Description'],
                batch_no=row['Batch No'],
                expiry=row['Expiry'],
                po_header="Header Placeholder"  # Replace with actual header if available
            )
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send emails based on Excel data.")
    parser.add_argument("file_path", help="Path to the Excel file.")
    args = parser.parse_args()

    process_excel(args.file_path)