# MedNotify

MedNotify is a simple command-line tool that reads data from an Excel file and sends automated email notifications about medicine expiry details. It's designed to help healthcare services streamline their communication and avoid manual tracking of critical information.

## Features

- Reads data from an Excel file.
- Sends personalized email notifications to recipients based on the data.
- Easy-to-use command-line interface.

## Prerequisites

- Python 3.7 or higher.
- Create a virtual environment and install the required Python libraries:

```
1. python3 -m venv .venv
2. source .venv/bin/activate
3. pip install -r requirements.txt
```

## For .env file

`create a .env file in the root directory by referring .env.example file and update the values accordingly.`:

### For EMAIL_PASSWORD

Enable 2-Step Verification (2SV) and Generate App Password:

Go to your Google account security settings: https://myaccount.google.com/intro/security.
Enable 2-step verification if it's not already enabled. This adds an extra layer of security.
Once enabled, navigate to "App passwords" and generate a new password specifically for your script.
Update Script with App Password:

Replace the PASSWORD variable in your .env file with the generated App Password, not your regular Gmail password.

## To run the application

- python main.py path-to-file.xlsx