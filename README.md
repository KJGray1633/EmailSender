# EmailSender

This project provides a simple Python class to send emails using Microsoft Outlook via the `win32com` library.

## Features
- Send emails from your Outlook account
- Specify subject, body, recipient, and CC addresses
- Choose to display the email for manual sending or send automatically

## Requirements
- Python 3.x
- `pywin32` package (`pip install pywin32`)
- Microsoft Outlook installed and configured

## Usage

```python
from EmailSender import EmailSender

sender = EmailSender()
subject = "Test Email"
body = "This is a test email sent from Python."
recipient = "recipient@example.com"
cc_emails = ["cc1@example.com", "cc2@example.com"]

# Display the email (manual send)
sender.send_email(subject, body, recipient, cc_emails, display_or_send='d')

# Send the email automatically
# sender.send_email(subject, body, recipient, cc_emails, display_or_send='s')
```

## Parameters
- `subject`: Subject of the email
- `body`: Body text of the email
- `recipient`: Recipient's email address
- `cc_emails`: List of CC email addresses (optional)
- `display_or_send`: `'d'` to display the email for manual sending, `'s'` to send automatically