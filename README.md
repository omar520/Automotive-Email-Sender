# Automotive-Email-Sender
This Python project automates the process of sending personalized emails with links or attachments. It reads recipient details from an Excel file, generates custom email messages, and delivers them via Gmail's SMTP service. This tool is ideal for sending certificates, announcements, or appreciation emails efficiently.

# Email Certificate Sender

This project is a Python-based tool that automates the process of sending personalized emails with certificates to recipients. It reads recipient information from an Excel file, customizes the email content, and sends the emails using Gmail's SMTP server.

## Features
- Reads recipient details (name, email, certificate link, etc.) from an Excel file.
- Sends personalized emails with a certificate link or attaches the certificate as a PDF.
- Can be customized for any type of email campaign (e.g., certificate delivery, announcements).
- Utilizes Gmail's SMTP server for secure email delivery.

## Prerequisites
- Python 3.x
- Required Python libraries:
  - `pandas` for reading the Excel file.
  - `smtplib`, `ssl` for sending emails.
  - `email` for composing the emails.
  - `Pillow` (optional, if you're generating certificates as images).

You can install these dependencies using pip:

```bash
pip install pandas Pillow
