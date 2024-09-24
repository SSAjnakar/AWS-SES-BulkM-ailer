# AWS-SES-BulkMailer


## Overview
The **AWS-SES-BulkMailer** is a Python application that allows users to send personalized emails to multiple recipients using AWS Simple Email Service (SES). It features a user-friendly graphical interface built with Tkinter, enabling users to easily manage contact and content files, configure email settings, and monitor the email sending process.

## Features
- **Tabbed Interface**: Organizes functionality into separate tabs for file selection, email configuration, and error logs.
- **File Management**: Supports CSV files for contacts and email content, allowing users to browse and select files easily.
- **Email Configuration**: Users can specify the sender's email and delay between sending emails.
- **Email Sending Process**: Emails are sent in the background to prevent freezing of the GUI.
- **Progress Tracking**: A progress bar and status messages keep users informed about the email sending process.
- **Log Generation**: Logs are created for each email sent, detailing the recipient and the status of the email.

## Prerequisites
- Python 3.x
- `boto3` library for AWS SES
- `pandas` library for handling CSV files
- Tkinter (comes pre-installed with Python)

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/ImprovedEmailClient.git
   cd ImprovedEmailClient

Contact File (campagin_name1_contact.csv): first name, Email
                                            Xyz, xyz@abc.com
Contact File (campagin_name1_content.csv): subject, Body
                                            Introduction Email, Sample text sample text sample text sample text sample textsample text sample textsample text sample textsample text sample textsample text sample textsample text sample textsample text sample textsample text sample textsample text sample textsample text sample text

                                            

