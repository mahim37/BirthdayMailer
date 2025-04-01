import os
import json
import logging
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from typing import List
import gspread


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GoogleSheet:
    """
    A class to interact with Google Sheets using gspread.
    """
    def __init__(self, file_name: str, sheet_name: str) -> None:
        # Load credentials from the environment variable
        google_creds = os.environ.get("GOOGLE_CREDENTIALS")
        if not google_creds:
            raise ValueError("GOOGLE_CREDENTIALS environment variable is not set.")
        try:
            credentials_dict = json.loads(google_creds)
        except Exception as e:
            raise ValueError("Failed to parse GOOGLE_CREDENTIALS as JSON.") from e

        self.sa = gspread.service_account_from_dict(credentials_dict)
        self.file = self.sa.open(file_name)
        self.sheet = self.file.worksheet(sheet_name)

    def col_string(self, column_int: int) -> str:
        """Converts a column number to its corresponding letters (e.g., 1 -> A, 27 -> AA)."""
        result = ""
        while column_int:
            column_int, remainder = divmod(column_int - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def write_list(self, row: int, column: int, list_entry: List[str], vertical: bool = False, user_entered: bool = True) -> None:
        """Writes a list of values to the sheet, either horizontally or vertically."""
        cell_range = self.letter_range(
            row,
            column,
            len(list_entry) - 1 if not vertical else 0,
            0 if not vertical else len(list_entry) - 1,
        )
        cell_list = self.sheet.range(cell_range)
        for i, val in enumerate(list_entry):
            cell_list[i].value = str(val)
        update_kwargs = {"value_input_option": "USER_ENTERED"} if user_entered else {}
        self.sheet.update_cells(cell_list, **update_kwargs)

    def letter_range(self, row: int, column: int, width: int, height: int) -> str:
        """Constructs a cell range string for Google Sheets."""
        start_cell = f"{self.col_string(column)}{row}"
        end_cell = f"{self.col_string(column + width)}{row + height}"
        return f"{start_cell}:{end_cell}"

    def col_search(self, column_header_name: str, header_row: int = 1) -> int:
        """Searches for a column header and returns its index (1-indexed)."""
        headers = self.sheet.row_values(header_row)
        for index, header in enumerate(headers, start=1):
            if header.strip().lower() == column_header_name.strip().lower():
                return index
        raise ValueError(f"Column header '{column_header_name}' not found.")

    def get_column(self, column_header_name: str, header_row: int = 1) -> List[str]:
        """Retrieves a column of values (excluding the header) from the sheet."""
        col_index = self.col_search(column_header_name, header_row)
        # Ensure we handle potential errors if col_values returns fewer rows than expected
        all_values = self.sheet.col_values(col_index)
        return all_values[header_row:] # Get values starting after the header row


def send_birthday_email(first_name: str, sender_email: str, receiver_email: str,
                        cc_emails: List[str], smtp_username: str, smtp_password: str,
                        image_path: str) -> None:
    """
    Sends a beautifully formatted birthday email with an embedded image.
    """
    subject = f"Happy Birthday, {first_name}!"
    html_body = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="light dark">
    <meta name="supported-color-schemes" content="light dark">
    <title>Happy Birthday, {first_name}!</title>
    <style type="text/css">
        /* Basic Reset & Defaults */
        body, table, td, div, p, a {{ -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }}
        table, td {{ mso-table-lspace: 0pt; mso-table-rspace: 0pt; border-collapse: collapse; }}
        img {{ -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }}
        body {{ margin: 0; padding: 0; width: 100% !important; background-color: #f4f4f4; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; }}

        /* Container */
        .container {{
            max-width: 600px;
            margin: 20px auto;
            background-color: #ffffff;
            border-radius: 10px; /* Softer corners */
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            overflow: hidden; /* Keeps corners rounded */
            border: 1px solid #dddddd; /* Subtle border */
        }}

        /* Header */
        .header {{
            background-color: #5E9CAE; /* A calmer blue/teal */
            padding: 30px 20px;
            text-align: center;
            color: #ffffff;
        }}
        .header h1 {{
            margin: 0;
            font-size: 26px;
            font-weight: 500; /* Slightly lighter weight */
        }}

        /* Content */
        .content {{
            padding: 30px 40px; /* More horizontal padding */
            color: #333333; /* Darker grey for better contrast */
            line-height: 1.6;
            font-size: 16px;
        }}
        .content p {{
            margin: 15px 0;
        }}
        .content .signature {{
            margin-top: 30px;
        }}

        /* Image */
        .image-container {{
            text-align: center;
            margin: 25px 0;
        }}
        .image-container img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px; /* Match container rounding */
            display: block; /* Fix potential spacing issues */
            margin: 0 auto;
        }}

        /* Footer */
        .footer {{
            background-color: #eeeeee; /* Lighter footer background */
            padding: 20px;
            text-align: center;
            font-size: 12px;
            color: #777777; /* Medium grey */
            line-height: 1.4;
        }}
        .footer p {{
            margin: 5px 0;
        }}

        /* Responsive Styles */
        @media screen and (max-width: 640px) {{
            .container {{
                width: 95% !important;
                margin: 10px auto;
                border-radius: 5px;
            }}
            .content {{
                padding: 20px 25px; /* Reduce padding on mobile */
            }}
            .header {{
                padding: 20px 15px;
            }}
             .header h1 {{
                font-size: 22px;
            }}
             .content {{
                font-size: 15px;
            }}
        }}

        /* Basic Dark Mode Styles (Optional - Client support varies) */
        :root {{
          color-scheme: light dark;
          supported-color-schemes: light dark;
        }}
        @media (prefers-color-scheme: dark) {{
          body, .container {{ background-color: #1e1e1e !important; }}
          .header {{ background-color: #3a6d7e !important; }} /* Darker shade for header */
          .content, .content p {{ color: #e0e0e0 !important; }}
          .footer {{ background-color: #2c2c2c !important; }}
          .footer p {{ color: #aaaaaa !important; }}
          .container {{ border-color: #444444 !important; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Happy Birthday, {first_name}!</h1>
        </div>

        <div class="content">
            <p>Dear {first_name},</p>
            <p>Wishing you a very happy birthday! May this special day bring you joy, happiness, success, and fulfillment in every aspect of your life.</p>

            <div class="image-container">
                <img src="cid:image1" alt="Happy Birthday from Team Fischer Jordan" width="520" style="max-width:100%; height:auto; border-radius: 8px;">
            </div>

            <p>We hope you have a fantastic day celebrating!</p>

            <p class="signature">Best regards,<br><strong>Team Fischer Jordan</strong></p>
        </div>

        <div class="footer">
            <p>This birthday wish was sent with care by the team at Fischer Jordan.</p>
            <p>&copy; {datetime.date.today().year} Fischer Jordan. All rights reserved.</p>
        </div>
    </div>
    </body>
</html>
"""

    # --- Plain Text Body (Keep it simple and clear) ---
    text_body = f"""
Dear {first_name},

Wishing you a very happy birthday! May this special day bring you joy, happiness, success, and fulfillment in every aspect of your life.

We hope you have a fantastic day celebrating!

Best regards,
Team Fischer Jordan

---
© {datetime.date.today().year} Fischer Jordan. All rights reserved.
"""

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = f"Team Fischer Jordan <{sender_email}>"
    message["To"] = receiver_email
    if cc_emails:
        message["Cc"] = ", ".join(cc_emails)

    # Attach plain text first, then HTML
    message.attach(MIMEText(text_body, "plain", "utf-8"))
    message.attach(MIMEText(html_body, "html", "utf-8"))

    # --- Attach Image ---
    try:
        if os.path.exists(image_path):
            with open(image_path, "rb") as image_file:
                image_data = image_file.read()
            image_part = MIMEImage(image_data, name=os.path.basename(image_path))
            # The Content-ID should match the src in the <img> tag (cid:image1)
            image_part.add_header("Content-ID", "<image1>")
            # Add Content-Disposition as inline for better compatibility
            image_part.add_header('Content-Disposition', 'inline', filename=os.path.basename(image_path))
            message.attach(image_part)
        else:
             logger.warning(f"Image file not found at {image_path}. Sending email without image.")

    except Exception as e:
        logger.error(f"Error attaching image from {image_path}: {e}")
        pass # Continue sending email without the image if attach fails

    # --- Send Email ---
    smtp_server = "smtp.gmail.com"
    smtp_port = 587 # Standard TLS port

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo() # Greet the server
            server.starttls() # Start TLS encryption
            server.ehlo() # Re-greet after TLS
            server.login(smtp_username, smtp_password)
            recipients = [receiver_email] + cc_emails
            server.sendmail(sender_email, recipients, message.as_string())
            logger.info(f"Successfully sent birthday email to {receiver_email} (CC: {', '.join(cc_emails) if cc_emails else 'None'})")
    except smtplib.SMTPAuthenticationError as e:
         logger.error(f"SMTP Authentication Error: Check username/password or app password settings. Details: {e}")
         raise # Reraise critical auth errors
    except smtplib.SMTPServerDisconnected:
        logger.error("SMTP Server disconnected unexpectedly. Check network or server status.")
        raise
    except Exception as e:
        logger.error(f"Error sending email to {receiver_email}: {e}")
        raise # Reraise other critical SMTP errors

def process_birthdays():
    """
    Processes the birthday list from Google Sheets and sends emails on matching birthdays.
    """
    # --- Configuration from Environment Variables ---
    sheet_file = os.environ.get("SHEET_FILE", "Birthday List")
    sheet_name = os.environ.get("SHEET_NAME", "test")
    image_path = os.environ.get("IMAGE_PATH", "./birthday_image.jpg")
    sender_email = os.environ.get("SENDER_EMAIL")
    smtp_username = os.environ.get("SMTP_USERNAME", sender_email)
    smtp_password = os.environ.get("SMTP_PASSWORD")

    # --- Validate Essential Configuration ---
    if not sender_email:
        logger.critical("SENDER_EMAIL environment variable not set. Exiting.")
        return
    if not smtp_password:
        logger.critical("SMTP_PASSWORD environment variable not set. Exiting.")
        return
    if not os.environ.get("GOOGLE_CREDENTIALS"):
         logger.critical("GOOGLE_CREDENTIALS environment variable not set. Exiting.")
         return
    # Check if image exists early
    if not os.path.exists(image_path):
        logger.warning(f"IMAGE_PATH '{image_path}' does not exist. Emails will be sent without an image.")


    # --- Initialize Google Sheet ---
    try:
        googlesheet = GoogleSheet(sheet_file, sheet_name)
        logger.info(f"Connected to Google Sheet: '{sheet_file}' -> Sheet: '{sheet_name}'")
    except Exception as e:
        logger.error(f"Failed to initialize Google Sheet connection: {e}")
        return # Cannot proceed without sheet connection

    # --- Fetch Data from Sheet ---
    try:
        # Specify the header row explicitly if it's not the first row
        header_row_number = 1
        names = googlesheet.get_column("Name", header_row=header_row_number)
        birthdays = googlesheet.get_column("Birthday", header_row=header_row_number)
        emails = googlesheet.get_column("Emails", header_row=header_row_number)

        # Basic validation: Check if columns have the same length
        if not (len(names) == len(birthdays) == len(emails)):
            logger.warning("Columns 'Name', 'Birthday', and 'Emails' have different lengths. Data might be misaligned.")
            # You might want to stop here or proceed cautiously
            return

    except ValueError as e:
         logger.error(f"Sheet Error: {e}. Make sure 'Name', 'Birthday', 'Emails' columns exist in row {header_row_number}.")
         return
    except Exception as e:
        logger.error(f"Error fetching data from the sheet: {e}")
        return

    # --- Process Birthdays ---
    today = datetime.date.today()
    logger.info(f"Processing birthdays for {today.strftime('%Y-%m-%d')}")
    emails_sent_today = 0
    
    # Collect all valid emails for CC
    all_valid_emails = [email.strip() for email in emails if email and '@' in email.strip()]

    for i, bday_str in enumerate(birthdays):
        if i >= len(names) or i >= len(emails):
            logger.warning(f"Skipping row {i + header_row_number + 1} due to data length mismatch.")
            continue

        person_name = names[i].strip()
        receiver_email = emails[i].strip()

        # Basic validation for email and name
        if not person_name:
            logger.warning(f"Skipping row {i + header_row_number + 1}: Name is missing.")
            continue
        if not receiver_email or '@' not in receiver_email:
            logger.warning(f"Skipping {person_name} (Row {i + header_row_number + 1}): Invalid or missing email '{receiver_email}'.")
            continue

        # --- Parse Birthday ---
        try:
            # Be flexible with date formats if possible, but standardization is best
            birthday_date = None
            possible_formats = ["%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d", "%Y-%m-%d", "%m/%d"]
            for fmt in possible_formats:
                try:
                    birthday_date = datetime.datetime.strptime(bday_str.strip(), fmt).date()
                    break # Stop trying formats once one works
                except ValueError:
                    continue # Try next format

            if birthday_date is None:
                 raise ValueError("Date format not recognized.")


        except Exception as e:
            logger.warning(f"Skipping {person_name} (Row {i + header_row_number + 1}): Invalid date format '{bday_str}'. Error: {e}")
            continue # Skip to next person

        # --- Check if today is the birthday ---
        if today.month == birthday_date.month and today.day == birthday_date.day:
            try:
                first_name = person_name.split()[0] # Get the first name

                # Prepare CC list: all *other* valid emails from the sheet
                cc_emails_for_this_person = [
                    email for email in all_valid_emails if email != receiver_email
                ]

                logger.info(f"MATCH FOUND: {person_name}'s birthday is today! Preparing email to {receiver_email}...")

                send_birthday_email(
                    first_name=first_name,
                    sender_email=sender_email,
                    receiver_email=receiver_email,
                    cc_emails=cc_emails_for_this_person,
                    smtp_username=smtp_username,
                    smtp_password=smtp_password,
                    image_path=image_path
                )
                emails_sent_today += 1

            except Exception as e:
                # Catch errors during the sending process for this specific person
                logger.error(f"FAILED to send email for {person_name} to {receiver_email}: {e}")
                # Continue processing the rest of the list

    logger.info(f"Birthday processing complete. Sent {emails_sent_today} emails today.")


def main(event=None, context=None):
    """
    Main entry point for running the birthday emailer.
    Suitable for direct execution or cloud function trigger.
    """
    logger.info("Birthday Emailer script started.")
    try:
        process_birthdays()
    except Exception as e:
        # Catch any unexpected errors not caught within process_birthdays
        logger.critical(f"CRITICAL UNEXPECTED ERROR in main execution: {e}", exc_info=True) # Log traceback
    finally:
        logger.info("Birthday Emailer script finished.")

if __name__ == "__main__":
    main()
