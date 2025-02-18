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
        return self.sheet.col_values(col_index)[1:]  # Exclude header row

def send_birthday_email(first_name: str, sender_email: str, receiver_email: str,
                        cc_emails: List[str], smtp_username: str, smtp_password: str,
                        image_path: str) -> None:
    """
    Sends a birthday email with an embedded image.
    """
    subject = f"Happy Birthday, {first_name}!"
    body = f"""
    <html>
        <body>
            <p>Dear {first_name},</p>
            <p>Wishing you a very happy birthday! May this special day bring you happiness, success, and fulfillment in every aspect of your life!</p>
            <p>Best regards,<br>Team Fischer Jordan</p>
            <p><img src="cid:image1"></p>
        </body>
    </html>
    """
    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email
    if cc_emails:
        message["Cc"] = ", ".join(cc_emails)
    message.attach(MIMEText(body, "html"))

    try:
        with open(image_path, "rb") as image_file:
            image_data = image_file.read()
        image_part = MIMEImage(image_data, name=os.path.basename(image_path))
        image_part.add_header("Content-ID", "<image1>")
        message.attach(image_part)
    except Exception as e:
        logger.error(f"Error attaching image from {image_path}: {e}")
        raise

    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            recipients = [receiver_email] + cc_emails
            server.sendmail(sender_email, recipients, message.as_string())
            logger.info(f"Email sent to {receiver_email}")
    except Exception as e:
        logger.error(f"Error sending email to {receiver_email}: {e}")
        raise

def process_birthdays():
    """
    Processes the birthday list from Google Sheets and sends emails on matching birthdays.
    """
    sheet_file = os.environ.get("SHEET_FILE", "Birthday List")
    sheet_name = os.environ.get("SHEET_NAME", "test")
    image_path = os.environ.get("IMAGE_PATH", "./image.jpg")
    sender_email = os.environ.get("SENDER_EMAIL", "mahim.gupta@fischerjordan.com")
    smtp_username = os.environ.get("SMTP_USERNAME", sender_email)
    smtp_password = os.environ.get("SMTP_PASSWORD", "")

    try:
        googlesheet = GoogleSheet(sheet_file, sheet_name)
    except Exception as e:
        logger.error(f"Failed to initialize Google Sheet: {e}")
        return

    try:
        names = googlesheet.get_column("Name")
        birthdays = googlesheet.get_column("Birthday")
        emails = googlesheet.get_column("Emails")
    except Exception as e:
        logger.error(f"Error fetching data from the sheet: {e}")
        return

    today = datetime.date.today()
    logger.info(f"Processing birthdays for {today}")

    for i, bday_str in enumerate(birthdays):
        try:
            birthday = datetime.datetime.strptime(bday_str.strip(), "%m/%d/%Y").date()
        except Exception as e:
            logger.warning(f"Skipping row {i+2} due to invalid date format '{bday_str}': {e}")
            continue

        if today.month == birthday.month and today.day == birthday.day:
            try:
                person_name = names[i].strip()
                first_name = person_name.split()[0]
                receiver_email = emails[i].strip()
                cc_emails = [email.strip() for j, email in enumerate(emails) if j != i and email.strip()]
                logger.info(f"Sending birthday email to {person_name} at {receiver_email}")
                send_birthday_email(first_name, sender_email, receiver_email,
                                    cc_emails, smtp_username, smtp_password, image_path)
            except Exception as e:
                logger.error(f"Failed to send email for {person_name}: {e}")

def main(event=None, context=None):
    """
    Main entry point for running the birthday emailer.
    When deployed as a cloud function or scheduled task, 'event' and 'context' can be provided.
    """
    try:
        process_birthdays()
    except Exception as e:
        logger.error(f"Unexpected error: {e}")

if __name__ == "__main__":
    main()
