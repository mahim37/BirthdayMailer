import logging
import os
import datetime
from typing import List, Optional, Tuple

from utils.google_sheets import GoogleSheet, GoogleSheetError
from utils.email_sender import EmailSender, EmailError
import config

logger = logging.getLogger(__name__)


def _parse_birthday(date_str: str) -> Optional[datetime.date]:
    """
    Tries to parse a date string using predefined formats.

    Args:
        date_str (str): The date string to parse.

    Returns:
        Optional[datetime.date]: The parsed date object, or None if parsing fails.
    """
    if not date_str:
        return None

    for fmt in config.SUPPORTED_DATE_FORMATS:
        try:
            # Use datetime.strptime and then .date()
            return datetime.datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue  # Try the next format
    logger.warning(
        f"Could not parse date string '{date_str}' with any supported format."
    )
    return None


def _get_sheet_data(
    sheet: GoogleSheet,
) -> Optional[Tuple[List[str], List[str], List[str]]]:
    """Fetches and validates required columns from the Google Sheet."""
    try:
        name_col_idx = sheet.find_column_index(
            config.NAME_COLUMN_HEADER, config.HEADER_ROW
        )
        bday_col_idx = sheet.find_column_index(
            config.BIRTHDAY_COLUMN_HEADER, config.HEADER_ROW
        )
        email_col_idx = sheet.find_column_index(
            config.EMAIL_COLUMN_HEADER, config.HEADER_ROW
        )

        if not all([name_col_idx, bday_col_idx, email_col_idx]):
            missing = [
                h
                for h, i in [
                    (config.NAME_COLUMN_HEADER, name_col_idx),
                    (config.BIRTHDAY_COLUMN_HEADER, bday_col_idx),
                    (config.EMAIL_COLUMN_HEADER, email_col_idx),
                ]
                if i is None
            ]
            logger.error(
                f"Missing required columns in row {config.HEADER_ROW}: {', '.join(missing)}"
            )
            return None

        data_start_row = config.HEADER_ROW + 1
        names = sheet.get_column_values(name_col_idx, start_row=data_start_row)
        birthdays = sheet.get_column_values(bday_col_idx, start_row=data_start_row)
        emails = sheet.get_column_values(email_col_idx, start_row=data_start_row)

        # Basic validation: Check if columns have reasonably consistent lengths
        # Allow for some trailing empty rows, but warn if lengths are drastically different.
        lengths = [len(names), len(birthdays), len(emails)]
        if len(set(lengths)) > 1:  # More than one unique length found
            logger.warning(
                f"Columns have potentially misaligned data. Lengths - "
                f"{config.NAME_COLUMN_HEADER}: {lengths[0]}, "
                f"{config.BIRTHDAY_COLUMN_HEADER}: {lengths[1]}, "
                f"{config.EMAIL_COLUMN_HEADER}: {lengths[2]}. Processing based on shortest length."
            )
            # return None

        return names, birthdays, emails

    except GoogleSheetError as e:
        logger.error(f"Failed to fetch data from Google Sheet: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error fetching sheet data: {e}", exc_info=True)
        return None


def process_birthdays() -> None:
    """
    Main logic to process birthdays and send emails.
    Connects to Google Sheets, fetches data, checks birthdays, and sends emails.
    """
    logger.info("Starting birthday processing...")

    try:
        # Initialize Google Sheet connection using config
        googlesheet = GoogleSheet(
            credentials=config.GOOGLE_CREDENTIALS,
            file_name=config.SHEET_FILE_NAME,
            sheet_name=config.SHEET_NAME,
        )
    except GoogleSheetError as e:
        logger.error(f"Failed to initialize Google Sheet connection: {e}. Aborting.")
        return
    except Exception as e:
        logger.error(f"Unexpected error initializing Google Sheet: {e}", exc_info=True)
        return

    # Fetch data
    sheet_data = _get_sheet_data(googlesheet)
    if sheet_data is None:
        logger.error("Aborting due to issues fetching or validating sheet data.")
        return

    names, birthdays, emails = sheet_data
    min_len = min(
        len(names), len(birthdays), len(emails)
    )  # Process up to the shortest column length

    # Initialize Email Sender using config
    try:
        email_sender = EmailSender(
            smtp_server=config.SMTP_SERVER,
            smtp_port=config.SMTP_PORT,
            smtp_username=config.SMTP_USERNAME,
            smtp_password=config.SMTP_PASSWORD,
            sender_email=config.SENDER_EMAIL,
        )
    except Exception as e:  # Catch potential init errors, though unlikely here
        logger.error(f"Failed to initialize EmailSender: {e}. Aborting.")
        return

    # Process Birthdays
    today = datetime.date.today()
    logger.info(f"Processing birthdays for {today.strftime('%Y-%m-%d')}")
    emails_sent_today = 0
    emails_failed_today = 0

    # Collect all valid emails from the sheet for potential CC list
    # Filter out the birthday person's email later.
    all_valid_emails_in_sheet = [
        email.strip() for email in emails[:min_len] if email and "@" in email.strip()
    ]

    for i in range(min_len):
        row_num = i + config.HEADER_ROW + 1
        person_name = names[i].strip()
        bday_str = birthdays[i].strip()
        receiver_email = emails[i].strip()

        if not person_name:
            logger.warning(f"Skipping Row {row_num}: Name is missing.")
            continue
        if not receiver_email or "@" not in receiver_email:
            logger.warning(
                f"Skipping Row {row_num} ({person_name}): Invalid or missing email '{receiver_email}'."
            )
            continue
        if not bday_str:
            logger.warning(
                f"Skipping Row {row_num} ({person_name}): Birthday is missing."
            )
            continue

        birthday_date = _parse_birthday(bday_str)
        if birthday_date is None:
            logger.warning(
                f"Skipping Row {row_num} ({person_name}): Could not parse birthday '{bday_str}'."
            )
            continue

        if today.month == birthday_date.month and today.day == birthday_date.day:
            logger.info(
                f"MATCH FOUND: Row {row_num} - {person_name}'s birthday is today!"
            )

            try:
                first_name = person_name.split()[0]

                # Prepare CC list: all *other* valid emails from the sheet
                cc_list = [
                    em
                    for em in all_valid_emails_in_sheet
                    if em.lower() != receiver_email.lower()
                ]

                logger.info(
                    f"Preparing email for {person_name} to {receiver_email} (CC: {len(cc_list)} others)..."
                )

                email_sender.send_birthday_email(
                    first_name=first_name,
                    receiver_email=receiver_email,
                    cc_emails=cc_list,
                    template_path=config.TEMPLATE_PATH,
                    image_path=config.IMAGE_PATH
                    if os.path.exists(config.IMAGE_PATH)
                    else None,
                    company_name=config.COMPANY_NAME,
                )
                emails_sent_today += 1

            except EmailError as e:
                # Catch errors specifically from email sending for this person
                logger.error(
                    f"FAILED to send email for {person_name} (Row {row_num}) to {receiver_email}: {e}"
                )
                emails_failed_today += 1
            except Exception as e:
                # Catch any other unexpected errors during this person's processing
                logger.error(
                    f"Unexpected error processing birthday for {person_name} (Row {row_num}): {e}",
                    exc_info=True,
                )
                emails_failed_today += 1
            # Continue processing the rest of the list regardless of individual failures

    logger.info("Birthday processing finished.")
    logger.info(
        f"Summary: Emails Sent = {emails_sent_today}, Emails Failed = {emails_failed_today}"
    )
