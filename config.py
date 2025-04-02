import os
import json
import logging
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

SUPPORTED_DATE_FORMATS = ["%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d", "%Y-%m-%d", "%m/%d"]


class ConfigError(Exception):
    """Custom exception for configuration errors."""

    pass


def _get_required_env(var_name: str) -> str:
    """Gets a required environment variable or raises ConfigError."""
    value = os.environ.get(var_name)
    if not value:
        logger.critical(f"Missing required environment variable: {var_name}")
        raise ConfigError(f"Missing required environment variable: {var_name}")
    return value


def _parse_google_creds(creds_json_str: str) -> dict:
    """Parses the Google Credentials JSON string."""
    try:
        return json.loads(creds_json_str)
    except json.JSONDecodeError as e:
        logger.critical("Failed to parse GOOGLE_CREDENTIALS JSON.")
        raise ConfigError("Invalid JSON format in GOOGLE_CREDENTIALS.") from e
    except Exception as e:
        logger.critical(f"Unexpected error parsing GOOGLE_CREDENTIALS: {e}")
        raise ConfigError("Failed to process GOOGLE_CREDENTIALS.") from e


# --- Core Configuration ---
try:
    # Google Sheets Config
    SHEET_FILE_NAME: str = os.environ.get("SHEET_FILE_NAME", "Birthday List")
    SHEET_NAME: str = os.environ.get(
        "SHEET_NAME", "test"
    )  # Default to 'Sheet1' which is common
    HEADER_ROW: int = int(os.environ.get("HEADER_ROW", "1"))
    NAME_COLUMN_HEADER: str = os.environ.get("NAME_COLUMN_HEADER", "Name")
    BIRTHDAY_COLUMN_HEADER: str = os.environ.get("BIRTHDAY_COLUMN_HEADER", "Birthday")
    EMAIL_COLUMN_HEADER: str = os.environ.get("EMAIL_COLUMN_HEADER", "Emails")

    # Email Config
    SENDER_EMAIL: str = _get_required_env("SENDER_EMAIL")
    # Default SMTP username to sender email if not provided
    SMTP_USERNAME: str = os.environ.get("SMTP_USERNAME", SENDER_EMAIL)
    SMTP_PASSWORD: str = _get_required_env("SMTP_PASSWORD")
    SMTP_SERVER: str = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT: int = int(os.environ.get("SMTP_PORT", "587"))  # TLS port

    # Content Config
    IMAGE_PATH: str = os.environ.get("IMAGE_PATH", "./birthday_image.jpg")
    TEMPLATE_PATH: str = os.environ.get(
        "TEMPLATE_PATH", "templates/birthday_email.html"
    )
    COMPANY_NAME: str = os.environ.get("COMPANY_NAME", "Fischer Jordan")

    # Google Credentials
    _google_creds_json = _get_required_env("GOOGLE_CREDENTIALS")
    GOOGLE_CREDENTIALS: dict = _parse_google_creds(_google_creds_json)

    # Validate Image Path Existence
    if not os.path.exists(IMAGE_PATH):
        logger.warning(
            f"IMAGE_PATH '{IMAGE_PATH}' does not exist. Emails will be sent without an image."
        )

    if not os.path.exists(TEMPLATE_PATH):
        logger.error(f"HTML template file not found at {TEMPLATE_PATH}.")
        raise ConfigError(f"TEMPLATE_PATH '{TEMPLATE_PATH}' not found.")


except ConfigError:
    # Logged already in helper functions
    exit(1)
except ValueError as e:
    logger.critical(
        f"Configuration Error: Non-integer value provided for a numeric setting (e.g., HEADER_ROW, SMTP_PORT). Details: {e}"
    )
    exit(1)
except Exception as e:
    logger.critical(
        f"An unexpected error occurred during configuration loading: {e}", exc_info=True
    )
    exit(1)

logger.info("Configuration loaded successfully.")
