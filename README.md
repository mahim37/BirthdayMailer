# Automated Birthday Emailer

## Description

This Python application automatically checks a Google Sheet for contacts whose birthday matches the current date and sends them a personalized, formatted birthday email using SMTP. It's designed to be run daily (e.g., via cloud scheduler).

## Features

* **Google Sheets Integration:** Reads contact names, birthdays, and email addresses from a specified Google Sheet.
* **Secure Authentication:** Uses Google Service Account credentials for secure API access.
* **SMTP Email Sending:** Sends emails via any standard SMTP server (e.g., Gmail, Outlook). 
* **Customizable HTML Emails:** Uses an HTML template (`templates/birthday_email.html`) for visually appealing emails.
* **Image Embedding:** Embeds a specified image within the email body.
* **CC Functionality:** Automatically CCs all other valid email addresses found in the sheet on each birthday email (excluding the recipient).
* **Configuration Driven:** All settings (credentials, sheet names, paths, server details) are managed via environment variables for flexibility and security. Supports `.env` file for local development.
* **Modular Design:** Code is split into logical modules (`config`, `google_sheets`, `email_sender`, `birthday_processor`) for better readability, maintainability, and testability.
* **Robust Error Handling:** Includes specific error handling for configuration issues, sheet access problems, date parsing errors, and SMTP failures.
* **Logging:** Provides informative console logging about the script's execution progress and any errors encountered.

## Prerequisites

* **Python:** Version 3.7 or higher.
* **Google Cloud Platform (GCP) Account:** Required to enable APIs and create service accounts.
* **Google Service Account:**
    * A Service Account created within your GCP project.
    * A JSON key file generated for this Service Account.
    * The **Google Sheets API** enabled in your GCP project.
* **Google Sheet:**
    * A Google Sheet containing contact information.
    * The sheet must be **shared** with the Service Account's email address (found in the JSON key file, looks like `your-service-account-name@your-project-id.iam.gserviceaccount.com`). Grant "Editor" permissions if you intend to use write features (though this script currently only reads).
    * The sheet needs specific columns (see **Google Sheet Format** below).
* **Email Account (SMTP):**
    * An email account that allows sending via SMTP.
    * If using **Gmail**, you **must** enable 2-Factor Authentication and generate an **App Password** to use in the configuration instead of your regular Gmail password. See: [Google App Passwords](https://myaccount.google.com/apppasswords)

## Setup & Installation

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/mahim37/BirthdayMailer/
    cd BirthdayMailer/
    ```

2.  **Create and Activate Virtual Environment:** (Recommended)
    ```bash
    python -m venv venv
    # On Linux/macOS:
    source venv/bin/activate
    # On Windows:
    venv\Scripts\activate
    ```

3.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

Configuration is handled via environment variables. For local development, you can use a `.env` file.

1.  **Create `.env` File:**
    * Copy the example file: `cp .env.example .env` (or `copy .env.example .env` on Windows).
    * **IMPORTANT:** Add `.env` to your `.gitignore` file to prevent accidentally committing secrets!
        ```gitignore
        # .gitignore
        .env
        venv/
        __pycache__/
        *.pyc
        *.log
        ```

2.  **Edit `.env` File:**
    * Open the `.env` file and fill in the values specific to your setup. Refer to the comments within `.env.example` for explanations of each variable.
    * **`GOOGLE_CREDENTIALS`**: Paste the *entire content* of your downloaded Service Account JSON key file as a single line string within single quotes. Be careful with special characters if pasting directly in some shells; setting it as an environment variable directly might be easier in deployment environments.
    * **`SMTP_PASSWORD`**: Use your email account's password or, **for Gmail, use the generated App Password.**
    * **`SHEET_FILE_NAME` / `SHEET_NAME`**: Ensure these exactly match your Google Sheet file and worksheet names.
    * **Column Headers**: Match the header text in your specified `HEADER_ROW`.
    * **`IMAGE_PATH` / `TEMPLATE_PATH`**: Ensure these paths are correct relative to where you run the script. Place your image (e.g., `birthday_image.png`) inside the `templates` directory or update the path accordingly.

## Usage

1.  **Ensure Configuration is Complete:** Double-check your `.env` file (or environment variables if deploying).
2.  **Run the Script:**
    ```bash
    python main.py
    ```

## Project Structure


├── .env.example           # Example environment variables template
├── .gitignore             # Specifies intentionally untracked files
├── config.py              # Loads and validates configuration from environment/.env
├── email_sender.py        # Class for composing and sending emails via SMTP
├── google_sheets.py       # Class for interacting with the Google Sheets API
├── birthday_processor.py  # Core logic: reads sheet, checks dates, triggers emails
├── main.py                # Main entry point for the application
├── README.md              # This documentation file
├── requirements.txt       # Project dependencies
└── templates/
  ├── birthday_email.html  # HTML template for the birthday email


## Google Sheet Format

The script expects the Google Sheet specified in the configuration (`SHEET_FILE_NAME` & `SHEET_NAME`) to have the following structure:

* **Header Row:** The row number containing headers is defined by `HEADER_ROW` in the config (defaults to 1).
* **Required Columns:** The sheet *must* contain columns with headers exactly matching the values set for:
    * `NAME_COLUMN_HEADER` (Default: "Name")
    * `BIRTHDAY_COLUMN_HEADER` (Default: "Birthday")
    * `EMAIL_COLUMN_HEADER` (Default: "Emails")
* **Birthday Format:** The dates in the Birthday column should match one of the formats defined in `config.SUPPORTED_DATE_FORMATS`. Default supported formats include:
    * `MM/DD/YYYY`
    * `MM-DD-YYYY`
    * `YYYY/MM/DD`
    * `YYYY-MM-DD`
    * `MM/DD` (Year will be ignored for matching)
    * Ensure dates are formatted consistently for best results.

## Error Handling & Logging

* The script logs information about its progress and any errors to the console using standard Python logging (`INFO` level by default).
* Configuration errors (missing variables, invalid JSON) are critical and will prevent the script from running.
* Errors related to accessing Google Sheets (authentication, sheet/worksheet not found, API errors) are logged and will typically halt execution.
* Invalid data within a specific row (e.g., unparseable date, missing email) will cause that row to be skipped, and a warning will be logged. The script will continue processing other rows.
* Errors during SMTP email sending (authentication failure, connection issues, recipient refused) for a specific person are logged, and the script will attempt to send emails for other matching birthdays.

## Security

* **Never commit your `.env` file or hardcode credentials (Service Account JSON, SMTP password) directly into the source code.** Use environment variables or a secure secrets management system for deployment.
* Ensure your Google Service Account key file (`.json`) is stored securely and its access is restricted.
* When using Gmail, always prefer **App Passwords** over your main account password for better security.

---





