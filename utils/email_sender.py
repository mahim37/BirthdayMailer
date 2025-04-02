import os
import smtplib
import logging
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from typing import List, Optional

logger = logging.getLogger(__name__)


class EmailError(Exception):
    """Custom exception for email sending errors."""

    pass


class EmailSender:
    """
    Handles the composition and sending of emails via SMTP.
    """

    def __init__(
        self,
        smtp_server: str,
        smtp_port: int,
        smtp_username: str,
        smtp_password: str,
        sender_email: str,
    ):
        """
        Initializes the EmailSender with SMTP configuration.

        Args:
            smtp_server (str): SMTP server address.
            smtp_port (int): SMTP server port.
            smtp_username (str): Username for SMTP authentication.
            smtp_password (str): Password or App Password for SMTP authentication.
            sender_email (str): The 'From' email address.
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password
        self.sender_email = sender_email
        logger.info(
            f"EmailSender initialized for server {smtp_server}:{smtp_port} with user {smtp_username}"
        )

    def _create_message(
        self,
        receiver_email: str,
        cc_emails: List[str],
        subject: str,
        html_body: str,
        text_body: str,
        image_path: Optional[str],
    ) -> MIMEMultipart:
        """Creates the MIMEMultipart email message."""
        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        message["From"] = self.sender_email
        message["To"] = receiver_email
        if cc_emails:
            message["Cc"] = ", ".join(cc_emails)

        # Attach plain text first, then HTML
        message.attach(MIMEText(text_body, "plain", "utf-8"))
        message.attach(MIMEText(html_body, "html", "utf-8"))

        # Attach Image if path is valid
        if image_path and os.path.exists(image_path):
            try:
                with open(image_path, "rb") as image_file:
                    image_data = image_file.read()
                image_part = MIMEImage(image_data, name=os.path.basename(image_path))
                # Content-ID should match the src in the <img> tag (e.g., cid:birthday_image)
                image_part.add_header("Content-ID", "<birthday_image>")
                image_part.add_header(
                    "Content-Disposition",
                    "inline",
                    filename=os.path.basename(image_path),
                )
                message.attach(image_part)
                logger.info(f"Successfully attached image: {image_path}")
            except FileNotFoundError:
                logger.warning(
                    f"Image file not found at {image_path} during message creation."
                )
            except Exception as e:
                logger.error(
                    f"Error attaching image from {image_path}: {e}", exc_info=True
                )
                # Continue without image if attachment fails
        elif image_path:
            logger.warning(
                f"Image file specified but not found at {image_path}. Sending without image."
            )

        return message

    def _load_html_template(self, template_path: str) -> str:
        """Loads the HTML template from a file."""
        try:
            with open(template_path, "r", encoding="utf-8") as f:
                return f.read()
        except FileNotFoundError:
            logger.error(f"HTML template file not found: {template_path}")
            raise EmailError(f"HTML template not found at {template_path}")
        except Exception as e:
            logger.error(
                f"Error reading HTML template {template_path}: {e}", exc_info=True
            )
            raise EmailError(f"Could not read HTML template: {template_path}") from e

    def send_birthday_email(
        self,
        first_name: str,
        receiver_email: str,
        cc_emails: List[str],
        template_path: str,
        image_path: Optional[str],
        company_name: str,
    ) -> None:
        """
        Composes and sends a personalized birthday email.

        Args:
            first_name (str): Recipient's first name.
            receiver_email (str): Recipient's email address.
            cc_emails (List[str]): List of emails to CC.
            template_path (str): Path to the HTML email template file.
            image_path (Optional[str]): Path to the image file to embed, or None.
            company_name (str): Name of the company for the signature.

        Raises:
            EmailError: If sending the email fails.
        """
        subject = f"Happy Birthday, {first_name}!"
        current_year = datetime.date.today().year

        text_body = f"""
Dear {first_name},

Wishing you a very happy birthday! May this special day bring you joy, happiness, success, and fulfillment.

We hope you have a fantastic day celebrating!

Best regards,
Team {company_name}

---
© {current_year} {company_name}. All rights reserved.
"""

        try:
            html_template = self._load_html_template(template_path)
            html_body = html_template.format(
                first_name=first_name, company_name=company_name, year=current_year
            )
        except KeyError as e:
            logger.error(
                f"Missing placeholder in HTML template {template_path}: {e}. Ensure {{first_name}}, {{company_name}}, {{year}} exist."
            )
            raise EmailError("HTML template formatting error.")
        except Exception as e:
            # Catch errors from _load_html_template or .format()
            logger.error(f"Failed to prepare HTML body: {e}")
            raise EmailError("Failed to prepare email body.") from e

        message = self._create_message(
            receiver_email=receiver_email,
            cc_emails=cc_emails,
            subject=subject,
            html_body=html_body,
            text_body=text_body,
            image_path=image_path,
        )

        recipients = [receiver_email] + cc_emails

        logger.info(
            f"Attempting to send email to {receiver_email} (CC: {', '.join(cc_emails) if cc_emails else 'None'})"
        )
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(self.smtp_username, self.smtp_password)
                server.sendmail(self.sender_email, recipients, message.as_string())
                logger.info(f"Successfully sent birthday email to {receiver_email}")
        except smtplib.SMTPAuthenticationError as e:
            logger.error(
                f"SMTP Authentication Error for user {self.smtp_username}. Check credentials/app password. Details: {e}"
            )
            raise EmailError("SMTP Authentication Failed.") from e
        except smtplib.SMTPServerDisconnected:
            logger.error(
                "SMTP Server disconnected unexpectedly. Check network or server status."
            )
            raise EmailError("SMTP Server Disconnected.")
        except smtplib.SMTPConnectError as e:
            logger.error(
                f"Failed to connect to SMTP server {self.smtp_server}:{self.smtp_port}. Details: {e}"
            )
            raise EmailError("SMTP Connection Failed.") from e
        except smtplib.SMTPRecipientsRefused as e:
            logger.error(
                f"SMTP server refused recipients: {recipients}. Details: {e.recipients}"
            )
            raise EmailError(f"Recipients Refused: {e.recipients}") from e
        except TimeoutError:
            logger.error(
                f"Timeout connecting or sending email via {self.smtp_server}:{self.smtp_port}."
            )
            raise EmailError("SMTP Operation Timed Out.")
        except Exception as e:
            logger.error(
                f"General error sending email to {receiver_email}: {e}", exc_info=True
            )
            raise EmailError("Failed to send email.") from e
