import logging
import sys
import time

try:
    import config
    from birthday_processor import process_birthdays
except ImportError as e:
    print(f"ERROR: Failed to import required modules: {e}", file=sys.stderr)
    sys.exit(1)
except config.ConfigError:
    sys.exit(1)  # Exit cleanly if config already determined failure

# Set up basic logging configuration
log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
logging.basicConfig(level=logging.INFO, format=log_format)


# Optionally, add a file handler:
# file_handler = logging.FileHandler("birthday_emailer.log")
# file_handler.setFormatter(logging.Formatter(log_format))
# logging.getLogger().addHandler(file_handler)

logger = logging.getLogger(__name__)


def main_runner(event=None, context=None):
    """
    Main entry point for running the birthday emailer.
    Suitable for direct execution or cloud function trigger.

    Args:
        event (dict, optional): Event data (e.g., from Cloud Function trigger). Defaults to None.
        context (object, optional): Context object (e.g., from Cloud Function trigger). Defaults to None.
    """
    start_time = time.time()
    logger.info("Birthday Emailer Script Started")

    if event:
        logger.info(f"Triggered by event: {type(event)}")
    if context:
        logger.info(f"Function context ID: {getattr(context, 'event_id', 'N/A')}")

    try:
        process_birthdays()

    except Exception as e:
        # Catch any unexpected errors not caught within process_birthdays
        logger.critical(
            f"CRITICAL UNHANDLED ERROR in main execution: {e}", exc_info=True
        )

    finally:
        end_time = time.time()
        duration = end_time - start_time
        logger.info(
            f"Birthday Emailer Script Finished. Duration: {duration:.2f} seconds"
        )


if __name__ == "__main__":
    main_runner()
