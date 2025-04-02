import time
import logging
import functools
from typing import Type, Tuple, Callable, Any, Optional

logger = logging.getLogger(__name__)


def retry_on_exception(
    exceptions_to_retry: Tuple[Type[Exception], ...],
    max_retries: int = 3,
    initial_delay: float = 1.0,
    backoff_factor: float = 2.0,
    check_exception_callback: Optional[Callable[[Exception], bool]] = None,
    log_prefix: str = "Retry",
) -> Callable:
    """
    Decorator factory to retry a function/method if specific exceptions occur.

    Args:
        exceptions_to_retry (Tuple[Type[Exception], ...]): A tuple of Exception types to catch and retry.
        max_retries (int): Maximum number of retry attempts (default: 3). Total calls = 1 + max_retries.
        initial_delay (float): Initial delay in seconds before the first retry (default: 1.0).
        backoff_factor (float): Multiplier for the delay on subsequent retries (default: 2.0).
        check_exception_callback (Optional[Callable[[Exception], bool]]):
            An optional function that takes the caught exception as input.
            It should return True if a retry should be attempted based on the
            exception's properties (e.g., status code), False otherwise.
            If None, retries happen for any caught exception in exceptions_to_retry.
        log_prefix (str): A prefix string for log messages related to retries.

    Returns:
        Callable: The actual decorator.
    """

    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)  # Preserves original function metadata (name, docstring)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            retries = 0
            delay = initial_delay
            last_exception = None

            while retries <= max_retries:
                try:
                    # Attempt to call the original function
                    result = func(*args, **kwargs)
                    # If successful before max_retries, return result
                    if retries > 0:
                        logger.info(
                            f"[{log_prefix}] Function '{func.__name__}' succeeded after {retries} retries."
                        )
                    return result
                except exceptions_to_retry as e:
                    last_exception = e
                    should_retry = True  # Assume retry unless callback says no

                    # If a callback is provided, use it to decide if we should retry this specific exception instance
                    if check_exception_callback:
                        try:
                            should_retry = check_exception_callback(e)
                        except Exception as check_err:
                            logger.error(
                                f"[{log_prefix}] Error in check_exception_callback for '{func.__name__}': {check_err}",
                                exc_info=True,
                            )
                            should_retry = (
                                False  # Don't retry if the check itself fails
                            )

                    if should_retry and retries < max_retries:
                        retries += 1
                        logger.warning(
                            f"[{log_prefix}] Function '{func.__name__}' failed with {type(e).__name__} "
                            f"(Attempt {retries}/{max_retries}). Retrying in {delay:.2f} seconds..."
                        )
                        time.sleep(delay)
                        delay *= backoff_factor  # Increase delay for next time
                    elif should_retry and retries >= max_retries:
                        logger.error(
                            f"[{log_prefix}] Function '{func.__name__}' failed after {max_retries} retries "
                            f"due to {type(e).__name__}. No more retries."
                        )
                        raise last_exception  # Re-raise the last exception after all retries failed
                    else:
                        # Not a retryable instance according to callback, or not a retryable exception type
                        logger.info(
                            f"[{log_prefix}] Function '{func.__name__}' failed with {type(e).__name__}. Not configured for retry or callback returned False."
                        )
                        raise e  # Re-raise the exception immediately
                # Non-retryable exceptions will pass through here and be raised

        return wrapper

    return decorator


# --- Example Callback Function for gspread APIError 503 ---
def is_gspread_503_error(exception: Exception) -> bool:
    """Checks if the exception is gspread.exceptions.APIError with status 503."""
    from gspread.exceptions import (
        APIError,
    )  # Import locally to avoid circular dependency if decorator is widely used

    if isinstance(exception, APIError):
        response = getattr(exception, "response", None)
        status_code = getattr(response, "status_code", None)
        if status_code == 503:
            logger.debug(
                "Detected gspread APIError with status code 503. Retry advised."
            )
            return True
        else:
            logger.debug(
                f"Detected gspread APIError with status code {status_code}. No retry advised by this callback."
            )
            return False
    return False  # Not a gspread.exceptions.APIError
