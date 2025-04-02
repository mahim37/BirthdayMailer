import logging
from typing import List, Optional, Dict
import gspread
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound

logger = logging.getLogger(__name__)


class GoogleSheetError(Exception):
    """Custom exception for Google Sheet operations."""

    pass


class GoogleSheet:
    """
    A class to interact with Google Sheets using gspread.
    Handles authentication and common sheet operations.
    """

    def __init__(self, credentials: Dict, file_name: str, sheet_name: str) -> None:
        """
        Initializes the Google Sheet connection.

        Args:
            credentials (Dict): Parsed Google service account credentials.
            file_name (str): The name of the Google Sheet file.
            sheet_name (str): The name of the worksheet within the file.

        Raises:
            GoogleSheetError: If connection or sheet access fails.
        """
        self.file_name = file_name
        self.sheet_name = sheet_name
        self._connect(credentials)
        self._open_sheet()

    def _connect(self, credentials: Dict) -> None:
        """Establishes connection to Google Sheets API."""
        try:
            logger.info("Attempting to connect to Google Sheets API...")
            self.sa = gspread.service_account_from_dict(credentials)
            logger.info("Successfully authenticated with Google Sheets API.")
        except Exception as e:
            logger.error(
                f"Failed to authenticate with Google Sheets API: {e}", exc_info=True
            )
            raise GoogleSheetError("Google Sheets authentication failed.") from e

    def _open_sheet(self) -> None:
        """Opens the specified spreadsheet and worksheet."""
        try:
            logger.info(f"Opening spreadsheet: '{self.file_name}'")
            self.file = self.sa.open(self.file_name)
            logger.info(f"Accessing worksheet: '{self.sheet_name}'")
            self.sheet = self.file.worksheet(self.sheet_name)
            logger.info(
                f"Successfully accessed sheet: '{self.file_name}' -> '{self.sheet_name}'"
            )
        except SpreadsheetNotFound:
            logger.error(
                f"Spreadsheet '{self.file_name}' not found or permission denied."
            )
            raise GoogleSheetError(f"Spreadsheet '{self.file_name}' not found.")
        except WorksheetNotFound:
            logger.error(
                f"Worksheet '{self.sheet_name}' not found in spreadsheet '{self.file_name}'."
            )
            raise GoogleSheetError(f"Worksheet '{self.sheet_name}' not found.")
        except APIError as e:
            logger.error(f"Google API error accessing sheet: {e}", exc_info=True)
            raise GoogleSheetError("Google API error during sheet access.") from e
        except Exception as e:
            logger.error(f"Unexpected error opening sheet: {e}", exc_info=True)
            raise GoogleSheetError("Unexpected error opening sheet.") from e

    @staticmethod
    def _col_string(column_int: int) -> str:
        """Converts a column number (1-indexed) to its corresponding letters (e.g., 1 -> A, 27 -> AA)."""
        if not isinstance(column_int, int) or column_int < 1:
            raise ValueError("Column number must be a positive integer.")
        result = ""
        num = column_int
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _letter_range(self, row: int, column: int, width: int, height: int) -> str:
        """Constructs a cell range string (e.g., "A1:B2")."""
        if not all(isinstance(i, int) and i >= 0 for i in [row, column, width, height]):
            raise ValueError(
                "Row, column, width, and height must be non-negative integers."
            )
        if row == 0 or column == 0:
            raise ValueError(
                "Row and column indices must be 1-based for range calculation."
            )

        start_cell = f"{self._col_string(column)}{row}"
        end_col_num = column + width
        end_row_num = row + height
        end_cell = f"{self._col_string(end_col_num)}{end_row_num}"
        return f"{start_cell}:{end_cell}"

    def write_list(
        self,
        row: int,
        column: int,
        list_entry: List[str],
        vertical: bool = False,
        user_entered: bool = True,
    ) -> None:
        """
        Writes a list of values to the sheet, either horizontally or vertically.

        Args:
            row (int): Starting row number (1-indexed).
            column (int): Starting column number (1-indexed).
            list_entry (List[str]): The list of values to write.
            vertical (bool): If True, write vertically; otherwise, horizontally. Defaults to False.
            user_entered (bool): Use 'USER_ENTERED' value input option. Defaults to True.
        """
        if not list_entry:
            logger.warning("Attempted to write an empty list. Skipping.")
            return
        try:
            height = len(list_entry) - 1 if vertical else 0
            width = 0 if vertical else len(list_entry) - 1
            cell_range_str = self._letter_range(row, column, width, height)

            logger.info(
                f"Writing {len(list_entry)} values to range {cell_range_str}..."
            )
            cell_list = self.sheet.range(cell_range_str)

            for i, val in enumerate(list_entry):
                cell_list[i].value = str(val)  # Ensure value is string

            update_kwargs = (
                {"value_input_option": "USER_ENTERED"} if user_entered else {}
            )
            self.sheet.update_cells(cell_list, **update_kwargs)
            logger.info(f"Successfully updated cells in range {cell_range_str}.")

        except APIError as e:
            logger.error(f"Google API error writing to sheet: {e}", exc_info=True)
            raise GoogleSheetError("Google API error during write operation.") from e
        except Exception as e:
            logger.error(f"Unexpected error writing list to sheet: {e}", exc_info=True)
            raise GoogleSheetError("Unexpected error writing list.") from e

    def find_column_index(
        self, column_header_name: str, header_row: int = 1
    ) -> Optional[int]:
        """
        Searches for a column header in the specified row and returns its index (1-indexed).

        Args:
            column_header_name (str): The exact header name to find (case-insensitive comparison).
            header_row (int): The row number where headers are located (1-indexed).

        Returns:
            Optional[int]: The 1-based index of the column, or None if not found.
        """
        try:
            logger.debug(
                f"Searching for header '{column_header_name}' in row {header_row}..."
            )
            headers = self.sheet.row_values(header_row)
            normalized_search_header = column_header_name.strip().lower()

            for index, header in enumerate(headers, start=1):
                if header.strip().lower() == normalized_search_header:
                    logger.debug(
                        f"Found header '{column_header_name}' at column index {index}."
                    )
                    return index
            logger.warning(
                f"Column header '{column_header_name}' not found in row {header_row}."
            )
            return None
        except APIError as e:
            logger.error(
                f"Google API error fetching header row {header_row}: {e}", exc_info=True
            )
            raise GoogleSheetError(f"API error reading header row {header_row}.") from e
        except Exception as e:
            logger.error(
                f"Unexpected error searching for column header: {e}", exc_info=True
            )
            raise GoogleSheetError("Unexpected error during column search.") from e

    def get_column_values(self, column_index: int, start_row: int = 2) -> List[str]:
        """
        Retrieves all values from a specific column index, starting from a given row.

        Args:
            column_index (int): The 1-based index of the column to retrieve.
            start_row (int): The row number to start fetching data from (1-indexed). Defaults to 2 (below header).

        Returns:
            List[str]: A list of cell values as strings.

        Raises:
            GoogleSheetError: If fetching column values fails.
        """
        try:
            logger.info(
                f"Fetching values from column index {column_index} starting at row {start_row}."
            )
            # Ensure we handle potential errors if col_values returns fewer rows than expected
            # or if the column index is out of bounds (gspread might handle this).
            all_values = self.sheet.col_values(column_index)

            # Slice the list correctly based on 1-based start_row index
            if start_row <= 0:
                logger.warning(
                    "start_row should be 1-based. Defaulting to fetching all values."
                )
                return all_values
            elif start_row > len(all_values):
                logger.warning(
                    f"start_row {start_row} is beyond the number of rows ({len(all_values)}) in column {column_index}. Returning empty list."
                )
                return []
            else:
                # Adjust for 0-based list indexing
                return all_values[start_row - 1 :]

        except APIError as e:
            logger.error(
                f"Google API error fetching column {column_index}: {e}", exc_info=True
            )
            raise GoogleSheetError(f"API error reading column {column_index}.") from e
        except Exception as e:
            logger.error(f"Unexpected error fetching column values: {e}", exc_info=True)
            raise GoogleSheetError("Unexpected error fetching column values.") from e

    def get_all_records(self, header_row: int = 1) -> List[Dict[str, str]]:
        """
        Fetches all rows from the sheet as a list of dictionaries (headers as keys).

        Args:
            header_row (int): The row number containing the headers (1-indexed).

        Returns:
            List[Dict[str, str]]: A list of dictionaries representing rows.
        """
        try:
            logger.info(f"Fetching all records using headers from row {header_row}...")
            # Note: gspread's get_all_records uses the *first* row as header by default.
            # If header_row is different, we might need to fetch manually or adjust.
            # For simplicity, assuming header_row=1 matches gspread's default behavior well enough,
            # but adjust if your header_row is truly different and causes issues.
            if header_row != 1:
                logger.warning(
                    f"get_all_records works best with header_row=1. Fetching based on row {header_row} might require custom implementation if issues arise."
                )
                # Potential custom implementation: fetch headers, fetch data range, zip them.
                headers = self.sheet.row_values(header_row)
                data_range = f"A{header_row + 1}:{self._col_string(self.sheet.col_count)}{self.sheet.row_count}"
                data_values = self.sheet.get(data_range)
                records = [dict(zip(headers, row)) for row in data_values]
                return records
            else:
                records = self.sheet.get_all_records()
                logger.info(f"Successfully fetched {len(records)} records.")
                return records

        except APIError as e:
            logger.error(f"Google API error fetching all records: {e}", exc_info=True)
            raise GoogleSheetError("API error fetching all records.") from e
        except Exception as e:
            logger.error(f"Unexpected error fetching all records: {e}", exc_info=True)
            raise GoogleSheetError("Unexpected error fetching all records.") from e
