from __future__ import annotations

from pathlib import Path


def append_report_row(
    credentials_path: Path,
    spreadsheet_name: str,
    sheet_name: str,
    row_values: list[str | bool],
) -> None:
    """Insert one metadata row directly under the sheet header row.

    Parameters:
    - `credentials_path`: Service-account JSON credential file.
    - `spreadsheet_name`: Human-readable spreadsheet title, for example
      `fpc-reports`.
    - `sheet_name`: Worksheet name inside the spreadsheet, for example `test`.
    - `row_values`: Ordered row values matching the expected sheet columns.

    Behavior:
    - Uses `gspread.service_account(...)` for authentication.
    - Opens the spreadsheet by name.
    - Selects the target worksheet by name.
    - Inserts a new row at index 2 so the header row remains at the top.
    """
    try:
        import gspread
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("gspread is required for sheet updates. Install it in the Python environment used for publish-bot.") from exc

    client = gspread.service_account(filename=str(credentials_path))
    spreadsheet = client.open(spreadsheet_name)
    worksheet = spreadsheet.worksheet(sheet_name)
    worksheet.insert_row(row_values, index=2, value_input_option="USER_ENTERED")
