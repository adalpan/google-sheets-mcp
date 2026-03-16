import os
import json
from typing import Optional
from mcp.server.fastmcp import FastMCP
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = os.environ.get("TOKEN_PATH", "/data/token.json")
CREDENTIALS_PATH = os.environ.get("CREDENTIALS_PATH", "/data/credentials.json")
SERVICE_ACCOUNT_PATH = os.environ.get("SERVICE_ACCOUNT_PATH", "/data/service_account.json")
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

mcp = FastMCP("google-sheets-mcp")


def get_service():
    """Auto-detect authentication method (priority order):
    1. GOOGLE_SERVICE_ACCOUNT_JSON env var (Docker Desktop config UI)
    2. service_account.json file (Docker volume mount)
    3. OAuth 2.0 credentials.json file (local/personal use)
    """
    # --- 1. Service Account via env var (Docker Desktop config UI) ---
    if SERVICE_ACCOUNT_JSON:
        info = json.loads(SERVICE_ACCOUNT_JSON)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
        return build("sheets", "v4", credentials=creds)

    # --- 2. Service Account via file (volume mount) ---
    if os.path.exists(SERVICE_ACCOUNT_PATH):
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_PATH, scopes=SCOPES
        )
        return build("sheets", "v4", credentials=creds)

    # --- 3. OAuth 2.0 (local/personal use) ---
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                raise FileNotFoundError(
                    "No credentials found. Please provide one of:\n"
                    "  1. GOOGLE_SERVICE_ACCOUNT_JSON env var (Docker Desktop config UI)\n"
                    f"  2. Service Account file: {SERVICE_ACCOUNT_PATH}\n"
                    f"  3. OAuth 2.0 credentials file: {CREDENTIALS_PATH}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


# ---------------------------------------------------------------------------
# Batch 1 — original 4 tools
# ---------------------------------------------------------------------------

@mcp.tool()
def read_sheet(spreadsheet_id: str, range: str) -> str:
    """Read data from a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        range: The A1 notation range to read (e.g. 'Sheet1!A1:D10').
    """
    service = get_service()
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range)
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return "No data found in the specified range."
    return json.dumps(values, ensure_ascii=False)


@mcp.tool()
def write_sheet(spreadsheet_id: str, range: str, values: list) -> str:
    """Write or update cell values in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        range: The A1 notation range to write (e.g. 'Sheet1!A1').
        values: 2D list of values to write (e.g. [['Name', 'Age'], ['Alice', 30]]).
    """
    service = get_service()
    body = {"values": values}
    result = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range,
            valueInputOption="USER_ENTERED",
            body=body,
        )
        .execute()
    )
    updated = result.get("updatedCells", 0)
    return f"Successfully updated {updated} cell(s)."


@mcp.tool()
def append_rows(spreadsheet_id: str, range: str, values: list) -> str:
    """Append rows to a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        range: The A1 notation range to append after (e.g. 'Sheet1!A1').
        values: 2D list of rows to append (e.g. [['Alice', 30], ['Bob', 25]]).
    """
    service = get_service()
    body = {"values": values}
    result = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body,
        )
        .execute()
    )
    updates = result.get("updates", {})
    updated = updates.get("updatedRows", 0)
    return f"Successfully appended {updated} row(s)."


@mcp.tool()
def create_spreadsheet(title: str, sheet_names: Optional[list] = None) -> str:
    """Create a new Google Spreadsheet.

    Args:
        title: The title of the new spreadsheet.
        sheet_names: Optional list of sheet tab names (e.g. ['Sales', 'Inventory']).
    """
    service = get_service()
    sheets = []
    if sheet_names:
        for i, name in enumerate(sheet_names):
            sheets.append({"properties": {"sheetId": i, "title": name}})
    else:
        sheets = [{"properties": {"title": "Sheet1"}}]
    body = {"properties": {"title": title}, "sheets": sheets}
    spreadsheet = service.spreadsheets().create(body=body).execute()
    sid = spreadsheet["spreadsheetId"]
    url = f"https://docs.google.com/spreadsheets/d/{sid}"
    return json.dumps({"spreadsheetId": sid, "url": url})


# ---------------------------------------------------------------------------
# Batch 2 — 7 tools
# ---------------------------------------------------------------------------

@mcp.tool()
def list_sheets(spreadsheet_id: str) -> str:
    """List all sheet tabs inside a Google Spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
    """
    service = get_service()
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = [
        {"sheetId": s["properties"]["sheetId"], "title": s["properties"]["title"]}
        for s in result.get("sheets", [])
    ]
    return json.dumps(sheets, ensure_ascii=False)


@mcp.tool()
def get_spreadsheet_info(spreadsheet_id: str) -> str:
    """Get metadata about a Google Spreadsheet (title, sheets, row/column counts).

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
    """
    service = get_service()
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    info = {
        "spreadsheetId": result["spreadsheetId"],
        "title": result["properties"]["title"],
        "url": result["spreadsheetUrl"],
        "sheets": [
            {
                "sheetId": s["properties"]["sheetId"],
                "title": s["properties"]["title"],
                "rowCount": s["properties"]["gridProperties"]["rowCount"],
                "columnCount": s["properties"]["gridProperties"]["columnCount"],
            }
            for s in result.get("sheets", [])
        ],
    }
    return json.dumps(info, ensure_ascii=False)


@mcp.tool()
def clear_range(spreadsheet_id: str, range: str) -> str:
    """Clear all values in a range of a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        range: The A1 notation range to clear (e.g. 'Sheet1!A1:D10').
    """
    service = get_service()
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range=range, body={}
    ).execute()
    return f"Successfully cleared range '{range}'."


@mcp.tool()
def add_sheet(spreadsheet_id: str, title: str) -> str:
    """Add a new sheet tab to an existing Google Spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        title: The name for the new sheet tab.
    """
    service = get_service()
    body = {"requests": [{"addSheet": {"properties": {"title": title}}}]}
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    new_sheet = result["replies"][0]["addSheet"]["properties"]
    return json.dumps({"sheetId": new_sheet["sheetId"], "title": new_sheet["title"]})


@mcp.tool()
def delete_rows(spreadsheet_id: str, sheet_id: int, start_index: int, end_index: int) -> str:
    """Delete rows from a Google Sheet by index (0-based).

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab (use list_sheets to find it).
        start_index: First row to delete (0-based, inclusive).
        end_index: Last row to delete (0-based, exclusive).
    """
    service = get_service()
    body = {"requests": [{"deleteDimension": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": start_index, "endIndex": end_index}}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Successfully deleted rows {start_index} to {end_index - 1}."


@mcp.tool()
def find_and_replace(spreadsheet_id: str, find: str, replacement: str, sheet_id: Optional[int] = None) -> str:
    """Find and replace text across a Google Spreadsheet or a specific sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        find: The text to search for.
        replacement: The text to replace with.
        sheet_id: Optional numeric sheet ID to limit the search to one tab.
    """
    service = get_service()
    find_replace = {"find": find, "replacement": replacement, "allSheets": sheet_id is None}
    if sheet_id is not None:
        find_replace["sheetId"] = sheet_id
    body = {"requests": [{"findReplace": find_replace}]}
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    replaced = result["replies"][0].get("findReplace", {}).get("valuesChanged", 0)
    return f"Successfully replaced {replaced} occurrence(s) of '{find}' with '{replacement}'."


@mcp.tool()
def format_cells(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    bold: bool = False,
    font_size: Optional[int] = None,
    background_color: Optional[dict] = None,
    text_color: Optional[dict] = None,
) -> str:
    """Apply formatting to a range of cells in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        bold: Whether to make the text bold.
        font_size: Font size in points.
        background_color: Dict with r, g, b keys (0.0-1.0).
        text_color: Dict with r, g, b keys (0.0-1.0).
    """
    service = get_service()
    cell_format, fields, text_format = {}, [], {}
    if bold:
        text_format["bold"] = True
        fields.append("userEnteredFormat.textFormat.bold")
    if font_size:
        text_format["fontSize"] = font_size
        fields.append("userEnteredFormat.textFormat.fontSize")
    if text_color:
        text_format["foregroundColor"] = text_color
        fields.append("userEnteredFormat.textFormat.foregroundColor")
    if text_format:
        cell_format["textFormat"] = text_format
    if background_color:
        cell_format["backgroundColor"] = background_color
        fields.append("userEnteredFormat.backgroundColor")
    if not fields:
        return "No formatting options specified."
    body = {"requests": [{"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}, "cell": {"userEnteredFormat": cell_format}, "fields": ",".join(fields)}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return "Formatting applied successfully."


# ---------------------------------------------------------------------------
# Batch 3 — 8 tools
# ---------------------------------------------------------------------------

@mcp.tool()
def rename_sheet(spreadsheet_id: str, sheet_id: int, new_title: str) -> str:
    """Rename an existing sheet tab in a Google Spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab (use list_sheets to find it).
        new_title: The new name for the sheet tab.
    """
    service = get_service()
    body = {"requests": [{"updateSheetProperties": {"properties": {"sheetId": sheet_id, "title": new_title}, "fields": "title"}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Sheet renamed to '{new_title}'."


@mcp.tool()
def delete_sheet(spreadsheet_id: str, sheet_id: int) -> str:
    """Delete a sheet tab from a Google Spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab to delete.
    """
    service = get_service()
    body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Sheet with ID {sheet_id} deleted successfully."


@mcp.tool()
def duplicate_sheet(spreadsheet_id: str, sheet_id: int, new_title: str, insert_at: Optional[int] = None) -> str:
    """Duplicate a sheet tab within a Google Spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet to duplicate.
        new_title: Name for the duplicated sheet.
        insert_at: Optional index position to insert the duplicate (0-based).
    """
    service = get_service()
    req = {"duplicateSheet": {"sourceSheetId": sheet_id, "newSheetName": new_title}}
    if insert_at is not None:
        req["duplicateSheet"]["insertSheetIndex"] = insert_at
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [req]}).execute()
    new_sheet = result["replies"][0]["duplicateSheet"]["properties"]
    return json.dumps({"sheetId": new_sheet["sheetId"], "title": new_sheet["title"]})


@mcp.tool()
def insert_rows_columns(spreadsheet_id: str, sheet_id: int, dimension: str, start_index: int, count: int) -> str:
    """Insert empty rows or columns at a specific position in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        dimension: Either 'ROWS' or 'COLUMNS'.
        start_index: Index before which to insert (0-based).
        count: Number of rows or columns to insert.
    """
    service = get_service()
    body = {"requests": [{"insertDimension": {"range": {"sheetId": sheet_id, "dimension": dimension.upper(), "startIndex": start_index, "endIndex": start_index + count}, "inheritFromBefore": False}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Successfully inserted {count} {dimension.lower()} at index {start_index}."


@mcp.tool()
def delete_columns(spreadsheet_id: str, sheet_id: int, start_index: int, end_index: int) -> str:
    """Delete columns from a Google Sheet by index (0-based).

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        start_index: First column to delete (0-based, inclusive).
        end_index: Last column to delete (0-based, exclusive).
    """
    service = get_service()
    body = {"requests": [{"deleteDimension": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": start_index, "endIndex": end_index}}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Successfully deleted columns {start_index} to {end_index - 1}."


@mcp.tool()
def sort_range(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    sort_column_index: int, ascending: bool = True,
) -> str:
    """Sort a range of cells by a specific column.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        sort_column_index: Column index to sort by (0-based).
        ascending: True for A-Z, False for Z-A.
    """
    service = get_service()
    body = {"requests": [{"sortRange": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}, "sortSpecs": [{"dimensionIndex": sort_column_index, "sortOrder": "ASCENDING" if ascending else "DESCENDING"}]}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Range sorted {'ascending' if ascending else 'descending'} by column index {sort_column_index}."


@mcp.tool()
def freeze_rows_columns(spreadsheet_id: str, sheet_id: int, frozen_row_count: int = 0, frozen_column_count: int = 0) -> str:
    """Freeze rows and/or columns in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        frozen_row_count: Number of rows to freeze from the top (0 to unfreeze).
        frozen_column_count: Number of columns to freeze from the left (0 to unfreeze).
    """
    service = get_service()
    body = {"requests": [{"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": frozen_row_count, "frozenColumnCount": frozen_column_count}}, "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Frozen {frozen_row_count} row(s) and {frozen_column_count} column(s)."


@mcp.tool()
def add_conditional_formatting(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    condition_type: str, condition_value: str, background_color: dict,
) -> str:
    """Add a conditional formatting rule to a range in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        condition_type: One of: NUMBER_GREATER, NUMBER_LESS, NUMBER_EQ, TEXT_CONTAINS, TEXT_EQ, BLANK, NOT_BLANK.
        condition_value: Value to compare against (e.g. '0' or 'Done').
        background_color: Dict with r, g, b keys (0.0-1.0) for highlight color.
    """
    service = get_service()
    body = {"requests": [{"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}], "booleanRule": {"condition": {"type": condition_type.upper(), "values": [{"userEnteredValue": condition_value}]}, "format": {"backgroundColor": background_color}}}, "index": 0}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Conditional formatting rule added: {condition_type} '{condition_value}'."


# ---------------------------------------------------------------------------
# Batch 4 — 9 tools
# ---------------------------------------------------------------------------

@mcp.tool()
def batch_get_values(spreadsheet_id: str, ranges: list) -> str:
    """Read multiple ranges at once from a Google Spreadsheet in a single API call.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        ranges: List of A1 notation ranges (e.g. ['Sheet1!A1:C10', 'Sheet2!B2:D5']).
    """
    service = get_service()
    result = (
        service.spreadsheets()
        .values()
        .batchGet(spreadsheetId=spreadsheet_id, ranges=ranges)
        .execute()
    )
    output = [
        {"range": vr.get("range"), "values": vr.get("values", [])}
        for vr in result.get("valueRanges", [])
    ]
    return json.dumps(output, ensure_ascii=False)


@mcp.tool()
def batch_update_values(spreadsheet_id: str, updates: list) -> str:
    """Write data to multiple ranges in a single API call.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        updates: List of dicts with 'range' and 'values' keys.
                 Example: [{'range': 'Sheet1!A1', 'values': [['Name', 'Age']]}, ...]
    """
    service = get_service()
    body = {
        "valueInputOption": "USER_ENTERED",
        "data": [{"range": u["range"], "values": u["values"]} for u in updates],
    }
    result = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    total = result.get("totalUpdatedCells", 0)
    return f"Successfully updated {total} cell(s) across {len(updates)} range(s)."


@mcp.tool()
def batch_clear_values(spreadsheet_id: str, ranges: list) -> str:
    """Clear multiple ranges at once in a single API call.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        ranges: List of A1 notation ranges to clear.
    """
    service = get_service()
    service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id, body={"ranges": ranges}
    ).execute()
    return f"Successfully cleared {len(ranges)} range(s)."


@mcp.tool()
def copy_sheet_to(source_spreadsheet_id: str, sheet_id: int, destination_spreadsheet_id: str) -> str:
    """Copy a sheet tab to a different Google Spreadsheet.

    Args:
        source_spreadsheet_id: The ID of the source spreadsheet.
        sheet_id: Numeric ID of the sheet to copy (use list_sheets to find it).
        destination_spreadsheet_id: The ID of the destination spreadsheet.
    """
    service = get_service()
    result = (
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=source_spreadsheet_id,
            sheetId=sheet_id,
            body={"destinationSpreadsheetId": destination_spreadsheet_id},
        )
        .execute()
    )
    return json.dumps({"newSheetId": result["sheetId"], "newSheetTitle": result["title"]})


@mcp.tool()
def move_dimension(
    spreadsheet_id: str, sheet_id: int, dimension: str,
    start_index: int, end_index: int, destination_index: int,
) -> str:
    """Move rows or columns to a new position in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        dimension: Either 'ROWS' or 'COLUMNS'.
        start_index: Start index of the range to move (0-based, inclusive).
        end_index: End index of the range to move (0-based, exclusive).
        destination_index: Index to move the rows/columns to.
    """
    service = get_service()
    body = {"requests": [{"moveDimension": {"source": {"sheetId": sheet_id, "dimension": dimension.upper(), "startIndex": start_index, "endIndex": end_index}, "destinationIndex": destination_index}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Moved {dimension.lower()} {start_index}-{end_index - 1} to position {destination_index}."


@mcp.tool()
def delete_duplicates(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    comparison_columns: Optional[list] = None,
) -> str:
    """Remove duplicate rows from a range in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        comparison_columns: Optional list of 0-based column indices to compare.
                            If omitted, all columns are compared.
    """
    service = get_service()
    request = {"deleteDuplicates": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}}}
    if comparison_columns:
        request["deleteDuplicates"]["comparisonColumns"] = [
            {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1}
            for c in comparison_columns
        ]
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [request]}).execute()
    removed = result["replies"][0].get("deleteDuplicates", {}).get("duplicatesRemovedCount", 0)
    return f"Removed {removed} duplicate row(s)."


@mcp.tool()
def trim_whitespace(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
) -> str:
    """Trim leading and trailing whitespace from all cells in a range.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
    """
    service = get_service()
    body = {"requests": [{"trimWhitespace": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}}}]}
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    trimmed = result["replies"][0].get("trimWhitespace", {}).get("cellsChangedCount", 0)
    return f"Trimmed whitespace from {trimmed} cell(s)."


@mcp.tool()
def insert_range(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    shift_dimension: str = "ROWS",
) -> str:
    """Insert a blank range and shift existing cells down or to the right.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        shift_dimension: 'ROWS' to shift cells down, 'COLUMNS' to shift right.
    """
    service = get_service()
    body = {"requests": [{"insertRange": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}, "shiftDimension": shift_dimension.upper()}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Inserted blank range and shifted existing cells {shift_dimension.lower()}."


@mcp.tool()
def delete_range(
    spreadsheet_id: str, sheet_id: int,
    range_start_row: int, range_end_row: int,
    range_start_col: int, range_end_col: int,
    shift_dimension: str = "ROWS",
) -> str:
    """Delete a range of cells and shift remaining cells up or to the left.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: Numeric ID of the sheet tab.
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        shift_dimension: 'ROWS' to shift cells up, 'COLUMNS' to shift left.
    """
    service = get_service()
    body = {"requests": [{"deleteRange": {"range": {"sheetId": sheet_id, "startRowIndex": range_start_row, "endRowIndex": range_end_row, "startColumnIndex": range_start_col, "endColumnIndex": range_end_col}, "shiftDimension": shift_dimension.upper()}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Deleted range and shifted remaining cells {shift_dimension.lower()}."


if __name__ == "__main__":
    mcp.run()
