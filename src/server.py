import os
import json
from mcp.server.fastmcp import FastMCP
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = os.environ.get("TOKEN_PATH", "/data/token.json")
CREDENTIALS_PATH = os.environ.get("CREDENTIALS_PATH", "/data/credentials.json")

mcp = FastMCP("google-sheets-mcp")


def get_service():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


# ---------------------------------------------------------------------------
# Original 4 tools
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
def create_spreadsheet(title: str, sheet_names: list = None) -> str:
    """Create a new Google Spreadsheet.

    Args:
        title: The title of the new spreadsheet.
        sheet_names: Optional list of sheet tab names (e.g. ['Sales', 'Inventory']).
                     Defaults to a single sheet named 'Sheet1'.
    """
    service = get_service()
    sheets = []
    if sheet_names:
        for i, name in enumerate(sheet_names):
            sheets.append({"properties": {"sheetId": i, "title": name}})
    else:
        sheets = [{"properties": {"title": "Sheet1"}}]

    body = {
        "properties": {"title": title},
        "sheets": sheets,
    }
    spreadsheet = service.spreadsheets().create(body=body).execute()
    sid = spreadsheet["spreadsheetId"]
    url = f"https://docs.google.com/spreadsheets/d/{sid}"
    return json.dumps({"spreadsheetId": sid, "url": url})


# ---------------------------------------------------------------------------
# New tools
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
    body = {
        "requests": [
            {"addSheet": {"properties": {"title": title}}}
        ]
    }
    result = (
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute()
    )
    new_sheet = result["replies"][0]["addSheet"]["properties"]
    return json.dumps(
        {"sheetId": new_sheet["sheetId"], "title": new_sheet["title"]},
        ensure_ascii=False,
    )


@mcp.tool()
def delete_rows(spreadsheet_id: str, sheet_id: int, start_index: int, end_index: int) -> str:
    """Delete rows from a Google Sheet by index (0-based).

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        sheet_id: The numeric ID of the sheet tab (use list_sheets to find it).
        start_index: The first row to delete (0-based, inclusive).
        end_index: The last row to delete (0-based, exclusive).
    """
    service = get_service()
    body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": end_index,
                    }
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return f"Successfully deleted rows {start_index} to {end_index - 1}."


@mcp.tool()
def find_and_replace(spreadsheet_id: str, find: str, replacement: str, sheet_id: int = None) -> str:
    """Find and replace text across a Google Spreadsheet or a specific sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        find: The text to search for.
        replacement: The text to replace with.
        sheet_id: Optional numeric sheet ID to limit the search to one tab.
                  If omitted, the replacement applies to all sheets.
    """
    service = get_service()
    find_replace = {
        "find": find,
        "replacement": replacement,
        "allSheets": sheet_id is None,
    }
    if sheet_id is not None:
        find_replace["sheetId"] = sheet_id

    body = {"requests": [{"findReplace": find_replace}]}
    result = (
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute()
    )
    stats = result["replies"][0].get("findReplace", {})
    replaced = stats.get("valuesChanged", 0)
    return f"Successfully replaced {replaced} occurrence(s) of '{find}' with '{replacement}'."


@mcp.tool()
def format_cells(
    spreadsheet_id: str,
    sheet_id: int,
    range_start_row: int,
    range_end_row: int,
    range_start_col: int,
    range_end_col: int,
    bold: bool = False,
    font_size: int = None,
    background_color: dict = None,
    text_color: dict = None,
) -> str:
    """Apply formatting to a range of cells in a Google Sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL).
        sheet_id: The numeric ID of the sheet tab (use list_sheets to find it).
        range_start_row: Start row index (0-based, inclusive).
        range_end_row: End row index (0-based, exclusive).
        range_start_col: Start column index (0-based, inclusive).
        range_end_col: End column index (0-based, exclusive).
        bold: Whether to make the text bold.
        font_size: Font size in points (e.g. 12).
        background_color: Dict with r, g, b keys (0.0-1.0) e.g. {"r": 1, "g": 0.9, "b": 0}.
        text_color: Dict with r, g, b keys (0.0-1.0) e.g. {"r": 0, "g": 0, "b": 0}.
    """
    service = get_service()

    cell_format = {}
    fields = []

    text_format = {}
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

    body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": range_start_row,
                        "endRowIndex": range_end_row,
                        "startColumnIndex": range_start_col,
                        "endColumnIndex": range_end_col,
                    },
                    "cell": {"userEnteredFormat": cell_format},
                    "fields": ",".join(fields),
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return "Formatting applied successfully."


if __name__ == "__main__":
    mcp.run()
