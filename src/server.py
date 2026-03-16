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


if __name__ == "__main__":
    mcp.run()
