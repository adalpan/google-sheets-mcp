# Google Sheets MCP Server

An MCP (Model Context Protocol) server that lets AI agents interact with Google Sheets via OAuth 2.0.

## Tools

| # | Tool | Description |
|---|------|-------------|
| 1 | `read_sheet` | Read data from a range in a spreadsheet |
| 2 | `write_sheet` | Write or update cell values |
| 3 | `append_rows` | Append new rows to a sheet |
| 4 | `create_spreadsheet` | Create a new Google Spreadsheet |
| 5 | `list_sheets` | List all sheet tabs inside a spreadsheet |
| 6 | `get_spreadsheet_info` | Get metadata: title, sheet names, row/column counts |
| 7 | `clear_range` | Clear all values in a cell range |
| 8 | `add_sheet` | Add a new tab to an existing spreadsheet |
| 9 | `delete_rows` | Delete rows by index (0-based) |
| 10 | `find_and_replace` | Search and replace text across sheets |
| 11 | `format_cells` | Apply bold, font size, and background/text colors |
| 12 | `rename_sheet` | Rename an existing sheet tab |
| 13 | `delete_sheet` | Delete a sheet tab from a spreadsheet |
| 14 | `duplicate_sheet` | Duplicate a sheet tab within a spreadsheet |
| 15 | `insert_rows_columns` | Insert empty rows or columns at a specific position |
| 16 | `delete_columns` | Delete columns by index (0-based) |
| 17 | `sort_range` | Sort a range by a specific column (asc or desc) |
| 18 | `freeze_rows_columns` | Freeze header rows/columns so they stay visible |
| 19 | `add_conditional_formatting` | Highlight cells based on rules (e.g. red if negative) |

## Setup

### 1. Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use an existing one)
3. Enable the **Google Sheets API**
4. Go to **APIs & Services → Credentials**
5. Create an **OAuth 2.0 Client ID** (Desktop app type)
6. Download the JSON file and save it as `credentials.json`

### 2. Run with Docker

```bash
docker run -i --rm \
  -v /path/to/your/data:/data \
  mcp/google-sheets
```

On first run, a browser window will open for you to authorize access. A `token.json` will be saved in your `/data` folder for future runs.

### 3. Connect to Claude Desktop

Add this to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-v", "/path/to/your/data:/data",
        "mcp/google-sheets"
      ]
    }
  }
}
```

### 4. Run locally without Docker

```bash
git clone https://github.com/adalpan/google-sheets-mcp
cd google-sheets-mcp
pip install -r requirements.txt
mkdir data
cp /path/to/credentials.json data/
TOKEN_PATH=data/token.json CREDENTIALS_PATH=data/credentials.json python src/server.py
```

## Testing with MCP Inspector

```bash
npx @modelcontextprotocol/inspector python src/server.py
```

Opens a web UI to call and test each tool interactively.

## Usage Examples

- *"List all the sheets in my spreadsheet"*
- *"Read cells A1 to D10 from Sheet1"*
- *"Append a new row with ['Alice', 30, 'Berlin'] to Sheet1"*
- *"Create a new spreadsheet called 'Q1 Budget' with sheets Sales and Expenses"*
- *"Update cell B2 to 'Completed'"*
- *"Clear the range Sheet1!A1:Z100"*
- *"Add a new tab called 'Archive'"*
- *"Delete rows 5 to 10 from Sheet1"*
- *"Find and replace 'Pending' with 'Done' across all sheets"*
- *"Make the header row bold with a yellow background"*
- *"Rename the sheet 'Sheet1' to 'Sales Q1'"*
- *"Delete the tab called 'Draft'"*
- *"Duplicate the 'Template' sheet and call it 'January'"*
- *"Insert 3 empty rows before row 5"*
- *"Delete columns C and D"*
- *"Sort the data range by the second column in descending order"*
- *"Freeze the first row and first column"*
- *"Highlight in red all cells in column B that are less than 0"*

## License

MIT
