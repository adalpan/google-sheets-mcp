# Google Sheets MCP Server

A fully-featured [MCP (Model Context Protocol)](https://modelcontextprotocol.io) server that lets AI agents interact with Google Sheets via the official Google Sheets API v4, using OAuth 2.0 authentication.

Built with Python and [FastMCP](https://github.com/jlowin/fastmcp). Designed to run as a Docker container and integrate with AI clients like Claude Desktop, Cursor, and others.

## Features

- **28 tools** covering the full Google Sheets API v4
- OAuth 2.0 authentication (token cached after first login)
- Runs as an isolated Docker container
- Works with Claude Desktop, Cursor, Zed, and any MCP-compatible client

## Tools (28 total)

### 📋 Values — Reading & Writing
| # | Tool | Description |
|---|------|-------------|
| 1 | `read_sheet` | Read data from a single range |
| 2 | `write_sheet` | Write or update cell values |
| 3 | `append_rows` | Append new rows to a sheet |
| 4 | `batch_get_values` | Read multiple ranges in a single API call |
| 5 | `batch_update_values` | Write to multiple ranges in a single API call |
| 6 | `batch_clear_values` | Clear multiple ranges in a single API call |
| 7 | `clear_range` | Clear all values in a single range |

### 🗂️ Spreadsheet & Sheet Management
| # | Tool | Description |
|---|------|-------------|
| 8 | `create_spreadsheet` | Create a new Google Spreadsheet |
| 9 | `get_spreadsheet_info` | Get title, sheet names, row/column counts |
| 10 | `list_sheets` | List all sheet tabs |
| 11 | `add_sheet` | Add a new sheet tab |
| 12 | `delete_sheet` | Delete a sheet tab |
| 13 | `rename_sheet` | Rename a sheet tab |
| 14 | `duplicate_sheet` | Duplicate a sheet tab within the same spreadsheet |
| 15 | `copy_sheet_to` | Copy a sheet tab to a different spreadsheet |

### 🔢 Rows & Columns
| # | Tool | Description |
|---|------|-------------|
| 16 | `insert_rows_columns` | Insert empty rows or columns at a specific position |
| 17 | `delete_rows` | Delete rows by index |
| 18 | `delete_columns` | Delete columns by index |
| 19 | `move_dimension` | Move rows or columns to a new position |
| 20 | `insert_range` | Insert a blank range and shift cells down or right |
| 21 | `delete_range` | Delete a range and shift cells up or left |

### 🧹 Data Cleaning
| # | Tool | Description |
|---|------|-------------|
| 22 | `delete_duplicates` | Remove duplicate rows from a range |
| 23 | `trim_whitespace` | Trim extra spaces from cells in a range |
| 24 | `find_and_replace` | Search and replace text across sheets |
| 25 | `sort_range` | Sort a range by a specific column (asc or desc) |

### 🎨 Formatting
| # | Tool | Description |
|---|------|-------------|
| 26 | `format_cells` | Apply bold, font size, background and text colors |
| 27 | `freeze_rows_columns` | Freeze header rows and/or columns |
| 28 | `add_conditional_formatting` | Highlight cells based on rules (e.g. red if negative) |

## Setup

### 1. Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use an existing one)
3. Enable the **Google Sheets API**
4. Go to **APIs & Services → Credentials**
5. Click **Create Credentials → OAuth 2.0 Client ID**
6. Choose **Desktop app** as the application type
7. Download the JSON file and save it as `credentials.json` inside a `data/` folder

### 2. Run with Docker

```bash
docker run -i --rm \
  -v /path/to/your/data:/data \
  mcp/google-sheets
```

On first run, a browser window will open for you to sign in to Google and authorize access. A `token.json` is saved in your `/data` folder for future runs — you won't need to log in again.

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

Restart Claude Desktop — you can now talk to your Google Sheets directly!

### 4. Run Locally Without Docker

```bash
git clone https://github.com/adalpan/google-sheets-mcp
cd google-sheets-mcp
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
mkdir data
cp /path/to/credentials.json data/
TOKEN_PATH=data/token.json CREDENTIALS_PATH=data/credentials.json python src/server.py
```

## Testing with MCP Inspector

The [MCP Inspector](https://github.com/modelcontextprotocol/inspector) lets you call and test each tool interactively via a web UI:

```bash
source venv/bin/activate
npx @modelcontextprotocol/inspector \
  env TOKEN_PATH=data/token.json \
  CREDENTIALS_PATH=data/credentials.json \
  python src/server.py
```

Open the URL shown in the terminal, set **Transport Type** to `STDIO`, and click **Connect**.

## Usage Examples

**Reading & Writing**
- *"Read cells A1 to D10 from Sheet1"*
- *"Write a header row with Name, Age, City to Sheet1"*
- *"Append a new row with Alice, 30, Berlin"*
- *"Read Sheet1!A1:C5 and Sheet2!B1:D3 in one call"*
- *"Update both Sheet1!A1 and Sheet2!A1 at once"*

**Sheet Management**
- *"Create a spreadsheet called Q1 Budget with sheets Sales and Expenses"*
- *"List all the tabs in my spreadsheet"*
- *"Rename Sheet1 to Sales Q1"*
- *"Duplicate the Template sheet and call it January"*
- *"Copy this sheet to my other spreadsheet"*

**Rows & Columns**
- *"Insert 3 empty rows before row 5"*
- *"Delete columns C and D"*
- *"Move column 3 to position 0"*
- *"Insert a blank range at B2:D5 and shift cells down"*

**Data Cleaning**
- *"Remove all duplicate rows from my data range"*
- *"Trim whitespace from all cells in column A"*
- *"Find and replace Pending with Done across all sheets"*
- *"Sort the data by the second column descending"*

**Formatting**
- *"Make the header row bold with a yellow background"*
- *"Freeze the first row and first column"*
- *"Highlight in red all cells in column B that are less than 0"*

## License

MIT License — Copyright (c) 2026 Adalberto Perez Salas (pan)
