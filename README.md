# Google Sheets MCP Server

A fully-featured [MCP (Model Context Protocol)](https://modelcontextprotocol.io) server that lets AI agents interact with Google Sheets via the official Google Sheets API v4.

Supports **two authentication methods**: Service Account (recommended for Docker) and OAuth 2.0 (recommended for local/personal use).

Built with Python and [FastMCP](https://github.com/jlowin/fastmcp). Designed to run as a Docker container and integrate with AI clients like Claude Desktop, Cursor, and others.

## Features

- **28 tools** covering the full Google Sheets API v4
- **Two auth methods**: Service Account and OAuth 2.0 — auto-detected
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

---

## Authentication

The server **auto-detects** which method to use based on which file is present in `/data`:

| File | Method | Best for |
|------|--------|----------|
| `service_account.json` | Service Account | ✅ Docker, teams, automation |
| `credentials.json` | OAuth 2.0 | ✅ Personal use, local dev |

---

## Option A — Service Account ✅ Recommended for Docker

Service Accounts are ideal for Docker because they require no browser login — just a single JSON key file.

### Step 1 — Create a Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click **Select a project → New Project**
3. Give it a name (e.g. `my-sheets-mcp`) and click **Create**

### Step 2 — Enable the Google Sheets API
1. In your project, go to **APIs & Services → Library**
2. Search for **Google Sheets API**
3. Click it and then click **Enable**

### Step 3 — Create a Service Account
1. Go to **APIs & Services → Credentials**
2. Click **Create Credentials → Service Account**
3. Give it a name (e.g. `sheets-mcp`) and click **Create and Continue**
4. Skip the optional steps and click **Done**

### Step 4 — Download the JSON Key
1. In the Credentials page, click on the service account you just created
2. Go to the **Keys** tab
3. Click **Add Key → Create new key**
4. Choose **JSON** and click **Create**
5. A file will be downloaded — rename it to `service_account.json`

### Step 5 — Share your Google Sheet with the Service Account
> ⚠️ This is the most important step! The service account can only access sheets that are explicitly shared with it.

1. Open the downloaded `service_account.json` and find the `client_email` field. It looks like:
   ```
   sheets-mcp@my-sheets-mcp.iam.gserviceaccount.com
   ```
2. Open your Google Sheet
3. Click **Share** (top right)
4. Paste the `client_email` address and give it **Editor** access
5. Click **Send**

### Step 6 — Run with Docker
```bash
mkdir -p data
cp /path/to/service_account.json data/

docker run -i --rm \
  -v $(pwd)/data:/data \
  mcp/google-sheets
```

No browser login needed! 🎉

---

## Option B — OAuth 2.0 ✅ Recommended for Local / Personal Use

OAuth 2.0 lets you log in with your personal Google account via a browser. Best for local development and personal use.

### Step 1 — Create a Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click **Select a project → New Project**
3. Give it a name and click **Create**

### Step 2 — Enable the Google Sheets API
1. Go to **APIs & Services → Library**
2. Search for **Google Sheets API** and click **Enable**

### Step 3 — Configure OAuth Consent Screen
1. Go to **APIs & Services → OAuth consent screen**
2. Choose **External** and click **Create**
3. Fill in the **App name** and your **email** in the required fields
4. Click **Save and Continue** through all steps

### Step 4 — Create OAuth 2.0 Credentials
1. Go to **APIs & Services → Credentials**
2. Click **Create Credentials → OAuth 2.0 Client ID**
3. Choose **Desktop app** as the application type
4. Give it a name and click **Create**
5. Click **Download JSON** and rename the file to `credentials.json`

### Step 5 — First-time Login
On first run, a browser window will open for you to sign in to Google and authorize access.
A `token.json` is saved in your `data/` folder — you won't need to log in again.

### Step 6 — Run Locally
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

---

## Connect to Claude Desktop

### Using Docker (Service Account)
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

### Using Python directly (OAuth 2.0)
```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "python",
      "args": ["/path/to/google-sheets-mcp/src/server.py"],
      "env": {
        "TOKEN_PATH": "/path/to/data/token.json",
        "CREDENTIALS_PATH": "/path/to/data/credentials.json"
      }
    }
  }
}
```

## Testing with MCP Inspector

```bash
source venv/bin/activate
npx @modelcontextprotocol/inspector \
  env TOKEN_PATH=data/token.json \
  CREDENTIALS_PATH=data/credentials.json \
  python src/server.py
```

Open the URL shown in the terminal, set **Transport Type** to `STDIO`, and click **Connect**.

## Usage Examples

- *"Read cells A1 to D10 from Sheet1"*
- *"Write a header row with Name, Age, City to Sheet1"*
- *"Append a new row with Alice, 30, Berlin"*
- *"Create a spreadsheet called Q1 Budget with sheets Sales and Expenses"*
- *"Rename Sheet1 to Sales Q1"*
- *"Remove all duplicate rows from my data range"*
- *"Sort the data by the second column descending"*
- *"Highlight in red all cells less than 0"*
- *"Freeze the first row and first column"*

## License

MIT License — Copyright (c) 2026 Adalberto Perez Salas (pan)
