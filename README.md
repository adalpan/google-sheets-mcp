# Google Sheets MCP Server

An MCP (Model Context Protocol) server that lets AI agents interact with Google Sheets via OAuth 2.0.

## Tools

| Tool | Description |
|------|-------------|
| `read_sheet` | Read data from a range in a spreadsheet |
| `write_sheet` | Write or update cell values |
| `append_rows` | Append new rows to a sheet |
| `create_spreadsheet` | Create a new Google Spreadsheet |

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

## Usage Examples

- *"Read cells A1 to D10 from my spreadsheet"*
- *"Append a new row with ['Alice', 30, 'Berlin'] to Sheet1"*
- *"Create a new spreadsheet called 'Q1 Budget' with sheets Sales and Expenses"*
- *"Update cell B2 in my spreadsheet to 'Completed'"*

## License

MIT
