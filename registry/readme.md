# Google Sheets MCP Server

Read, write, append rows, and create Google Spreadsheets using natural language via AI agents.

## Authentication

This server uses **OAuth 2.0**. On first run, a browser window opens to authorize your Google account. A `token.json` is saved in the `/data` volume for future runs.

You must provide a `credentials.json` file (OAuth Desktop App) from [Google Cloud Console](https://console.cloud.google.com/).

## Configuration

Mount a local folder with your credentials:

```bash
docker run -i --rm \
  -v /path/to/data:/data \
  mcp/google-sheets
```

## Tools

- **read_sheet** — Read a cell range from a spreadsheet
- **write_sheet** — Write or update cell values
- **append_rows** — Append new rows to a sheet
- **create_spreadsheet** — Create a new spreadsheet with optional sheet names

## Source

[github.com/adalpan/google-sheets-mcp](https://github.com/adalpan/google-sheets-mcp)
