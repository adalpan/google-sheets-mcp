FROM python:3.12-slim

LABEL org.opencontainers.image.title="Google Sheets MCP Server" \
      org.opencontainers.image.description="A fully-featured MCP server for Google Sheets with 28 tools covering reading, writing, batch operations, sheet management, row/column manipulation, data cleaning, and formatting." \
      org.opencontainers.image.source="https://github.com/adalpan/google-sheets-mcp" \
      org.opencontainers.image.licenses="MIT" \
      org.opencontainers.image.authors="Adalberto Perez Salas (pan)"

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ ./src/

# Create a non-root user and switch to it
RUN useradd --create-home --shell /bin/bash mcpuser
RUN mkdir -p /data && chown mcpuser:mcpuser /data
USER mcpuser

# /data is where OAuth credentials and token will be mounted at runtime
VOLUME ["/data"]

ENV TOKEN_PATH=/data/token.json
ENV CREDENTIALS_PATH=/data/credentials.json

CMD ["python", "src/server.py"]
