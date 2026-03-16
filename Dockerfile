FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ ./src/

# /data is where OAuth credentials and token will be mounted at runtime
VOLUME ["/data"]

ENV TOKEN_PATH=/data/token.json
ENV CREDENTIALS_PATH=/data/credentials.json

CMD ["python", "src/server.py"]
