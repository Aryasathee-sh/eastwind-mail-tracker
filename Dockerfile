FROM python:3.11-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8080
ENV HOST=0.0.0.0

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY excel_loader.py .
COPY query_engine.py .
COPY mcp_server.py .
COPY server_cloud.py .
COPY Eastwind_Proposals_-_Mail_Tracker.xlsx .

EXPOSE 8080

CMD ["python", "server_cloud.py"]
