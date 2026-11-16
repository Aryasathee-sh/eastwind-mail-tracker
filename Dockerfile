FROM python:3.11-slim

WORKDIR /app

# Prevent Python from writing .pyc and enable unbuffered output
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8080
ENV HOST=0.0.0.0

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files and workbook
COPY excel_loader.py .
COPY mcp_server.py .
COPY server_cloud.py .
COPY Eastwind_Proposals_-_Mail_Tracker.xlsx .

EXPOSE 8080

CMD ["python", "server_cloud.py"]
