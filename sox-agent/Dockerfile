FROM python:3.12-slim

WORKDIR /app

# Install dependencies first (cached layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

# Cloud Run Jobs execute the entrypoint to completion, then exit.
ENTRYPOINT ["python", "main.py"]
