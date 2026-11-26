FROM python:3.12-slim

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend, trained knowledge base, and prebuilt Flutter Web UI
COPY backend/ ./backend/
COPY data/trained_formats/ ./data/trained_formats/
COPY huliot_flutter_app/build/web/ ./huliot_flutter_app/build/web/

# Ensure raw_boq directory exists
RUN mkdir -p ./data/raw_boq

WORKDIR /app/backend
EXPOSE 8000

CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${PORT:-8000}"]
