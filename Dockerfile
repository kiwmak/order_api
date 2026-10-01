FROM python:3.12-slim

WORKDIR /app

# System deps for psycopg2 / openpyxl
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpq5 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Ensure runtime dirs exist
RUN mkdir -p data static/uploads templates

ENV PYTHONUNBUFFERED=1
ENV PORT=8000

EXPOSE 8000

# Koyeb sets PORT; default 8000
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
