FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DB_PATH=/data/tempform.db \
    PORT=5205

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /data
VOLUME ["/data"]

EXPOSE 5205

CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5205", "--timeout", "120", "--access-logfile", "-", "app:app"]
