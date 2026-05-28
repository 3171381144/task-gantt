FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TASK_GANTT_HOST=0.0.0.0 \
    TASK_GANTT_PORT=8010 \
    TASK_GANTT_DATA_DIR=/app/data

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py ./app.py
COPY static ./static

RUN mkdir -p /app/data

VOLUME ["/app/data"]
EXPOSE 8010

CMD ["python", "app.py"]
