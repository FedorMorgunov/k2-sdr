# Контейнер для K2 Mailer в серверном (многопользовательском) режиме.
# Сборка:  docker build -t k2-mailer .
# Запуск:  см. README, раздел «Запуск в Docker».
FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV K2_MAILER_SERVER=1 \
    K2_MAILER_HOST=0.0.0.0 \
    K2_MAILER_PORT=8765 \
    K2_MAILER_HOME=/data

VOLUME /data
EXPOSE 8765
CMD ["python", "web_app.py"]
