#!/usr/bin/env bash
#
# Установщик K2 Mailer в серверном режиме на Ubuntu (для VM в K2 Cloud).
# Делает всё за один проход: пакеты, пользователь, venv, зависимости, служба
# systemd с автозапуском. После него останется только настроить HTTPS (nginx).
#
# Запуск (из папки с кодом):
#   sudo K2_EWS_URL="https://mail.k2.cloud/EWS/Exchange.asmx" bash deploy/install.sh
#
# Переменные (необязательные):
#   K2_EWS_URL    адрес EWS вашего Exchange (можно задать позже в юните)
#   K2_AUTH_TYPE  NTLM (по умолчанию) | basic | auto
#   K2_VERIFY_SSL true (по умолчанию) | false
#   K2_PORT       внутренний порт приложения (по умолчанию 8765)

set -euo pipefail

if [[ $EUID -ne 0 ]]; then
  echo "Запустите через sudo: sudo bash deploy/install.sh" >&2
  exit 1
fi

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
DATA_DIR="/var/lib/k2-mailer"
SVC_USER="k2mailer"
PORT="${K2_PORT:-8765}"
EWS_URL="${K2_EWS_URL:-https://mail.ЗАМЕНИТЕ.ru/EWS/Exchange.asmx}"
AUTH_TYPE="${K2_AUTH_TYPE:-NTLM}"
VERIFY_SSL="${K2_VERIFY_SSL:-true}"

echo "==> Код приложения: $APP_DIR"

echo "==> Устанавливаю системные пакеты"
apt-get update -y
apt-get install -y python3-venv python3-pip

echo "==> Создаю служебного пользователя $SVC_USER и папку данных $DATA_DIR"
id -u "$SVC_USER" >/dev/null 2>&1 || useradd -r -m -s /usr/sbin/nologin "$SVC_USER"
mkdir -p "$DATA_DIR"
chown -R "$SVC_USER:$SVC_USER" "$DATA_DIR"

echo "==> Создаю виртуальное окружение и ставлю зависимости"
python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install --upgrade pip >/dev/null
"$APP_DIR/.venv/bin/pip" install -r "$APP_DIR/requirements.txt"
chown -R "$SVC_USER:$SVC_USER" "$APP_DIR"

SECRET="$(python3 -c 'import secrets; print(secrets.token_hex(32))')"

echo "==> Пишу службу /etc/systemd/system/k2-mailer.service"
cat > /etc/systemd/system/k2-mailer.service <<EOF
[Unit]
Description=K2 Mailer (рассылка из OWA, веб-интерфейс)
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=$SVC_USER
Group=$SVC_USER
WorkingDirectory=$APP_DIR
Environment=K2_MAILER_SERVER=1
Environment=K2_MAILER_HOST=127.0.0.1
Environment=K2_MAILER_PORT=$PORT
Environment=K2_MAILER_SECRET=$SECRET
Environment=K2_MAILER_HOME=$DATA_DIR
Environment=K2_EWS_URL=$EWS_URL
Environment=K2_AUTH_TYPE=$AUTH_TYPE
Environment=K2_VERIFY_SSL=$VERIFY_SSL
ExecStart=$APP_DIR/.venv/bin/python $APP_DIR/web_app.py
Restart=on-failure
RestartSec=5
NoNewPrivileges=true
PrivateTmp=true
ReadWritePaths=$DATA_DIR

[Install]
WantedBy=multi-user.target
EOF

echo "==> Запускаю службу"
systemctl daemon-reload
systemctl enable --now k2-mailer
sleep 2
systemctl --no-pager --full status k2-mailer | head -n 12 || true

echo
echo "============================================================"
echo " Готово. Приложение слушает 127.0.0.1:$PORT (внутри машины)."
if [[ "$EWS_URL" == *ЗАМЕНИТЕ* ]]; then
  echo " ⚠ Не задан K2_EWS_URL. Откройте юнит и впишите адрес EWS:"
  echo "     sudo nano /etc/systemd/system/k2-mailer.service"
  echo "     sudo systemctl daemon-reload && sudo systemctl restart k2-mailer"
fi
echo " Следующий шаг — открыть наружу по HTTPS (nginx):"
echo "   sudo apt-get install -y nginx"
echo "   sudo cp deploy/nginx-k2-mailer.conf /etc/nginx/sites-available/k2-mailer"
echo "   # впишите server_name и сертификаты, затем:"
echo "   sudo ln -s /etc/nginx/sites-available/k2-mailer /etc/nginx/sites-enabled/"
echo "   sudo nginx -t && sudo systemctl reload nginx"
echo "============================================================"
