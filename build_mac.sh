#!/usr/bin/env bash
#
# Сборка приложения «K2 Mailer.app» для macOS.
# Запускать НА МАШИНЕ С macOS (PyInstaller собирает под ту ОС, на которой запущен).
#
# Требуется установленный Python 3.10+ (python3 --version).
#
# Использование:
#   chmod +x build_mac.sh
#   ./build_mac.sh
#
# Результат: dist/K2 Mailer.app  — раздавайте сотрудникам (можно заархивировать).

set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="K2 Mailer"
BUNDLE_ID="ru.k2sdr.mailer"

echo "==> Создаю изолированное окружение для сборки (.venv-build)"
python3 -m venv .venv-build
# shellcheck disable=SC1091
source .venv-build/bin/activate

echo "==> Ставлю зависимости и PyInstaller"
pip install --upgrade pip
pip install -r requirements.txt pyinstaller

echo "==> Собираю .app"
pyinstaller --noconfirm --clean --windowed \
  --name "$APP_NAME" \
  --osx-bundle-identifier "$BUNDLE_ID" \
  --collect-all exchangelib \
  --copy-metadata exchangelib \
  --collect-submodules openpyxl \
  --collect-submodules flask \
  web_app.py

echo
echo "==> Готово: dist/$APP_NAME.app"
echo
echo "Чтобы у сотрудников не блокировался Gatekeeper (приложение не подписано),"
echo "после копирования к ним выполните один раз в терминале:"
echo "    xattr -dr com.apple.quarantine \"/путь/к/$APP_NAME.app\""
echo "либо первый запуск: правый клик по приложению -> «Открыть» -> «Открыть»."
