@echo off
REM Сборка K2 Mailer.exe для Windows. Запускать НА WINDOWS.
REM Требуется установленный Python 3.10+ (python --version).
REM Результат: dist\K2 Mailer\K2 Mailer.exe  — раздавайте сотрудникам (папку целиком).

setlocal
cd /d "%~dp0"

echo ==> Создаю окружение для сборки
python -m venv .venv-build
call .venv-build\Scripts\activate.bat

echo ==> Ставлю зависимости и PyInstaller
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller

echo ==> Собираю приложение
pyinstaller --noconfirm --clean --windowed ^
  --name "K2 Mailer" ^
  --collect-all exchangelib ^
  --copy-metadata exchangelib ^
  --collect-submodules openpyxl ^
  --collect-submodules flask ^
  web_app.py

echo.
echo ==> Готово: dist\K2 Mailer\K2 Mailer.exe
endlocal
