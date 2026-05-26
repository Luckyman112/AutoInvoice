@echo off
:: Включаем поддержку русского языка в консоли
chcp 65001 >nul

:: Переходим в папку, где лежит этот bat-файл (магия Windows)
cd /d "%~dp0"

echo ==================================================
echo Проверка обновлений Auto-Invoice...
echo ==================================================

:: Проверяем, настроен ли здесь Git (есть ли скрытая папка .git)
if exist .git\ (
    git pull origin main
) else (
    git init
    git remote add origin https://github.com/Luckyman112/AutoInvoice.git
    git fetch --all
    git reset --hard origin/main
)

echo ==================================================
echo Запуск программы...
echo ==================================================

:: Проверяем наличие виртуального окружения (папки venv)
if not exist venv\ (
    echo Первый запуск: Устанавливаем библиотеки (это займет около минуты)...
    python -m venv venv
    call venv\Scripts\activate
    pip install customtkinter openpyxl
) else (
    call venv\Scripts\activate
)

start venv\Scripts\pythonw.exe main.py