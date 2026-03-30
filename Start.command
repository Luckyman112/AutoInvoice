#!/bin/bash

# Переходим в папку со скриптом
cd "$(dirname "$0")"

echo "=================================================="
echo "Проверка обновлений Auto-Invoice..."
echo "=================================================="

# есть ли скрытая папка .git
if [ -d ".git" ]; then
    # Если да — подтягиваем последние изменения
    git pull origin main
else
    # принудительно превращаем папку в Git-репозиторий и качаем свежий код
    git init
    git remote add origin https://github.com/Luckyman112/AutoInvoice.git
    git fetch --all
    git reset --hard origin/main
fi

echo "=================================================="
echo "Запуск программы..."
echo "=================================================="

# Проверяем и запускаем виртуальное окружение
if [ ! -d "venv" ]; then
    echo "Первый запуск: Устанавливаем библиотеки (это займет около минуты)..."
    python3 -m venv venv
    source venv/bin/activate
    pip install customtkinter openpyxl
else
    source venv/bin/activate
fi

# Запускаем питон!
python3 main.py