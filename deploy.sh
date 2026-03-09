#!/bin/bash

# Скрипт для загрузки обновлений на сервер

# Настройки (замените на свои)
SERVER_USER="your_user"
SERVER_HOST="your_server_ip"
SERVER_PATH="/opt/smetchikbot"

echo "🚀 Загрузка файлов на сервер..."

# Загружаем измененные файлы
scp src/bot.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/
scp src/pptx_generator.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/
scp src/pdf_generator.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/
scp src/excel_generator.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/
scp src/parser.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/
scp src/calculator.py ${SERVER_USER}@${SERVER_HOST}:${SERVER_PATH}/src/

echo "✅ Файлы загружены"

# Перезапускаем бота на сервере
echo "🔄 Перезапуск бота..."
ssh ${SERVER_USER}@${SERVER_HOST} << 'EOF'
cd /opt/smetchikbot
sudo systemctl restart smetchikbot
echo "✅ Бот перезапущен"
sudo systemctl status smetchikbot --no-pager
EOF

echo "🎉 Деплой завершен!"
