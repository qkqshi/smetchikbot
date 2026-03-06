#!/bin/bash
# Скрипт для загрузки обложки и финальных слайдов на сервер

SERVER="root@v3061083.hosted-by-vdsina.ru"
REMOTE_PATH="/opt/smetchikbot/templates"

echo "Создание папки templates на сервере..."
ssh $SERVER "mkdir -p $REMOTE_PATH"

echo "Загрузка всех изображений из templates/..."
scp templates/*.png templates/*.jpg templates/*.jpeg $SERVER:$REMOTE_PATH/ 2>/dev/null

echo "Проверка загрузки..."
ssh $SERVER "ls -lh $REMOTE_PATH/"

echo "Загрузка обновленного генератора..."
scp src/pptx_generator.py $SERVER:/opt/smetchikbot/src/

echo "Перезапуск бота..."
ssh $SERVER "systemctl restart smetchikbot.service"

echo "Готово! Проверьте логи:"
echo "ssh $SERVER 'journalctl -u smetchikbot.service -f'"
