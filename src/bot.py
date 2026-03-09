import os
import sys
import json
import logging
import re
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from dotenv import load_dotenv
from parser import extract_furniture_data
from calculator import Calculator
from main import get_facade_features
from pptx_generator import generate_kp_pptx

# Загружаем переменные окружения
load_dotenv()

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Отключаем логирование httpx
logging.getLogger("httpx").setLevel(logging.WARNING)

# Инициализация калькулятора
calc = Calculator('data/prices.json')

# Временное хранилище для фото (user_id -> photo_path)
user_photos = {}

# Временное хранилище последних расчетов (user_id -> calculation_data)
user_calculations = {}

# Временное хранилище номеров телефонов (user_id -> phone_number)
user_phone_numbers = {}

def clean_number(text):
    """Извлекает первое целое число из строки и возвращает его как int"""
    if not text:
        return 0
    if isinstance(text, (int, float)):
        return int(text)
    # Находим все числа в строке
    numbers = re.findall(r'\d+', str(text).replace(' ', '').replace(',', ''))
    if numbers:
        return int(numbers[0])
    return 0

async def send_long_message(update: Update, text: str, max_length: int = 4000):
    """Отправляет длинное сообщение, разбивая его на части если нужно"""
    if len(text) <= max_length:
        await update.message.reply_text(text)
        return
    
    # Разбиваем по строкам
    lines = text.split('\n')
    current_message = ""
    
    for line in lines:
        # Если добавление строки превысит лимит, отправляем текущее сообщение
        if len(current_message) + len(line) + 1 > max_length:
            if current_message:
                await update.message.reply_text(current_message)
                current_message = ""
        
        current_message += line + "\n"
    
    # Отправляем остаток
    if current_message:
        await update.message.reply_text(current_message)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start"""
    await update.message.reply_text(
        "👋 Привет! Я бот для расчета стоимости мебели.\n\n"
        "Как работать:\n"
        "1. Отправь фото проекта (опционально)\n"
        "2. Отправь .docx файл с таблицей изделий\n"
        "3. Получи КП в виде текста\n\n"
        "Дополнительные команды:\n"
        "• 'pdf' - получить PDF файл\n"
        "• 'excel' или 'таблица' - получить Excel файл\n"
        "• 'pptx' или 'презентация' - получить презентацию PowerPoint\n"
        "• 'добавь 15%' - добавить наценку ко всем позициям\n"
        "• 'сбросить' - вернуться к исходным ценам\n"
        "• 'добавь 10% к позиции 1,3' - к конкретным позициям\n"
        "• 'отправь pdf' или 'пдф' - получить КП в PDF\n"
        "• 'отправь excel' или 'эксель' - получить таблицу расчетов\n"
        "• 'номер 89151386664' - изменить номер на обложке"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик фотографий"""
    user_id = update.message.from_user.id
    photo = update.message.photo[-1]  # Берем фото наибольшего размера
    
    try:
        # Скачиваем фото
        file = await context.bot.get_file(photo.file_id)
        photo_path = f"uploads/photo_{user_id}.jpg"
        await file.download_to_drive(photo_path)
        
        # Сохраняем путь к фото для пользователя
        user_photos[user_id] = photo_path
        
        await update.message.reply_text(
            "✅ Фото получено!\n"
            "Теперь отправь .docx файл с таблицей изделий."
        )
    except Exception as e:
        logger.error(f"Ошибка обработки фото: {e}")
        await update.message.reply_text("⚠️ Ошибка при обработке фото")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик текстовых сообщений для добавления процентов, PDF и номера телефона"""
    user_id = update.message.from_user.id
    text = update.message.text.lower()
    
    # Проверка на запрос PDF
    if 'pdf' in text or 'пдф' in text:
        if user_id not in user_calculations:
            await update.message.reply_text(
                "⚠️ Сначала отправьте .docx файл для расчета.\n"
                "После получения расчета вы сможете запросить PDF."
            )
            return
        
        await update.message.reply_text("📄 Генерирую PDF...")
        
        try:
            from pdf_generator import generate_kp_pdf
            
            calc_data = user_calculations[user_id]
            project_name = calc_data['project_name']
            photo_path = calc_data.get('photo_path')
            phone_number = user_phone_numbers.get(user_id)
            
            # Формируем данные для PDF
            pdf_items = []
            for item_data in calc_data['items']:
                pdf_items.append({
                    'name': item_data['item']['name'],
                    'body_material': 'ЛДСП',
                    'facade_description': item_data['item']['facade'],
                    'additional_info': item_data.get('additional_info', []),
                    'total_cost': item_data['result']['total'],
                    'image': item_data['item'].get('image')  # Добавляем изображение
                })
            
            furniture_items = calc_data.get('furniture_items', [])
            
            pdf_path = generate_kp_pdf(project_name, pdf_items, furniture_items, photo_path, phone_number)
            
            with open(pdf_path, 'rb') as pdf_file:
                await update.message.reply_document(
                    document=pdf_file,
                    filename=f"{project_name}_КП.pdf",
                    caption="📄 Коммерческое предложение (PDF)"
                )
        except Exception as e:
            logger.error(f"Ошибка генерации PDF: {e}", exc_info=True)
            error_msg = f"⚠️ Не удалось создать PDF файл: {str(e)}"
            await send_long_message(update, error_msg)
        return
    
    # Проверка на запрос Excel
    if 'excel' in text or 'эксель' in text or 'xlsx' in text or 'таблиц' in text:
        if user_id not in user_calculations:
            await update.message.reply_text(
                "⚠️ Сначала отправьте .docx файл для расчета.\n"
                "После получения расчета вы сможете запросить Excel."
            )
            return
        
        await update.message.reply_text("📊 Генерирую Excel...")
        
        try:
            from excel_generator import generate_kp_excel
            
            calc_data = user_calculations[user_id]
            project_name = calc_data['project_name']
            
            # Используем данные из расчета
            items_data = calc_data['items']
            furniture_items = calc_data.get('furniture_items', [])
            
            excel_path = generate_kp_excel(project_name, items_data, furniture_items)
            
            with open(excel_path, 'rb') as excel_file:
                await update.message.reply_document(
                    document=excel_file,
                    filename=f"{project_name}_Расчет.xlsx",
                    caption="📊 Таблица расчетов (Excel)"
                )
        except Exception as e:
            logger.error(f"Ошибка генерации Excel: {e}", exc_info=True)
            error_msg = f"⚠️ Не удалось создать Excel файл: {str(e)}"
            await send_long_message(update, error_msg)
        return
    
    # Проверка на запрос презентации PPTX
    if 'pptx' in text or 'презентац' in text or 'powerpoint' in text or 'слайд' in text:
        if user_id not in user_calculations:
            await update.message.reply_text(
                "⚠️ Сначала отправьте .docx файл для расчета.\n"
                "После получения расчета вы сможете запросить презентацию."
            )
            return
        
        await update.message.reply_text("📊 Генерирую презентацию...")
        
        try:
            calc_data = user_calculations[user_id]
            project_name = calc_data['project_name']
            photo_path = calc_data.get('photo_path')
            phone_number = user_phone_numbers.get(user_id)
            
            # Формируем данные для презентации
            pptx_items = []
            for item_data in calc_data['items']:
                pptx_items.append({
                    'name': item_data['item']['name'],
                    'body_material': 'ЛДСП',
                    'facade_description': item_data['item']['facade'],
                    'additional_info': item_data.get('additional_info', []),
                    'total_cost': item_data['result']['total'],
                    'image': item_data['item'].get('image')  # Добавляем изображение
                })
            
            furniture_items = calc_data.get('furniture_items', [])
            
            pptx_path = generate_kp_pptx(project_name, pptx_items, furniture_items, photo_path, phone_number)
            
            with open(pptx_path, 'rb') as pptx_file:
                await update.message.reply_document(
                    document=pptx_file,
                    filename=f"{project_name}_КП.pptx",
                    caption="📊 Коммерческое предложение (презентация)"
                )
        except Exception as e:
            logger.error(f"Ошибка генерации PPTX: {e}", exc_info=True)
            error_msg = f"⚠️ Не удалось создать презентацию: {str(e)}"
            await send_long_message(update, error_msg)
        return
    
    # Проверка на команду "сбросить"
    if 'сброс' in text or 'reset' in text or 'вернуть' in text:
        if user_id not in user_calculations:
            await update.message.reply_text("⚠️ Нет сохраненных расчетов для сброса.")
            return
        
        calc_data = user_calculations[user_id]
        base_items = calc_data.get('base_items')
        
        if not base_items:
            await update.message.reply_text("⚠️ Базовые данные не найдены.")
            return
        
        await update.message.reply_text("🔄 Сбрасываю все наценки...")
        
        # Восстанавливаем базовые данные
        response = "📊 СБРОС К ИСХОДНЫМ ЦЕНАМ\n"
        response += "=" * 40 + "\n\n"
        
        total_sum = 0
        pdf_items = []
        
        for item_data in base_items:
            quantity = item_data.get('quantity', 1)
            original_total = item_data['result']['total']
            total_with_quantity = original_total * quantity
            
            response += f"🔹 {item_data['item']['name']}\n"
            response += f"   Размеры: {item_data['item']['width']}x{item_data['item']['height']}x{item_data['item']['depth']} мм\n"
            response += f"   Площадь: {item_data['result']['area']} м²\n"
            response += f"   Количество: {quantity} шт.\n"
            response += f"   ИТОГО ({quantity} шт.): {total_with_quantity:,.0f} руб.\n\n"
            
            total_sum += total_with_quantity
            
            pdf_items.append({
                'name': item_data['item']['name'],
                'body_material': 'ЛДСП',
                'facade_description': item_data['item']['facade'],
                'additional_info': item_data.get('additional_info', []),
                'total_cost': total_with_quantity,
                'image': item_data['item'].get('image'),
                'quantity': quantity
            })
        
        response += "=" * 40 + "\n"
        response += f"💰 ОБЩАЯ СУММА МАТЕРИАЛОВ: {total_sum:,.0f} руб.\n\n"
        
        # Добавляем мебель
        furniture_items = calc_data.get('furniture_items', [])
        furniture_total = 0
        
        if furniture_items:
            response += "=" * 40 + "\n"
            response += "🪑 МЕБЕЛЬ\n"
            response += "=" * 40 + "\n\n"
            
            for furn in furniture_items:
                response += f"🔹 {furn['name']}\n"
                response += f"   Количество: {furn['quantity']} шт.\n"
                response += f"   Цена за 1 шт: {furn['price_per_unit']} руб.\n"
                
                total_price = furn['total_price']
                if isinstance(total_price, str):
                    total_price = int(total_price.replace(' ', '').replace(',', ''))
                
                response += f"   ИТОГО: {total_price:,} руб.\n\n".replace(',', ' ')
                furniture_total += total_price
            
            response += "=" * 40 + "\n"
            response += f"💰 ОБЩАЯ СУММА МЕБЕЛИ: {furniture_total:,} руб.\n\n".replace(',', ' ')
            response += "=" * 40 + "\n"
            response += f"💵 ИТОГО ПО ПРОЕКТУ: {(total_sum + furniture_total):,} руб.".replace(',', ' ')
        
        await send_long_message(update, response)
        
        # Обновляем сохраненные данные - возвращаем к базовым
        base_project_name = calc_data.get('base_project_name', calc_data['project_name'])
        
        user_calculations[user_id] = {
            'items': base_items.copy(),
            'base_items': base_items,
            'furniture_items': furniture_items,
            'project_name': base_project_name,
            'base_project_name': base_project_name,
            'photo_path': calc_data.get('photo_path')
        }
        
        # Генерируем новую презентацию
        try:
            project_name = base_project_name
            photo_path = calc_data.get('photo_path')
            phone_number = user_phone_numbers.get(user_id)
            
            pptx_path = generate_kp_pptx(project_name, pdf_items, furniture_items, photo_path, phone_number)
            
            with open(pptx_path, 'rb') as pptx_file:
                await update.message.reply_document(
                    document=pptx_file,
                    filename=f"{project_name}_КП.pptx",
                    caption="📊 КП с исходными ценами"
                )
        except Exception as e:
            logger.error(f"Ошибка генерации PPTX: {e}", exc_info=True)
            error_msg = f"⚠️ Не удалось создать PPTX файл: {str(e)}"
            await send_long_message(update, error_msg)
        
        return
    
    # Проверка на номер телефона
    phone_match = re.search(r'(?:номер|телефон|тел\.?|phone)?\s*[:\-]?\s*(\+?[78]?\s*\(?\d{3}\)?\s*\d{3}[\s\-]?\d{2}[\s\-]?\d{2})', text)
    if phone_match:
        phone_number = phone_match.group(1).strip()
        # Очищаем номер от лишних символов
        phone_clean = re.sub(r'[^\d+]', '', phone_number)
        
        user_phone_numbers[user_id] = phone_clean
        
        await update.message.reply_text(
            f"✅ Номер телефона сохранен: {phone_number}\n"
            "Он будет использован на обложке следующих презентаций."
        )
        return
    
    # Проверяем, есть ли сохраненный расчет для добавления процентов
    if user_id not in user_calculations:
        await update.message.reply_text(
            "⚠️ Сначала отправьте .docx файл для расчета.\n"
            "После получения расчета вы сможете:\n"
            "• Добавить наценку: 'добавь 15%'\n"
            "• Сбросить наценки: 'сбросить'\n"
            "• Запросить PDF: 'отправь pdf'\n"
            "• Указать номер: 'номер 89151386664'"
        )
        return
    
    # Парсим команду "добавь N%"
    
    # Ищем процент
    percent_match = re.search(r'(\d+)\s*%', text)
    if not percent_match:
        await update.message.reply_text(
            "⚠️ Не могу понять команду. Используйте формат:\n"
            "• 'добавь 15%' - ко всем позициям\n"
            "• 'добавь 10% к позиции 1,3' - к конкретным позициям\n"
            "• 'отправь pdf' - получить PDF\n"
            "• 'номер 89151386664' - изменить номер на обложке"
        )
        return
    
    percent = int(percent_match.group(1))
    
    # Ищем номера позиций
    position_match = re.search(r'позиции?\s+([\d,\s]+)', text)
    target_positions = None
    
    if position_match:
        # Извлекаем номера позиций
        positions_str = position_match.group(1)
        target_positions = [int(p.strip()) for p in re.findall(r'\d+', positions_str)]
    
    # Получаем сохраненные данные
    calc_data = user_calculations[user_id]
    
    await update.message.reply_text(f"⏳ Пересчитываю с наценкой {percent}%...")
    
    # Пересчитываем от ТЕКУЩИХ данных (с учетом предыдущих наценок)
    current_items = calc_data['items']
    
    # Пересчитываем
    response = f"📊 ПЕРЕРАСЧЕТ С НАЦЕНКОЙ {percent}%\n"
    response += "=" * 40 + "\n\n"
    
    total_sum = 0
    pdf_items = []
    
    for idx, item_data in enumerate(current_items, start=1):
        # Применяем наценку если нужно
        apply_markup = target_positions is None or idx in target_positions
        markup_multiplier = (1 + percent / 100) if apply_markup else 1
        
        # Получаем количество
        quantity = item_data.get('quantity', 1)
        
        # Берем ТЕКУЩУЮ стоимость (уже с предыдущими наценками)
        current_total = item_data['result']['total']
        new_total = current_total * markup_multiplier
        
        # Умножаем на количество
        total_with_quantity = new_total * quantity
        
        response += f"🔹 {item_data['item']['name']}"
        if apply_markup:
            response += f" (+{percent}%)"
        response += "\n"
        response += f"   Размеры: {item_data['item']['width']}x{item_data['item']['height']}x{item_data['item']['depth']} мм\n"
        response += f"   Площадь: {item_data['result']['area']} м²\n"
        response += f"   Количество: {quantity} шт.\n"
        
        if apply_markup:
            response += f"   Было: {current_total:,.0f} руб.\n"
            response += f"   С наценкой: {new_total:,.0f} руб.\n"
            response += f"   ИТОГО ({quantity} шт.): {total_with_quantity:,.0f} руб.\n\n"
        else:
            response += f"   ИТОГО ({quantity} шт.): {total_with_quantity:,.0f} руб.\n\n"
        
        total_sum += total_with_quantity
        
        pdf_items.append({
            'name': item_data['item']['name'],
            'body_material': 'ЛДСП',
            'facade_description': item_data['item']['facade'],
            'additional_info': item_data.get('additional_info', []),
            'total_cost': total_with_quantity,
            'image': item_data['item'].get('image'),  # Добавляем изображение
            'quantity': quantity
        })
    
    response += "=" * 40 + "\n"
    response += f"💰 ОБЩАЯ СУММА МАТЕРИАЛОВ: {total_sum:,.0f} руб.\n\n"
    
    # Добавляем мебель (без изменений)
    furniture_items = calc_data.get('furniture_items', [])
    furniture_total = 0
    
    if furniture_items:
        response += "=" * 40 + "\n"
        response += "🪑 МЕБЕЛЬ\n"
        response += "=" * 40 + "\n\n"
        
        for furn in furniture_items:
            response += f"🔹 {furn['name']}\n"
            response += f"   Количество: {furn['quantity']} шт.\n"
            response += f"   Цена за 1 шт: {furn['price_per_unit']} руб.\n"
            
            total_price = furn['total_price']
            if isinstance(total_price, str):
                total_price = int(total_price.replace(' ', '').replace(',', ''))
            
            response += f"   ИТОГО: {total_price:,} руб.\n\n".replace(',', ' ')
            furniture_total += total_price
        
        response += "=" * 40 + "\n"
        response += f"💰 ОБЩАЯ СУММА МЕБЕЛИ: {furniture_total:,} руб.\n\n".replace(',', ' ')
        response += "=" * 40 + "\n"
        response += f"💵 ИТОГО ПО ПРОЕКТУ: {(total_sum + furniture_total):,} руб.".replace(',', ' ')
    
    await send_long_message(update, response)
    
    # Обновляем сохраненные данные с учетом наценки
    updated_items = []
    for idx, item_data in enumerate(current_items, start=1):
        apply_markup = target_positions is None or idx in target_positions
        markup_multiplier = (1 + percent / 100) if apply_markup else 1
        
        quantity = item_data.get('quantity', 1)
        
        # Создаем копию данных с обновленной стоимостью
        updated_result = item_data['result'].copy()
        updated_result['total'] = item_data['result']['total'] * markup_multiplier
        
        updated_items.append({
            'item': item_data['item'],
            'features': item_data['features'],
            'result': updated_result,
            'additional_info': item_data.get('additional_info', []),
            'quantity': quantity
        })
    
    # Обновляем сохраненные данные (текущие становятся новыми текущими)
    base_project_name = calc_data.get('base_project_name', calc_data['project_name'])
    
    user_calculations[user_id] = {
        'items': updated_items,
        'furniture_items': calc_data.get('furniture_items', []),
        'project_name': base_project_name,  # Убрали добавление процента
        'base_project_name': base_project_name,  # Сохраняем базовое название
        'photo_path': calc_data.get('photo_path')
    }
    
    # Генерируем новый PPTX
    try:
        project_name = base_project_name  # Используем базовое название без процента
        photo_path = calc_data.get('photo_path')
        phone_number = user_phone_numbers.get(user_id)
        
        pptx_path = generate_kp_pptx(project_name, pdf_items, furniture_items, photo_path, phone_number)
        
        with open(pptx_path, 'rb') as pptx_file:
            await update.message.reply_document(
                document=pptx_file,
                filename=f"{project_name}_КП.pptx",
                caption=f"📊 КП с наценкой {percent}%"
            )
    except Exception as e:
        logger.error(f"Ошибка генерации PPTX: {e}", exc_info=True)
        error_msg = f"⚠️ Не удалось создать PPTX файл: {str(e)}"
        await send_long_message(update, error_msg)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик документов"""
    user_id = update.message.from_user.id
    document = update.message.document
    
    # Проверяем формат файла
    if not document.file_name.endswith('.docx'):
        await update.message.reply_text("⚠️ Пожалуйста, отправьте файл в формате .docx")
        return
    
    await update.message.reply_text("📥 Получил файл, начинаю обработку...")
    
    try:
        # Скачиваем файл
        file = await context.bot.get_file(document.file_id)
        file_path = f"uploads/{document.file_name}"
        await file.download_to_drive(file_path)
        
        # Извлекаем данные
        items = extract_furniture_data(file_path)
        
        if isinstance(items, tuple):
            items, furniture_items = items
        else:
            furniture_items = []
        
        if not items and not furniture_items:
            await update.message.reply_text("❌ Не удалось извлечь данные из файла.")
            return
        
        # Обрабатываем каждое изделие
        response = "📊 РАСЧЕТ КОММЕРЧЕСКОГО ПРЕДЛОЖЕНИЯ\n"
        response += "=" * 40 + "\n\n"
        
        total_sum = 0
        pdf_items = []
        project_image = None  # Изображение для PDF
        
        # Сохраняем данные для возможности пересчета
        calculation_items = []
        
        for item in items:
            # Сохраняем первое найденное изображение для обложки PDF
            if not project_image and item.get('image'):
                project_image = item['image']
            # Определяем параметры фасада через ИИ
            try:
                features = get_facade_features(item['facade'])
                logger.info(f"Получены параметры фасада: {features}")
            except Exception as e:
                logger.error(f"Ошибка определения параметров фасада: {e}")
                # Используем параметры по умолчанию
                features = {"material": "egger", "two_sided": False, "milling": False}
            
            # Считаем стоимость
            try:
                # Безопасно извлекаем размеры
                w = clean_number(item['width'])
                h = clean_number(item['height'])
                d = clean_number(item['depth'])
                
                result = calc.calculate_cost(
                    width=w,
                    height=h,
                    depth=d,
                    body_type="egger",
                    facade_type=features.get('material', 'egger'),
                    is_two_sided=features.get('two_sided', False),
                    has_milling=features.get('milling', False)
                )
            except Exception as e:
                logger.error(f"Ошибка расчета изделия {item['name']}: {e}")
                # Если расчет не удался, пропускаем или используем нули
                continue
            
            # Переводим название материала на русский
            material_names = {
                'egger': 'ЛДСП',
                'enamel': 'Эмаль',
                'veneer': 'Шпон'
            }
            material_ru = material_names.get(features.get('material', 'egger'), 'ЛДСП')
            
            # Получаем количество
            quantity = item.get('quantity', 1)
            if isinstance(quantity, str):
                quantity = int(quantity) if quantity.isdigit() else 1
            
            # Формируем ответ
            response += f"🔹 {item['name']}\n"
            response += f"   Размеры: {item['width']}x{item['height']}x{item['depth']} мм\n"
            response += f"   Площадь: {result['area']} м²\n"
            response += f"   Количество: {quantity} шт.\n"
            response += f"   Корпус: {result['body_sum']:,.0f} руб.\n"
            response += f"   Фасад ({material_ru}): {result['facade_sum']:,.0f} руб.\n"
            
            if features.get('two_sided'):
                response += f"   ✓ Двухсторонняя отделка\n"
            if features.get('milling'):
                response += f"   ✓ Фрезеровка\n"
            
            # Умножаем на количество
            item_total = result['total'] * quantity
            response += f"   ИТОГО: {item_total:,.0f} руб.\n\n"
            
            total_sum += item_total
            
            # Данные для PDF и сохранения
            additional_info = []
            if features.get('two_sided'):
                additional_info.append("Двухсторонняя отделка")
            if features.get('milling'):
                additional_info.append("Фрезеровка")
            
            # Сохраняем данные для пересчета
            calculation_items.append({
                'item': item,
                'features': features,
                'result': result,
                'additional_info': additional_info,
                'quantity': quantity
            })
            
            pdf_items.append({
                'name': item['name'],
                'body_material': 'ЛДСП',
                'facade_description': item['facade'],
                'additional_info': additional_info,
                'total_cost': item_total,  # Уже умноженная на количество
                'image': item.get('image'),  # Добавляем изображение
                'quantity': quantity
            })
        
        response += "=" * 40 + "\n"
        response += f"💰 ОБЩАЯ СУММА МАТЕРИАЛОВ: {total_sum:,.0f} руб.\n\n"
        
        # Добавляем информацию о мебели
        furniture_total = 0
        if furniture_items:
            response += "=" * 40 + "\n"
            response += "🪑 МЕБЕЛЬ\n"
            response += "=" * 40 + "\n\n"
            
            for furn in furniture_items:
                response += f"🔹 {furn['name']}\n"
                response += f"   Количество: {furn['quantity']} шт.\n"
                response += f"   Цена за 1 шт: {furn['price_per_unit']} руб.\n"
                
                # Преобразуем итоговую стоимость в число (теперь это уже число из parser.py)
                total_price = furn['total_price']
                if isinstance(total_price, str):
                    total_price = clean_number(total_price)
                
                response += f"   ИТОГО: {total_price:,} руб.\n\n".replace(',', ' ')
                furniture_total += total_price
            
            response += "=" * 40 + "\n"
            response += f"💰 ОБЩАЯ СУММА МЕБЕЛИ: {furniture_total:,} руб.\n\n".replace(',', ' ')
            response += "=" * 40 + "\n"
            response += f"💵 ИТОГО ПО ПРОЕКТУ: {(total_sum + furniture_total):,} руб.".replace(',', ' ')
        
        # Отправляем текстовый ответ, разбивая на части если нужно
        await send_long_message(update, response)
        
        # Генерируем PPTX
        project_name = document.file_name.replace('.docx', '')
        
        # Используем изображение из docx, если нет фото от пользователя
        photo_path = user_photos.get(user_id) or project_image
        
        # Сохраняем расчет для возможности добавления процентов
        user_calculations[user_id] = {
            'items': calculation_items,
            'base_items': calculation_items.copy(),  # Сохраняем базовые данные для сброса
            'furniture_items': furniture_items,
            'project_name': project_name,
            'base_project_name': project_name,  # Базовое название без суффиксов
            'photo_path': photo_path
        }
        
        try:
            # Генерируем PPTX
            logger.info(f"Генерация PPTX для проекта: {project_name}")
            logger.info(f"Фото проекта: {photo_path}")
            logger.info(f"Количество материалов: {len(pdf_items)}")
            logger.info(f"Количество мебели: {len(furniture_items)}")
            
            phone_number = user_phone_numbers.get(user_id)
            
            pptx_path = generate_kp_pptx(project_name, pdf_items, furniture_items, photo_path, phone_number)
            
            # Отправляем PPTX
            with open(pptx_path, 'rb') as pptx_file:
                await update.message.reply_document(
                    document=pptx_file,
                    filename=f"{project_name}_КП.pptx",
                    caption="📊 Коммерческое предложение"
                )
            
            # Очищаем фото пользователя после использования
            if user_id in user_photos:
                del user_photos[user_id]
                
        except Exception as e:
            logger.error(f"Ошибка генерации PPTX: {e}", exc_info=True)
            error_msg = f"⚠️ Не удалось создать PPTX файл: {str(e)}"
            await send_long_message(update, error_msg)
        
    except Exception as e:
        logger.error(f"Ошибка обработки файла: {e}")
        error_message = f"❌ Ошибка при обработке файла: {str(e)}"
        await send_long_message(update, error_message)

def main():
    """Запуск бота"""
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    
    if not token:
        logger.error("❌ Не найден TELEGRAM_BOT_TOKEN в .env файле")
        return
    
    # Создаем приложение
    application = Application.builder().token(token).build()
    
    # Регистрируем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    # Запускаем бота
    logger.info("🤖 Бот запущен...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
