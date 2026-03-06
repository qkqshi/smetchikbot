import os
import sys

# Устанавливаем кодировку UTF-8 ДО импорта других библиотек
os.environ['PYTHONIOENCODING'] = 'utf-8'
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import json
import logging
from openai import OpenAI
from dotenv import load_dotenv
from parser import extract_furniture_data
from calculator import Calculator

# Отключаем логирование httpx для избежания проблем с кодировкой
logging.getLogger("httpx").setLevel(logging.WARNING)

# Загружаем переменные окружения (.env)
# Указываем путь к .env относительно текущего файла
from pathlib import Path
env_path = Path(__file__).parent.parent / '.env'
load_dotenv(dotenv_path=env_path)

# Настройка клиента OpenRouter
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("OPENROUTER_API_KEY"),
)

def get_facade_features(description):
    """
    Отправляет текстовое описание фасада ИИ-агенту для определения типа и наценок.
    """
    print(f"🔍 Анализирую описание фасада: '{description}'")
    
    prompt = f"""
    Ты — технический аналитик Ergonomika. Твоя задача — классифицировать фасады.
    
    Извлеки из текста менеджера следующие параметры для JSON:
    
    1. material:
       - "egger" (если ЛДСП)
       - "enamel" (если эмаль/покраска)
       - "veneer" (если шпон)
    
    2. two_sided: true (если есть "2х стор", "двухсторонняя", "2 стороны")
    3. milling: true (если есть "фреза", "фрезеровка", "классика")
    
    Текст: "{description}"
    
    Ответь только чистым JSON без пояснений.
    Пример: {{"material": "enamel", "two_sided": true, "milling": false}}
    """
    
    # Проверяем наличие API ключа
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key or api_key == "твой_ключ_здесь":
        print("⚠️ API ключ не настроен. Использую параметры по ключевым словам.")
        # Пытаемся определить параметры по ключевым словам
        desc_lower = description.lower()
        
        # Определяем материал
        material = "egger"
        if "эмаль" in desc_lower or "краск" in desc_lower or "покраск" in desc_lower:
            material = "enamel"
            print(f"   ✓ Найдена эмаль")
        elif "шпон" in desc_lower:
            material = "veneer"
            print(f"   ✓ Найден шпон")
        else:
            print(f"   ℹ️ Материал по умолчанию: egger (ЛДСП)")
        
        # Определяем двухсторонность
        two_sided = False
        if "2х" in description or "2 х" in description or "двух" in desc_lower or "2стор" in desc_lower or "2-стор" in desc_lower:
            two_sided = True
            print(f"   ✓ Найдена двухсторонность")
        
        # Определяем фрезеровку
        milling = False
        if "фрез" in desc_lower or "классик" in desc_lower or "рисун" in desc_lower:
            milling = True
            print(f"   ✓ Найдена фрезеровка")
        
        result = {"material": material, "two_sided": two_sided, "milling": milling}
        print(f"   📋 Результат: {result}")
        return result
    
    try:
        response = client.chat.completions.create(
            model=os.getenv("MODEL_NAME", "google/gemini-2.0-flash-001"),
            messages=[{"role": "user", "content": prompt}]
        )
        # Очищаем ответ от возможных Markdown-тегов и парсим в словарь
        content = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
        ai_result = json.loads(content)
        
        # Преобразуем структуру ответа ИИ в нужный формат
        result = {
            "material": ai_result.get("material_type", ai_result.get("material", "egger")),
            "two_sided": ai_result.get("surcharges", {}).get("two_sided", ai_result.get("two_sided", False)),
            "milling": ai_result.get("surcharges", {}).get("milling", ai_result.get("milling", False))
        }
        
        print(f"   📋 Результат от ИИ: {result}")
        return result
    except Exception as e:
        print(f"⚠️ Ошибка ИИ: {str(e)}. Использую параметры по ключевым словам.")
        # Пытаемся определить параметры по ключевым словам
        desc_lower = description.lower()
        
        # Определяем материал
        material = "egger"
        if "эмаль" in desc_lower or "краск" in desc_lower or "покраск" in desc_lower:
            material = "enamel"
            print(f"   ✓ Найдена эмаль")
        elif "шпон" in desc_lower:
            material = "veneer"
            print(f"   ✓ Найден шпон")
        else:
            print(f"   ℹ️ Материал по умолчанию: egger (ЛДСП)")
        
        # Определяем двухсторонность
        two_sided = False
        if "2х" in description or "2 х" in description or "двух" in desc_lower or "2стор" in desc_lower or "2-стор" in desc_lower:
            two_sided = True
            print(f"   ✓ Найдена двухсторонность")
        
        # Определяем фрезеровку
        milling = False
        if "фрез" in desc_lower or "классик" in desc_lower or "рисун" in desc_lower:
            milling = True
            print(f"   ✓ Найдена фрезеровка")
        
        result = {"material": material, "two_sided": two_sided, "milling": milling}
        print(f"   📋 Результат fallback: {result}")
        return result

def run_agent():
    # 1. Инициализация
    calc = Calculator('data/prices.json')
    items = extract_furniture_data('uploads/пример кухни.docx')
    
    if not items:
        print("❌ Нет данных для расчета.")
        return

    print("\n" + "="*40)
    print("🚀 ЗАПУСК ИИ-РАСЧЕТА КП")
    print("="*40)
    
    for item in items:
        print(f"\n🔍 Анализирую изделие: {item['name']}...")
        
        # 2. ИИ определяет параметры фасада из текста менеджера
        features = get_facade_features(item['facade'])
        
        # 3. Считаем стоимость через калькулятор
        result = calc.calculate_cost(
            width=item['width'],
            height=item['height'],
            depth=item['depth'],
            body_type="egger",
            facade_type=features['material'],
            is_two_sided=features['two_sided'],
            has_milling=features['milling']
        )

        # 4. Вывод результата в консоль
        print(f"✅ Результат для '{item['name']}':")
        print(f"   📐 Габариты: {item['width']}x{item['height']}x{item['depth']} мм (S={result['area']} м2)")
        print(f"   🎭 Фасад: {features['material'].upper()} (2-стор: {features['two_sided']}, Фреза: {features['milling']})")
        print(f"   💰 Корпус: {result['body_sum']} руб.")
        print(f"   💰 Фасад: {result['facade_sum']} руб.")
        print(f"   ИТОГО: {result['total']} руб.")
        print("-" * 40)

if __name__ == "__main__":
    run_agent()
