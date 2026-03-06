import json

class Calculator:
    def __init__(self, price_file='data/prices.json'):
        with open(price_file, 'r', encoding='utf-8') as f:
            self.prices = json.load(f)

    def calculate_cost(self, width, height, depth, body_type, facade_type, is_two_sided=False, has_milling=False):
        # 1. Общая площадь по передней проекции в м2
        area = (width / 1000) * (height / 1000)
        
        # 2. Расчет КОРПУСА
        # Базовая цена за м2 из прайса (например, 24 000 для Egger)
        p_body = self.prices['body_materials'].get(body_type, 0)
        # Коррекция по глубине: (Ш * В) * Цена * (Глубина / 600)
        cost_body = area * p_body * (depth / 600)
        
        # 3. Расчет ФАСАДА
        # Базовая цена за м2 (Egger: 6500, Эмаль: 17000, Шпон: 30000)
        p_facade = self.prices['facade_materials'].get(facade_type, 0)
        
        # Логика наценок
        if facade_type == "egger":
            # Для Egger двухсторонняя отделка просто удваивает базу
            if is_two_sided:
                p_facade *= 2
        else:
            # Для Эмали и Шпона прибавляем фиксированные суммы за м2
            if is_two_sided:
                p_facade += self.prices['surcharges'].get('two_sided', 0) # +12000
            if has_milling:
                p_facade += self.prices['surcharges'].get('milling', 0)   # +8000
        
        # Итоговая стоимость фасада
        cost_facade = area * p_facade
        
        return {
            "area": round(area, 2),
            "body_sum": round(cost_body, 2),
            "facade_sum": round(cost_facade, 2),
            "total": round(cost_body + cost_facade, 2)
        }