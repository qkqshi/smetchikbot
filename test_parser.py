from src.parser import extract_furniture_data

items, furniture = extract_furniture_data('uploads/ВЯЗНИКОВСКАЯ просчет.docx')

print(f"\n=== РЕЗУЛЬТАТЫ ===")
print(f"Найдено элементов: {len(items)}")
print(f"Найдено мебели: {len(furniture)}")

for idx, item in enumerate(items, 1):
    print(f"\n{idx}. {item['name']}")
    print(f"   Размеры: {item['width']}x{item['height']}x{item['depth']}")
    print(f"   Количество: {item['quantity']}")
    print(f"   Корпус: {item['body']}")
    print(f"   Фасад: {item['facade']}")
    print(f"   Фото: {item['image']}")
