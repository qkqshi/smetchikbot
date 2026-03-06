import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image

class PDFGenerator:
    def __init__(self):
        self.width, self.height = A4
        self.bg_color = HexColor('#FFFFFF')
        self.text_color = HexColor('#000000')
        
        # Регистрируем шрифты с поддержкой кириллицы
        try:
            # Пытаемся использовать системные шрифты
            import platform
            if platform.system() == 'Windows':
                pdfmetrics.registerFont(TTFont('Arial', 'C:/Windows/Fonts/arial.ttf'))
                pdfmetrics.registerFont(TTFont('Arial-Bold', 'C:/Windows/Fonts/arialbd.ttf'))
            else:
                # Linux
                pdfmetrics.registerFont(TTFont('Arial', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
                pdfmetrics.registerFont(TTFont('Arial-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
            self.font_regular = 'Arial'
            self.font_bold = 'Arial-Bold'
            print("✅ Шрифты с кириллицей загружены успешно")
        except Exception as e:
            print(f"⚠️ Ошибка загрузки шрифтов: {e}")
            # Используем встроенные (без кириллицы)
            self.font_regular = 'Helvetica'
            self.font_bold = 'Helvetica-Bold'
        
    def create_kp(self, output_path, project_name, items, furniture_items=None, photo_path=None, total_sum=0, phone_number=None):
        """
        Создает PDF коммерческого предложения
        """
        c = canvas.Canvas(output_path, pagesize=A4)
        
        # СТРАНИЦА 1: Обложка (вертикальная)
        cover_names = ['cover_vert.png', 'cover_vert.jpg', 'cover_vert.jpeg', 'cover-vert.png', 'cover-vert.jpg', 'cover-vert.jpeg']
        base_paths = ['templates', '/opt/smetchikbot/templates']
        
        cover_found = False
        for base in base_paths:
            for name in cover_names:
                cover_path = os.path.join(base, name)
                if os.path.exists(cover_path):
                    try:
                        # Для вертикальной обложки используем портретную ориентацию A4
                        c.drawImage(cover_path, 0, 0, width=self.width, height=self.height, preserveAspectRatio=False)
                        
                        # Если указан номер телефона, добавляем его в пустой прямоугольник
                        if phone_number:
                            # Добавляем номер в позицию пустого прямоугольника (внизу слева)
                            c.setFont(self.font_bold, 14)
                            c.setFillColorRGB(1, 1, 1)  # Белый текст
                            c.drawString(40, 55, f"тел. {phone_number}")
                        
                        c.showPage()
                        cover_found = True
                        break
                    except Exception as e:
                        print(f"Ошибка загрузки обложки: {e}")
            if cover_found:
                break
        
        # СТРАНИЦА 2+: Материалы (каждый на отдельной странице)
        for idx, item in enumerate(items):
            if idx > 0 or cover_found:  # Создаем новую страницу только если это не первый материал или была обложка
                c.showPage()
            
            # Заголовок материала
            c.setFont(self.font_bold, 20)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(50, self.height - 60, item['name'])
            
            y_pos = self.height - 120
            
            # Фото материала (если есть)
            item_photo = item.get('image')
            if not item_photo or not os.path.exists(item_photo):
                item_photo = photo_path  # Fallback на общее фото
            
            if item_photo and os.path.exists(item_photo):
                try:
                    img = Image.open(item_photo)
                    img_width, img_height = img.size
                    
                    max_width = 500
                    max_height = 400
                    
                    scale = min(max_width / img_width, max_height / img_height)
                    new_width = img_width * scale
                    new_height = img_height * scale
                    
                    x_pos = (self.width - new_width) / 2
                    
                    c.drawImage(item_photo, x_pos, y_pos - new_height,
                               width=new_width, height=new_height, preserveAspectRatio=True)
                    
                    y_pos -= (new_height + 50)
                except Exception as e:
                    print(f"Ошибка загрузки изображения материала: {e}")
            
            # Детали материала
            c.setFont(self.font_regular, 11)
            details = item.get('details', [])
            for detail in details:
                if y_pos < 100:
                    c.showPage()
                    y_pos = self.height - 80
                c.drawString(50, y_pos, detail)
                y_pos -= 20
            
            # Стоимость материала
            y_pos -= 20
            if y_pos < 100:
                c.showPage()
                y_pos = self.height - 80
            
            c.setFont(self.font_bold, 16)
            item_cost = item.get('cost', 0)
            cost_text = f"Итоговая стоимость: {item_cost:,.1f} руб".replace(',', ' ')
            c.drawString(50, y_pos, cost_text)
        
        # СТРАНИЦЫ С МЕБЕЛЬЮ (каждая позиция на отдельной странице)
        if furniture_items:
            for furniture in furniture_items:
                c.showPage()
                
                # Название мебели
                c.setFont(self.font_bold, 20)
                c.setFillColorRGB(0, 0, 0)
                c.drawString(50, self.height - 60, furniture['name'])
                
                y_pos = self.height - 120
                
                # Изображение мебели
                if furniture.get('image') and os.path.exists(furniture['image']):
                    try:
                        img = Image.open(furniture['image'])
                        img_width, img_height = img.size
                        
                        max_width = 500
                        max_height = 400
                        
                        scale = min(max_width / img_width, max_height / img_height)
                        new_width = img_width * scale
                        new_height = img_height * scale
                        
                        x_pos = (self.width - new_width) / 2
                        
                        c.drawImage(furniture['image'], x_pos, y_pos - new_height,
                                   width=new_width, height=new_height, preserveAspectRatio=True)
                        
                        y_pos -= (new_height + 50)
                    except Exception as e:
                        print(f"Ошибка загрузки изображения мебели: {e}")
                
                # Информация о мебели
                c.setFont(self.font_regular, 14)
                c.drawString(50, y_pos, f"Количество: {furniture['quantity']}")
                y_pos -= 30
                c.drawString(50, y_pos, f"Стоимость за 1 шт: {furniture['price_per_unit']} руб")
                y_pos -= 30
                c.setFont(self.font_bold, 16)
                c.drawString(50, y_pos, f"Итоговая стоимость: {furniture['total_price']} руб")
        
        # СОВМЕЩЕННАЯ ТАБЛИЦА (материалы + мебель)
        c.showPage()
        
        # Заголовок
        c.setFont(self.font_bold, 18)
        c.setFillColorRGB(0, 0, 0)
        c.drawCentredString(self.width / 2, self.height - 60, project_name)
        
        # Собираем все элементы
        all_items = []
        
        # Добавляем материалы с их стоимостью
        for item in items:
            item_cost = item.get('cost', 0)
            quantity = item.get('quantity', 1)
            all_items.append({
                'name': item['name'],
                'quantity': quantity,
                'price_c': round(item_cost, 1),
                'price_d': round(item_cost, 1)
            })
        
        # Добавляем мебель
        if furniture_items:
            for furn in furniture_items:
                quantity = furn['quantity']
                if isinstance(quantity, str):
                    quantity = int(quantity)
                
                price_per_unit = furn['price_per_unit']
                if isinstance(price_per_unit, str):
                    price_per_unit = float(price_per_unit.replace(' ', '').replace(',', '').replace('руб', '').strip())
                
                total_price = furn['total_price']
                if isinstance(total_price, str):
                    total_price = float(total_price.replace(' ', '').replace(',', '').replace('руб', '').strip())
                
                all_items.append({
                    'name': furn['name'],
                    'quantity': quantity,
                    'price_c': round(price_per_unit, 1),
                    'price_d': round(total_price, 1)
                })
        
        # Рисуем таблицу
        y_start = self.height - 120
        x_start = 50
        
        # Ширины столбцов
        col_widths = [280, 70, 90, 90]
        row_height = 25
        
        # Рисуем данные (без заголовков)
        y_pos = y_start
        c.setFont(self.font_regular, 10)
        
        total_col_c = 0
        total_col_d = 0
        
        for item in all_items:
            # Рисуем границы ячеек
            c.setStrokeColorRGB(0, 0, 0)
            c.setFillColorRGB(1, 1, 1)  # Белый фон
            c.rect(x_start, y_pos - row_height, sum(col_widths), row_height, fill=1)
            
            c.setFillColorRGB(0, 0, 0)  # Черный текст
            
            # Наименование
            c.drawString(x_start + 5, y_pos - row_height + 8, item['name'])
            
            # Количество
            c.drawString(x_start + col_widths[0] + 5, y_pos - row_height + 8, str(item['quantity']))
            
            # Цена 1
            if item['price_c'] > 0:
                price_c_text = f"{item['price_c']:,.1f}".replace(',', ' ')
                c.drawRightString(x_start + col_widths[0] + col_widths[1] + col_widths[2] - 5, 
                                y_pos - row_height + 8, price_c_text)
                total_col_c += item['price_c']
            
            # Цена 2
            if item['price_d'] > 0:
                price_d_text = f"{item['price_d']:,.1f}".replace(',', ' ')
                c.drawRightString(x_start + sum(col_widths) - 5, y_pos - row_height + 8, price_d_text)
                total_col_d += item['price_d']
            
            y_pos -= row_height
        
        # Строка итого
        c.setFillColorRGB(0.9, 0.9, 0.9)  # Светло-серый фон
        c.rect(x_start, y_pos - row_height, sum(col_widths), row_height, fill=1)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont(self.font_bold, 11)
        c.drawString(x_start + 5, y_pos - row_height + 8, "Итого")
        
        # Итого колонка C
        if total_col_c > 0:
            c.drawRightString(x_start + col_widths[0] + col_widths[1] + col_widths[2] - 5, 
                             y_pos - row_height + 8, f"{total_col_c:,.1f}".replace(',', ' '))
        
        # Итого колонка D (с желтым фоном)
        c.setFillColorRGB(1, 1, 0)  # Желтый фон
        c.rect(x_start + col_widths[0] + col_widths[1] + col_widths[2], y_pos - row_height, 
               col_widths[3], row_height, fill=1)
        
        c.setFillColorRGB(0, 0, 0)
        c.drawRightString(x_start + sum(col_widths) - 5, y_pos - row_height + 8, 
                         f"{total_col_d:,.1f}".replace(',', ' '))
        
        # Добавляем примечание под таблицей
        y_pos -= row_height + 20
        c.setFont(self.font_regular, 10)
        c.setFillColorRGB(0, 0, 0)
        
        note_lines = [
            "Услуги упаковка, доставка, подъем, сборка и монтаж, сбор мусора +10% к стоимости",
            "-3% Скидка за наличные",
            "Срок 45 рабочих дней"
        ]
        
        for line in note_lines:
            c.drawString(x_start, y_pos, line)
            y_pos -= 15
        
        # ФИНАЛЬНЫЕ СТРАНИЦЫ (end1, end2, end3) - добавляем только один раз
        end_names = ['end1', 'end2', 'end3']
        end_formats = ['.png', '.jpg', '.jpeg']
        
        for end_name in end_names:
            end_found = False
            for base in base_paths:
                if end_found:
                    break
                for fmt in end_formats:
                    if end_found:
                        break
                    end_path = os.path.join(base, end_name + fmt)
                    if os.path.exists(end_path):
                        try:
                            c.showPage()
                            
                            # Загружаем изображение
                            img = Image.open(end_path)
                            img_width, img_height = img.size
                            
                            # Конвертируем в дюймы (72 DPI для PDF)
                            img_width_pts = img_width * 72 / 96
                            img_height_pts = img_height * 72 / 96
                            
                            # Масштабируем если больше страницы
                            if img_width_pts > self.width or img_height_pts > self.height:
                                scale = min(self.width / img_width_pts, self.height / img_height_pts)
                                img_width_pts *= scale
                                img_height_pts *= scale
                            
                            # Центрируем
                            x = (self.width - img_width_pts) / 2
                            y = (self.height - img_height_pts) / 2
                            
                            c.drawImage(end_path, x, y, width=img_width_pts, height=img_height_pts, preserveAspectRatio=True)
                            end_found = True
                        except Exception as e:
                            print(f"Ошибка добавления финальной страницы {end_path}: {e}")
        
        c.save()
        return output_path

def generate_kp_pdf(project_name, items_data, furniture_data=None, photo_path=None, phone_number=None):
    """
    Генерирует PDF КП
    
    Args:
        project_name: название проекта
        items_data: список словарей с данными об изделиях (материалы)
        furniture_data: список словарей с данными о мебели
        photo_path: путь к фото проекта
        phone_number: номер телефона для обложки
    
    Returns:
        путь к созданному PDF файлу
    """
    generator = PDFGenerator()
    
    # Формируем данные для PDF
    items = []
    total = 0
    
    # Сохраняем стоимости материалов для таблицы
    materials_costs = []
    
    for item_data in items_data:
        details = [
            f"Каркас из {item_data.get('body_material', 'ЛДСП')}",
            f"Фасады {item_data.get('facade_description', 'МДФ')}",
        ]
        
        if item_data.get('additional_info'):
            details.extend(item_data['additional_info'])
        
        item_cost = item_data.get('total_cost', 0)
        
        items.append({
            'name': item_data['name'],
            'details': details,
            'cost': item_cost,
            'image': item_data.get('image'),  # Добавляем изображение
            'quantity': item_data.get('quantity', 1)  # Добавляем количество
        })
        
        materials_costs.append(item_cost)
        total += item_cost
    
    # Формируем данные о мебели
    furniture_items = []
    if furniture_data:
        for furn in furniture_data:
            furniture_items.append({
                'name': furn['name'],
                'image': furn.get('image'),
                'quantity': furn['quantity'],
                'price_per_unit': furn['price_per_unit'],
                'total_price': furn['total_price']
            })
    
    # Создаем PDF
    output_path = f"outputs/{project_name.replace(' ', '_')}.pdf"
    os.makedirs('outputs', exist_ok=True)
    
    generator.create_kp(output_path, project_name, items, furniture_items, photo_path, total, phone_number)
    
    return output_path
