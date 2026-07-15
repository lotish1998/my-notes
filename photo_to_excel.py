import os
from openpyxl import openpyxl #, load_workbook, Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
# from openpyxl.styles import Alignment
    # За допомогою цього коду були добавленні фото до ексел таблиці
    # Є представлені два варіанти

# Вкажіть назву вашого Excel файлу
excel_file = 'тут шлях до таблиці'
# Вкажіть шлях до папки з фото (наприклад, 'C:/photos/' або './photos/')
photo_folder = 'тут шлях до папки з фото'

# Відкриваємо файл Excel та активний аркуш
wb = openpyxl.load_workbook("book_it_py.xlsx")
ws = wb.active

# Вкажіть стовпчик з ІМ'ЯМ (наприклад, A) та стовпчик для ФОТО (наприклад, B)
name_column = 'B'
photo_column = 'E'

# Проходимося по всіх рядках від 2 до останнього (припускаємо, що 1-й рядок — це заголовки)
for row in range(2, ws.max_row + 1):
    name_cell = ws[f'{name_column}{row}']
    
    # Якщо ім'я є в комірці
    if name_cell.value:
        # Формуємо ім'я файлу (додаємо .png)
        file_name = f"{name_cell.value}.jpg"
        photo_path = os.path.join(photo_folder, file_name)
        
        # Перевіряємо, чи існує файл фотографії
        if os.path.exists(photo_path):
            img = Image(photo_path)
            
            # Вказуємо розмір фото (за потреби змініть цифри 100 на інші)
            img.width = 70
            img.height = 100
            
            # Встановлюємо висоту рядка та ширину стовпчика відповідно до розміру фото
            ws.row_dimensions[row].height = 105
            ws.column_dimensions[photo_column].width = 15
            
            # Вказуємо клітинку, куди вставити фото
            cell_coordinate = f'{photo_column}{row}'
            ws.add_image(img, cell_coordinate)

# Зберігаємо результати у новий файл
wb.save('таблиця_з_фото.xlsx')
print("Фото успішно вставлено!")



# # 1. Шляхи до файлів (Оригінал і Новий результат)
# SOURCE_EXCEL = "/Users//.xlsx"  # Звідси ТІЛЬКИ ЧИТАЄМО
# RESULT_EXCEL = "/Users//.xlsx" # Сюди ПИШЕМО (новий файл)
# IMAGE_FOLDER = "/Users//Photo_books"
# IMAGE_EXTENSION = ".jpg"

# if not os.path.exists(SOURCE_EXCEL):
#     print(f"Помилка: Оригінальний файл {SOURCE_EXCEL} не знайдено!")
#     exit()

# # 2. Створюємо абсолютно нову чисту книгу для результату
# wb_source = load_workbook(SOURCE_EXCEL, data_only=True)
# ws_source = wb_source.active

# wb_res = Workbook()
# ws_res = wb_res.active
# ws_res.title = ws_source.title

# # Налаштовуємо стовпчик E (5) у новій таблиці
# ws_res.column_dimensions['E'].width = 70

# print("Обробка файлу розпочата...")

# # 3. Переносимо заголовки (рядок 1)
# for col in range(1, ws_source.max_column + 1):
#     ws_res.cell(row=1, column=col, value=ws_source.cell(row=1, column=col).value)

# # 4. Переносимо дані та вставляємо фотографії
# for row in range(2, ws_source.max_row + 1):
#     # Спочатку копіюємо всі текстові дані з оригінального рядка в новий
#     for col in range(1, ws_source.max_column + 1):
#         ws_res.cell(row=row, column=col, value=ws_source.cell(row=row, column=col).value)
    
#     # Налаштовуємо висоту та перенесення тексту для нового рядка
#     ws_res.row_dimensions[row].height = 100
#     for col in range(1, ws_source.max_column + 1):
#         ws_res.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    
#     # Беремо назву книги з 2-го стовпчика (B)
#     book_name = ws_source.cell(row=row, column=2).value
#     if not book_name:
#         continue
        
#     book_name = str(book_name).strip()
    
#     # Формуємо шлях до фото
#     full_image_name = f"{book_name}{IMAGE_EXTENSION}"
#     image_path = os.path.join(IMAGE_FOLDER, full_image_name)
    
#     # Перевіряємо та вставляємо
#     if os.path.exists(image_path):
#         try:
#             img = OpenpyxlImage(image_path)
#             # Задаємо розміри фото
#             img.height = 130
#             img.width = 130
            
#             # Вставляємо у стовпчик E (5) нового файлу
#             ws_res.add_image(img, f"E{row}")
#             print(f"Успішно: Фото для '{book_name}' додано в E{row}")
#         except Exception as e:
#             print(f"Помилка з фото '{book_name}': {e}")
#     else:
#         # Якщо фото немає — чітко пишемо текст у стовпчик E
#         ws_res.cell(row=row, column=5, value="Немає фото")
#         print(f"Файл не знайдено: {image_path} (вписано 'Немає фото')")

# # 5. Зберігаємо у НОВИЙ файл
# wb_res.save(RESULT_EXCEL)
# print(f"\nВсе готово! Новий чистий файл збережено у: {RESULT_EXCEL}")
