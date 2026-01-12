import os
from barcode import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import tempfile


def generate_barcodes_excel():
    # Создаем новую рабочую книгу
    wb = Workbook()
    ws = wb.active
    ws.title = "Штрихкоды"

    # Заголовки
    ws['A1'] = 'Номер'
    ws['B1'] = 'Штрихкод'

    # Устанавливаем ширину колонок
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 30

    # Временная директория для хранения изображений штрихкодов
    temp_dir = tempfile.mkdtemp()

    try:
        for i in range(1, 201):
            # Генерируем значение штрихкода (CC001 до CC200)
            barcode_value = f"CC{i:03d}"
            
            # Заполняем первый столбец (порядковый номер)
            ws[f'A{i+1}'] = i
            
            # Генерируем изображение штрихкода
            barcode_filename = os.path.join(temp_dir, f'barcode_{i}')
            
            # Создаем штрихкод Code128
            code128 = Code128(barcode_value, writer=ImageWriter())
            code128.save(barcode_filename)
            
            # Добавляем расширение .png, так как библиотека сохраняет файлы с этим расширением
            barcode_filename_with_ext = barcode_filename + '.png'
            
            # Проверяем, существует ли файл
            if not os.path.exists(barcode_filename_with_ext):
                raise FileNotFoundError(f"Barcode image was not created: {barcode_filename_with_ext}")
            
            # Открываем изображение и изменяем его размер для соответствия требованиям
            img = PILImage.open(barcode_filename_with_ext)
            
            # Устанавливаем высоту строки в 30 мм (в Excel единицы измерения - точки)
            # 30 мм примерно равно 85.04 точкам (1 мм ≈ 2.8346 точек)
            row_height_points = 85.04
            ws.row_dimensions[i+1].height = row_height_points
            
            # Добавляем изображение в ячейку
            excel_img = Image(barcode_filename_with_ext)
            
            # Масштабируем изображение, чтобы оно поместилось в ячейку
            # Учитываем, что ширина и высота изображения должны соответствовать размеру ячейки
            cell_width_pixels = 200  # Приблизительно соответствует ширине колонки B
            cell_height_pixels = int(row_height_points * 1.33)  # 1 pt ≈ 1.33 px
            
            original_width, original_height = img.size
            scale_factor = min(cell_width_pixels/original_width, cell_height_pixels/original_height)
            
            new_width = int(original_width * scale_factor)
            new_height = int(original_height * scale_factor)
            
            excel_img.width = new_width
            excel_img.height = new_height
            
            # Добавляем изображение в ячейку
            ws.add_image(excel_img, f'B{i+1}')
            
            # Применяем форматирование ячеек для лучшего центрирования
            cell = ws[f'B{i+1}']
            from openpyxl.styles import Alignment
            # Центрируем содержимое ячейки по горизонтали и вертикали
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
        # Сохраняем файл
        wb.save('barcodes_CC001_to_CC200.xlsx')
        print("Excel-файл с штрихкодами успешно создан: barcodes_CC001_to_CC200.xlsx")
        
    finally:
        # Удаляем временные файлы
        import shutil
        shutil.rmtree(temp_dir)


if __name__ == "__main__":
    generate_barcodes_excel()