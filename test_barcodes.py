from barcode import Code128
from barcode.writer import ImageWriter
import os

# Тестирование генерации штрихкодов
test_codes = ["CC001", "CC050", "CC100", "CC150", "CC200"]

for code in test_codes:
    # Создаем штрихкод
    barcode = Code128(code, writer=ImageWriter())
    filename = f'/workspace/test_{code}'
    barcode.save(filename)
    print(f"Штрихкод для {code} сохранен как {filename}.png")
    
# Проверим, что все файлы были созданы
for code in test_codes:
    filename = f'/workspace/test_{code}.png'
    if os.path.exists(filename):
        print(f"✓ Файл {filename} существует")
        # Удалим тестовый файл
        os.remove(filename)
    else:
        print(f"✗ Файл {filename} НЕ существует")