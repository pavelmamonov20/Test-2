import pandas as pd

# Читаем Excel файл
df = pd.read_excel('/workspace/barcodes_CC001_to_CC200.xlsx', header=None, nrows=5)

print("Первые 5 строк Excel-файла:")
print(df.head())

# Проверим общее количество строк
wb = pd.ExcelFile('/workspace/barcodes_CC001_to_CC200.xlsx')
df_full = pd.read_excel(wb, header=None)
print(f"\nОбщее количество строк (включая заголовок): {len(df_full)}")
print(f"Количество данных строк: {len(df_full) - 1}")  # минус заголовок