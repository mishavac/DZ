import re
import pandas as pd
from openpyxl.utils import get_column_letter
import csv

# Загрузите файл Excel в DataFrame
df = pd.read_excel(r'C:\Users\NOVIKOV\Desktop\ebala\raspisanie.xlsx', sheet_name='16.10.2023-21.10.2023', header=None)

# Ищем колонку с названием "Группа:"
column_name = 'Группа:'
group_row = None

for index, row in df.iterrows():
    for col in df.columns:
        cell_value = row[col]
        if isinstance(cell_value, str) and column_name.lower() in cell_value.lower():
            print(f'Найдено в строке {index + 1}, колонка {col + 1}')
            group_row = row
            break

    if group_row is not None:
        break

if group_row is not None:
    print("Найденная строка:")
    data = []
    for i, value in enumerate(group_row):
        column_letter = get_column_letter(i + 1)
        data.append((index + 1, i + 1, value))

    # Удаляем строки с NaN
    data = [line for line in data if 'nan' not in str(line[2]).lower()]

    # Записываем результаты в CSV файл
    csv_filename = 'result.csv'
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['Строка', 'Колонка', 'Группа', 'Преподаватель', 'Номер пары', 'Время пары'])

        for line in data:
            row_data = df.iloc[line[0]-1:, line[1]-1]
            for number_pair, time_pair, value_row in zip(df.iloc[line[0]-1:, 1], df.iloc[line[0]-1:, 3], row_data.dropna()):
                csv_writer.writerow([line[0], line[1], value_row, df.iloc[line[0]-1, line[1]-1], number_pair, time_pair])

    print(f'Результаты сохранены в файл {csv_filename}')
