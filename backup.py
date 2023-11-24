import pandas as pd

# Загрузите файл Excel в DataFrame
df = pd.read_excel('raspisanie.xlsx', header=None)

start_row_index = 9
start_column_index = 5

# Количество итераций цикла
num_iterations = 15  # Измените это на нужное количество итераций
for i in range(num_iterations):
    # Выбор конкретной ячейки
    cell_value = df.iat[start_row_index, start_column_index]
    
    # Вывод значения
    print(f"Значение в строке {start_row_index} и колонке {start_column_index}: {cell_value}")
    
    # Увеличение индекса колонки на 4
    start_column_index += 5


start_row_indexs = 9
start_column_indexs = 9

for i in range(num_iterations):
    cell_value = df.iat[start_row_indexs, start_column_indexs]

    # Проверка, не является ли значение NaN
    if pd.notna(cell_value):
        print(f"Значение в строке {start_row_indexs} и колонке {start_column_indexs}: {cell_value}")
        
        # Увеличение индекса колонки на 5
        
        start_column_indexs += 4
        print(f" колонка кабинет  {start_column_indexs}: {cell_value}")
    
    # Сброс индекса колонки только после завершения внутреннего условия
    start_column_indexs = 5
    start_row_indexs += 1

