import pandas as pd
from collections import Counter
from datetime import datetime, timedelta
from UliPlot.XLSX import auto_adjust_xlsx_column_width




def find_duplicates():
    date_format = '%Y_%m_%d'
    yesterday = datetime.now() - timedelta(days=1)
    yesterday = yesterday.strftime(date_format)

    file_name = f'Alarms_{yesterday}.xls'
    sheet_name = 'Лист1'
# Чтение данных из Excel
    df = pd.read_excel(file_name, sheet_name=sheet_name)

# Выбор столбцов B, C и D, начиная с четвертой строки
    df_selected = df.iloc[3:, [1, 2, 4, 5]]

# Объединение значений в каждой строке в одну строку
    df_combined = df_selected.apply(lambda row: ';'.join(row.values.astype(str)), axis=1)
# Преобразование объединенных значений в список
    data = df_combined.tolist()
    # Подсчет повторяющихся значений
    counter = Counter(data)
    # Создание нового DataFrame и запись данных
    duplicates_df = pd.DataFrame.from_records(list(counter.items()), columns=['Сообщение;Класс сообщения;Состояние;Мнемосхема', 'Dublicates'])
        # Сортировка DataFrame по количеству повторений в порядке убывания
    duplicates_df = duplicates_df.sort_values(by='Dublicates', ascending=False)
# Разделение столбца 0 на несколько столбцов
    duplicates_df[['Сообщение', 'Класс сообщения', 'Состояние', 'Мнемосхема']] = duplicates_df['Сообщение;Класс сообщения;Состояние;Мнемосхема'].str.split(';', expand=True)
    duplicates_df.drop(columns=['Сообщение;Класс сообщения;Состояние;Мнемосхема'])
    duplicates_df = duplicates_df[['Сообщение', 'Класс сообщения', 'Состояние', 'Мнемосхема', 'Dublicates']]
    column_to_sum = duplicates_df[['Dublicates']]
    duplicates_df.loc['Total'] = column_to_sum.sum()
    

# Запись DataFrame обратно в Excel
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        duplicates_df.to_excel(writer, sheet_name='NoDublicates', index=False)
        auto_adjust_xlsx_column_width(duplicates_df, writer, sheet_name="NoDublicates", index=False)
        
    print('Данные успешно записаны в новый лист "NoDublicates"')

if __name__ == '__main__':
    find_duplicates()