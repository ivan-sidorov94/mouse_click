import pyautogui
import subprocess
import os
import time
import configparser
import pandas as pd
import babel.numbers
import tzdata
import tkinter as tk

from datetime import datetime, timedelta
from tkcalendar import DateEntry
from UliPlot.XLSX import auto_adjust_xlsx_column_width
from collections import Counter

config_file = "C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms\settings.ini"
config = configparser.RawConfigParser()

if not os.path.exists(config_file):
    config.add_section('Time')
    config.set('Time', 'time1', '45')
    config.set('Time', 'time2', '0.1')
    config.set('Time', 'time3', '20')
    with open(config_file, 'w') as configfile:
        config.write(configfile)

if config.read(config_file):
    try:
        time1 = config.getfloat('Time', 'time1')
        time2 = config.getfloat('Time', 'time2')
        time3 = config.getfloat('Time', 'time3')
    except:
        time1 = 45
        time2 = 0.5
        time3 = 15
        config.clear()
        config.add_section('Time')
        config.set('Time', 'time1', '45')
        config.set('Time', 'time2', '0.1')
        config.set('Time', 'time3', '20')
        with open(config_file, 'w') as configfile:
            config.write(configfile)

def alarms(date_format, int1, int2, int3, date1, date2):
    date1 = datetime.strftime(date1, date_format)
    date2 = datetime.strftime(date2, date_format)

    config.set("Time", "time1", int1)
    config.set("Time", "time2", int2)
    config.set("Time", "time3", int3)
    with open(config_file, 'w') as configfile:
        config.write(configfile)

    cmd = f'C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms\InfinityAlarms.exe HISTORY DBEG="{date1}" DEND"={date2}" Bookmark="Новая закладка"'
    p = subprocess.Popen(cmd)
    time.sleep(int1)
    pyautogui.click(x=1939, y=31)
    pyautogui.sleep(int2)
    pyautogui.click(x=1960, y=74)
    pyautogui.sleep(int2)
    pyautogui.click(x=2667, y=582)
    pyautogui.sleep(int2)
    pyautogui.click(x=2980, y=745)
    pyautogui.sleep(int2)
    pyautogui.click(x=2958, y=689)
    pyautogui.sleep(int3)
    #pyautogui.click(x=3050, y=623)
    p.kill()

def open_directory():

    directory = 'C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms'
    os.startfile(directory)

def obrabotka(date):
    date_format = '%Y_%m_%d'
    date = datetime.strftime(date, date_format)


    file_name = f'C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms\Alarms_{date}.xls'
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

def main():

    # Вычисление вчерашней и позавчерашней даты
    date_format = '%d.%m.%Y'
    #before_yesterday = datetime.now() - timedelta(days=2)
    yesterday = datetime.now() - timedelta(days=1)
    today = datetime.now()
    #before_yesterday = before_yesterday.strftime(date_format)
    yesterday = yesterday.strftime(date_format)
    today = today.strftime(date_format)

    root = tk.Tk()
    root.title("Выгрузка алармов за предыдущий день")

    frame1 = tk.Frame(root)
    frame2 = tk.Frame(root)
    frame3 = tk.Frame(root)

    # Создание меток и полей выбора даты
    label1 = tk.Label(frame1, text="Выберите начальную дату:")
    date_entry1 = DateEntry(frame2, date_pattern='dd.mm.yyyy')
    date_entry1.set_date(yesterday)

    label2 = tk.Label(frame1, text="Выберите конечную дату:")
    date_entry2 = DateEntry(frame2, date_pattern='dd.mm.yyyy')
    date_entry2.set_date(today)

# Создание меток и полей ввода
    label3 = tk.Label(frame3, text="Введите задержку открытия программы:")
    int1 = tk.Entry(frame3)
    int1.insert(0, str(time1))

    label4 = tk.Label(frame3, text="Введите задержку перемещения мыши:")
    int2 = tk.Entry(frame3)
    int2.insert(0, str(time2))

    label5 = tk.Label(frame3, text="Введите задержку перед сохранением:")
    int3 = tk.Entry(frame3)
    int3.insert(0, str(time3))

    # # Создание кнопок
    button1 = tk.Button(frame3, text="Запустить программу", command=lambda: alarms(date_format, float(int1.get()), float(int2.get()), 
                                                                                   float(int3.get()), date_entry1.get_date(), date_entry2.get_date()))    
    button2 = tk.Button(frame3, text="Открыть каталог", command=open_directory)

    button3 = tk.Button(frame3, text="Запустить обработку", command=obrabotka(date_entry1.get_date()))
   
# Размещение меток и полей ввода на экране
    label1.pack(side='left', padx=5)
    label2.pack(side='left', padx=5)
    date_entry1.pack(side='left', padx=30)
    date_entry2.pack(side='left', padx=30)

    label3.pack()
    int1.pack()
    label4.pack()
    int2.pack()
    label5.pack()
    int3.pack()
    button1.pack()
    button2.pack()
    button3.pack()

    frame1.pack()
    frame2.pack()
    frame3.pack()

# Запуск главного цикла
    root.mainloop()

if __name__ == "__main__":
    main()
