import pyautogui
import subprocess
import os
import time
import configparser
import pandas as pd
import babel.numbers
import tzdata
import tkinter as tk
import shutil

from tkinter import filedialog
from openpyxl.styles import Border, Side
from openpyxl.worksheet.table import Table
from datetime import datetime, timedelta
from tkcalendar import DateEntry
from UliPlot.XLSX import auto_adjust_xlsx_column_width

path_to_config = "C:\\Program Files (x86)\\EleSy\\SCADA Infinity\\InfinityAlarms\\settings.ini"


path_to_InfinityAlarms = "C:\\Program Files (x86)\\EleSy\\SCADA Infinity\\InfinityAlarms\\InfinityAlarms.exe"

path_to_raw_file = "C:\\Program Files (x86)\\EleSy\\SCADA Infinity\\InfinityAlarms"


default_path_to_save = "D:\\Alarms_ASUTP"

config_file = path_to_config
config = configparser.RawConfigParser()

if not os.path.exists(config_file):
    config.add_section('Time')
    config.set('Time', 'time1', '45')
    config.set('Time', 'time2', '0.1')
    config.set('Time', 'time3', '20')
    config.add_section('Start')
    config.set('Start', 'autostart', 'False')
    config.set('Save Directory', 'save_folder_path', f'{default_path_to_save}')
    with open(config_file, 'w') as configfile:
        config.write(configfile)

if config.read(config_file):
    try:
        time1 = config.getfloat('Time', 'time1')
        time2 = config.getfloat('Time', 'time2')
        time3 = config.getfloat('Time', 'time3')
        autostart = config.getboolean('Start', 'autostart')
        save_folder = config.get('Save Directory', 'save_folder_path')

    except (configparser.NoSectionError, configparser.NoOptionError):
        time1 = 45
        time2 = 0.5
        time3 = 15
        config.clear()
        config.add_section('Time')
        config.set('Time', 'time1', '45')
        config.set('Time', 'time2', '0.1')
        config.set('Time', 'time3', '20')
        config.add_section('Start')
        config.set('Start', 'autostart', 'False')
        config.add_section('Save Directory')
        config.set('Save Directory', 'save_folder_path', f'{default_path_to_save}')

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

    cmd = (f'{path_to_InfinityAlarms} HISTORY DBEG="{date1}" DEND"={date2}" '
           'Bookmark="Новая закладка"')
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
    # pyautogui.click(x=3050, y=623)
    p.kill()


def open_directory():
    save_folder_path = config.get('Save Directory', 'save_folder_path')
    try:
        os.startfile(save_folder_path)
    except FileNotFoundError:
        os.makedirs(save_folder_path, exist_ok=True)
        os.startfile(save_folder_path)


def obrabotka(date1, date2):

    date_format = '%Y_%m_%d'
    date1 = datetime.strftime(date1, date_format)
    date2 = datetime.strftime(date2, '%Y.%m.%d')

    file_name = rf'{path_to_raw_file}\Alarms_{date1}.xls'
    sheet_name = 'Лист1'

    # Читаем только необходимые столбцы
    df = pd.read_excel(file_name, sheet_name=sheet_name, usecols=[0, 1, 2, 4, 5],
                       skiprows=2, names=['Время','Сообщение', 'Класс сообщения',
                                          'Состояние', 'Мнемосхема'], decimal=',', dtype=str)

    # Заменяем NaN на пустую строку
    df = df.fillna('')

    # Создаем новый фрейм для ремонта
    df_rem = df[df['Сообщение'].str.contains('ремонт"') | df['Сообщение'].str.contains('ремонта"')]

# Удаляем ненужные столбцы из фреймов
    df = df.drop(columns=['Время'])
    df_rem = df_rem.drop(columns=['Класс сообщения','Состояние', 'Мнемосхема'])

# Создаем временный фрейм для обработки сообщений о ремонте, разбивая столбец Сообщение на 2
    new_df = df_rem['Сообщение'].str.split(' П', expand=True, n=-1)

# Удаляем столбец перед конкатенацией
    df_rem = df_rem.drop(columns='Сообщение')

# Соединяем оба фрейма
    df_rem = pd.concat([df_rem, new_df], axis=1)

# Удаляем из памяти временный фрейм
    del new_df

# Переименовываем столбцы
    df_rem.columns = ['Время','Датчик|ИМ','Команда']

# Изменяем значения в столбце Команда в зависимости от значения
    df_rem.loc[df_rem['Команда'].str.contains('Снять'), 'Команда'] =  'Снят с ремонта'
    df_rem.loc[df_rem['Команда'].str.contains('Вывести'), 'Команда'] =  'Выведен в ремонт'

# Подсчитываем дубликаты напрямую
    df = df.groupby(list(df.columns)).size().reset_index(name='Dublicates')

# Сортируем по количеству дубликатов
    df = df.sort_values(by='Dublicates', ascending=False)

# Считаем общее количество
    total_dub = df['Dublicates'].sum()

# Добавляем строку с общим количеством
    df.loc[len(df.index)] = ['', '', '','Всего сообщений:',total_dub]
# Создаем объект Border с границами
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name='NoDublicates', index=False)

    # Получаем объект workbook
        workbook = writer.book

# Получаем объект worksheet для другого листа
        worksheet = writer.sheets['Лист1']

# Получаем номер последней строки с данными
        num_rows = worksheet.max_row

# Создаем объект Table для диапазона данных на другом листе, начиная с 3 строки
        last_col_letter = 'F'  # Замените на букву последнего столбца
        tab = Table(displayName="Table1",ref="A3:" + last_col_letter + str(num_rows))

# Добавляем таблицу в другой лист
        worksheet.add_table(tab)

# Получаем объект worksheet
        worksheet = writer.sheets['NoDublicates']

        for row in worksheet.iter_rows():
            for cell in row:
# Применяем границы к каждой ячейке
                cell.border = thin_border

# Создаем объект Table для всего диапазона данных и добавляем фильтр
        tab = Table(displayName="Table2",ref="A1:E" + str(len(df.index) + 1))

# Добавляем таблицу в лист
        worksheet.add_table(tab)

# Автоматическое расширение столбцов
        auto_adjust_xlsx_column_width(df, writer, sheet_name="NoDublicates", index=False)

        df_rem.to_excel(writer, sheet_name='Ремонт', index=False)

        worksheet = writer.sheets['Ремонт']

        for row in worksheet.iter_rows():
            for cell in row:
# Применяем границы к каждой ячейке
                cell.border = thin_border

# Создаем объект Table для всего диапазона данных и добавляем фильтр
        tab = Table(displayName="Table3",ref="A1:C" + str(len(df.index) + 1))

# Добавляем таблицу в лист
        worksheet.add_table(tab)

        auto_adjust_xlsx_column_width(df_rem, writer, sheet_name="Ремонт", index=False)

    shutil.move(file_name,
              f'{save_folder}\\Отчет по алармам за период_{date2}.xls')

    if var:
        root.destroy()


def autostart_cmd():
    if var.get() == 1:
        config.set('Start', 'autostart', 'True')
    else:
        config.set('Start', 'autostart', 'False')
    with open(config_file, 'w') as f:
        config.write(f)


def set_save_folder():
    save_folder_path = filedialog.askdirectory()
    config.set('Save Directory', 'save_folder_path', f'{save_folder_path}')
    with open(config_file, 'w') as f:
        config.write(f)


def main():

    # Вычисление вчерашней и позавчерашней даты
    date_format = '%d.%m.%Y'
    # before_yesterday = datetime.now() - timedelta(days=2)
    yesterday = datetime.now() - timedelta(days=1)
    today = datetime.now()
    # before_yesterday = before_yesterday.strftime(date_format)
    yesterday = yesterday.strftime(date_format)
    today = today.strftime(date_format)

    global var
    global root

    root = tk.Tk()
    var = tk.IntVar(value=config.getboolean('Start', 'autostart'))
    root.title("Выгрузка алармов за предыдущий день")

    frame1 = tk.Frame(root)
    frame2 = tk.Frame(root)
    frame3 = tk.Frame(root)
    frame4 = tk.Frame(root)

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

    dir_button = tk.Button(frame3, text="Выбрать папку для сохранения",
                           command=set_save_folder)

    # # Создание кнопок
    button1 = tk.Button(frame3, text="Запустить Infinity Alarms", command=lambda:
                        alarms(date_format, float(int1.get()), float(int2.get()),
                               float(int3.get()), date_entry1.get_date(),
                               date_entry2.get_date()))
    button2 = tk.Button(frame3, text="Открыть каталог", command=open_directory)

    button3 = tk.Button(frame3, text="Запустить обработку", command=lambda:
                        obrabotka(date_entry1.get_date(),
                                  date_entry2.get_date()))
    check_button = tk.Checkbutton(frame3, text='Автостарт', variable=var,
                                  command=autostart_cmd)

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
    button3.pack()
    dir_button.pack()
    button2.pack()
    check_button.pack()

    frame1.pack()
    frame2.pack()
    frame4.pack()
    frame3.pack()

    if config.getboolean('Start', 'autostart'):
        alarms(date_format, float(int1.get()), float(int2.get()),
               float(int3.get()), date_entry1.get_date(),
               date_entry2.get_date())
        time.sleep(5)
        obrabotka(date_entry1.get_date(), date_entry2.get_date())

# Запуск главного цикла
    root.mainloop()


if __name__ == "__main__":
    main()
