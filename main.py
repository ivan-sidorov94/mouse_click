import pyautogui
import subprocess
import datetime as dt
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkcalendar import DateEntry



def alarms(int1, int2, int3, date1, date2):
    print(int1, int2, int3, date1, date2)

    # cmd = f'C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms\InfinityAlarms.exe HISTORY DBEG="{date1}" DEND"={date2}"'
    # p = subprocess.Popen(cmd)
    # time.sleep(int1)
    # pyautogui.click(x=1939, y=31)
    # pyautogui.sleep(int2)
    # pyautogui.click(x=1960, y=74)
    # pyautogui.sleep(int2)
    # pyautogui.click(x=2667, y=582)
    # pyautogui.sleep(int2)
    # pyautogui.click(x=2980, y=745)
    # pyautogui.sleep(int2)
    # pyautogui.click(x=2958, y=689)
    # pyautogui.sleep(int3)
    # pyautogui.click(x=3050, y=623)
    # p.kill()

def open_directory():

    directory = 'C:\Program Files (x86)\EleSy\SCADA Infinity\InfinityAlarms'
    os.startfile(directory)

def main():
    date_format = '%d.%m.%Y'
    before_yesterday = datetime.now() - timedelta(days=2)
    yesterday = datetime.now() - timedelta(days=1)
    before_yesterday = before_yesterday.strftime(date_format)
    yesterday = yesterday.strftime(date_format)
    root = tk.Tk()
    root.title("Выгрузка алармов за предыдущий день")
# Создание меток и полей ввода
    label1 = tk.Label(root, text="Введите задержку открытия программы:")
    int1 = tk.Entry(root)
    label2 = tk.Label(root, text="Введите задержку перемещения мыши:")
    int2 = tk.Entry(root)
    label3 = tk.Label(root, text="Введите задержку перед сохранением:")
    int3 = tk.Entry(root)
    # Создание меток и полей выбора даты
    label4 = tk.Label(root, text="Выберите первую дату:")
    date_entry1 = DateEntry(root, date_pattern='dd.mm.yyyy')
    date_entry1.set_date(before_yesterday)
    label5 = tk.Label(root, text="Выберите вторую дату:")
    date_entry2 = DateEntry(root)
    date_entry2.set_date(yesterday)
    
    
    date1 = date_entry1.get_date()
    date2 = date_entry2.get_date()


    

    if date_entry1.format_date() != before_yesterday or date_entry2.format_date() != yesterday:
        date1 = date1
        date2 = date2
        print(1)
    else:
        date1 = before_yesterday
        date2 = yesterday
        print(2)
    # # Создание кнопок
    button1 = tk.Button(root, text="Запустить программу", command=lambda: alarms(float(int1.get()), float(int2.get()), float(int3.get()), date1, date2))    
    button2 = tk.Button(root, text="Открыть каталог", command=open_directory)
    int1.insert(0, "20")
    int2.insert(0, "0.5")
    int3.insert(0, "10")

   
# Размещение меток и полей ввода на экране
    label1.pack()
    int1.pack()
    label2.pack()
    int2.pack()
    label3.pack()
    int3.pack()
    button1.pack()
    button2.pack()
    label4.pack()
    date_entry1.pack()
    label5.pack()
    date_entry2.pack()
    #root.geometry("300x200")
# Запуск главного цикла
    root.mainloop()

if __name__ == "__main__":
    main()