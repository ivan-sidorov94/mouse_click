import pyautogui
import subprocess
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkcalendar import DateEntry



def alarms(date_format, int1, int2, int3, date1, date2):
    date1 = datetime.strftime(date1, date_format)
    date2 = datetime.strftime(date2, date_format)

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

    # Вычисление вчерашней и позавчерашней даты
    date_format = '%d.%m.%Y'
    before_yesterday = datetime.now() - timedelta(days=2)
    yesterday = datetime.now() - timedelta(days=1)
    before_yesterday = before_yesterday.strftime(date_format)
    yesterday = yesterday.strftime(date_format)

    root = tk.Tk()
    root.title("Выгрузка алармов за предыдущий день")

    frame1 = tk.Frame(root)
    frame2 = tk.Frame(root)
    frame3 = tk.Frame(root)

# Создание меток и полей ввода
    label1 = tk.Label(frame2, text="Введите задержку открытия программы:")
    int1 = tk.Entry(frame2)
    int1.insert(0, "20")

    label2 = tk.Label(frame2, text="Введите задержку перемещения мыши:")
    int2 = tk.Entry(frame2)
    int2.insert(0, "0.5")

    label3 = tk.Label(frame2, text="Введите задержку перед сохранением:")
    int3 = tk.Entry(frame2)
    int3.insert(0, "10")

    # Создание меток и полей выбора даты
    label4 = tk.Label(frame1, text="Выберите начальную дату:")
    date_entry1 = DateEntry(frame3)
    date_entry1.set_date(before_yesterday)

    label5 = tk.Label(frame1, text="Выберите конечную дату:")
    date_entry2 = DateEntry(frame3)
    date_entry2.set_date(yesterday)

    # # Создание кнопок
    button1 = tk.Button(frame2, text="Запустить программу", command=lambda: alarms(date_format, float(int1.get()), 
                                                                                 float(int2.get()), float(int3.get()), 
                                                                                 date_entry1.get_date(), date_entry2.get_date()))    
    button2 = tk.Button(frame2, text="Открыть каталог", command=open_directory)
   
# Размещение меток и полей ввода на экране
    label4.pack(side='left', padx=5)
    label5.pack(side='left', padx=5)
    date_entry1.pack(side='left', padx=30)
    date_entry2.pack(side='left', padx=30)

    label1.pack()
    int1.pack()
    label2.pack()
    int2.pack()
    label3.pack()
    int3.pack()
    button1.pack()
    button2.pack()

    frame1.pack()
    frame3.pack()
    frame2.pack()
    #root.geometry("300x200")
# Запуск главного цикла
    root.mainloop()

if __name__ == "__main__":
    main()