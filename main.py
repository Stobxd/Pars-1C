import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import *
from data import vetsion
from buratino_mode import pars_buratino
from pars import solo_mode
#pyinstaller --onefile  --icon=1.ico --windowed main.py

# Функция для отображения нужного экрана
def show_screen(screen):
    # Скрываем все экраны
    main_screen.pack_forget()
    settings_screen.pack_forget()
    info_screen.pack_forget()
    # Показываем нужный экран
    screen.pack(fill="both", expand=True)

# Функция для выбора файла и вставки его пути в текстовое поле
def select_file():
    file_path = filedialog.askopenfilename()  # Открываем диалоговое окно для выбора файла
    if file_path:  # Если файл был выбран
        a.delete(0, tk.END)  # Очищаем текстовое поле
        a.insert(0, file_path)  # Вставляем путь в текстовое поле

def select_filejpg():
    file_path = filedialog.askopenfilename()  # Открываем диалоговое окно для выбора файла
    if file_path:  # Если файл был выбран
        imginp.delete(0, tk.END)  # Очищаем текстовое поле
        imginp.insert(0, file_path)  # Вставляем путь в текстовое поле
    
def Buratino():
    ggg = a.get()
    dateTable = s.get()
    nameinput = name.get()
    jjpgname = imginp.get()
    pars_buratino(ggg, dateTable,  nameinput, jjpgname)

def standart():
    ggg = a.get()
    dateTable = s.get()
    solo_mode(ggg, dateTable)

def printtt3():
    pass


# Создаем главное окно
root = tk.Tk()

root.title(vetsion)
root.geometry("400x400")
#root.iconbitmap("1.ico")

# Главный экран
main_screen = tk.Frame(root)
ttk.Label(main_screen, text="Главный экран", font=("Arial", 16)).pack(pady=10)




s = ttk.Entry(main_screen, width=50)
s.pack(pady=5)
s.insert(0, "Дату укажи (работает только в буратино)") 

name = ttk.Entry(main_screen, width=50)
name.pack(pady=5)
name.insert(0, "Укажи имя (работает только в буратино)") 



# Создаем текстовое поле отдельно и вызываем метод pack на отдельной строке
a = ttk.Entry(main_screen, width=50)
a.pack(pady=5)  # Добавляем текстовое поле на экран
a.insert(0, "Выбери название файла")  # Добавляем начальный текст

# Кнопка для выбора файла
b = ttk.Button(main_screen, text="Выбрать файл", command=select_file, width=30)
b.pack(pady=5)








# Создаем текстовое поле отдельно и вызываем метод pack на отдельной строке
imginp = ttk.Entry(main_screen, width=50) 
imginp.pack(pady=5)  # Добавляем текстовое поле на экран
imginp.insert(0, "Название фотки ")  # Добавляем начальный текст

# Кнопка для выбора файла
imgbutt = ttk.Button(main_screen, text="Выбрать файл jpg", command=select_filejpg, width=30)
imgbutt.pack(pady=5)




ttk.Button(main_screen, text="БУРАТИНО", command=lambda: Buratino(), width=30).pack(pady=10)
ttk.Button(main_screen, text="СТАНДАРТ", command=lambda: standart(), width=30).pack(pady=10)







# Кнопки для перехода на другие экраны
#tk.Button(main_screen, text="Перейти к настройкам", command=lambda: show_screen(settings_screen)).pack(pady=10)
#tk.Button(main_screen, text="Перейти к информации", command=lambda: show_screen(info_screen)).pack(pady=10)

# Экран настроек
settings_screen = tk.Frame(root)
tk.Label(settings_screen, text="Экран настроек", font=("Arial", 16)).pack(pady=10)
tk.Button(settings_screen, text="Назад", command=lambda: show_screen(main_screen)).pack(pady=10)
tk.Checkbutton(settings_screen, text="Включить опцию 1").pack(anchor="w", padx=20)
tk.Checkbutton(settings_screen, text="Включить опцию 2").pack(anchor="w", padx=20)

# Экран информации
info_screen = tk.Frame(root)
tk.Label(info_screen, text="Экран информации", font=("Arial", 16)).pack(pady=10)
tk.Label(info_screen, text="Здесь находится информация о приложении.").pack(pady=5)
tk.Button(info_screen, text="Назад", command=lambda: show_screen(main_screen)).pack(pady=10)

# Запускаем с главного экрана
show_screen(main_screen)

# Запуск основного цикла приложения
root.mainloop()
