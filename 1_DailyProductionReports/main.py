import Format_1
import Format_2
import Format_3
import Format_4
from tkinter import *
from tkinter import ttk
from tkinter import messagebox


def process_Format_1():
    """Производит обработку файлов формата Format_1


    :return:
    """
    folder_to_process = tab1_process_folder.get()
    result_file = tab1_result_file.get()
    Format_1.main(folder_to_process, result_file)
    messagebox.showinfo('Рассчет окончен', 'Выполнение программы завершено!')


def process_Format_2():
    """Производит обработку файлов формата Format_2


    :return:
    """
    folder_to_process = tab2_process_folder.get()
    result_file = tab2_result_file.get()
    Format_2.main(folder_to_process, result_file)
    messagebox.showinfo('Рассчет окончен', 'Выполнение программы завершено!')


def process_Format_3():
    """Производит обработку файлов формата Format_3


    :return:
    """
    folder_to_process = tab3_process_folder.get()
    result_file = tab3_result_file.get()
    Format_3.main(folder_to_process, result_file)
    messagebox.showinfo('Рассчет окончен', 'Выполнение программы завершено!')


def process_Format_4():
    """Производит обработку файлов формата Format_4


    :return:
    """
    folder_to_process = tab4_process_folder.get()
    result_file = tab4_result_file.get()
    Format_4.main(folder_to_process, result_file)
    messagebox.showinfo('Рассчет окончен', 'Выполнение программы завершено!')


# Создание рабочего окна
window = Tk()
window.title("Обработка сводок")
window.geometry("600x250")

# Условные линии сетки
X_FIRST_LANE = 0
X_SECOND_LANE = 350
X_THIRD_LANE = 600


# Создание вкладок
tab_control = ttk.Notebook(window)
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)
tab3 = ttk.Frame(tab_control)
tab4 = ttk.Frame(tab_control)
tab_control.add(tab1, text='Format_1')
tab_control.add(tab2, text='Format_2')
tab_control.add(tab3, text='Format_3')
tab_control.add(tab4, text='Format_4')
tab_control.pack(expand=5, fill='both')


# Создание наполнения вкладок
# Вкладка 1
# Блок лейбл/инпут
tab1_label_one = Label(tab1, text="Вставьте ссылку на папку с архивами для обработки:", font=("Arial Bold", 12))
tab1_label_one.place(x=X_FIRST_LANE, y=30)
tab1_process_folder = Entry(tab1, width=90)
tab1_process_folder.place(x=X_FIRST_LANE, y=60)

# Блок лейбл/инпут
tab1_label_two = Label(tab1, text="Вставьте ссылку на файл для вставки значений:", font=("Arial Bold", 12))
tab1_label_two.place(x=X_FIRST_LANE, y=110)
tab1_result_file = Entry(tab1, width=90)
tab1_result_file.place(x=X_FIRST_LANE, y=140)

# Кнопка запуска
tab1_run_button = Button(tab1, text="Провести обработку", command=Format_1)
tab1_run_button.place(x=225, y=175)

# Вкладка 2
# Блок лейбл/инпут
tab2_label_one = Label(tab2, text="Вставьте ссылку на папку с архивами для обработки:", font=("Arial Bold", 12))
tab2_label_one.place(x=X_FIRST_LANE, y=30)
tab2_process_folder = Entry(tab2, width=90)
tab2_process_folder.place(x=X_FIRST_LANE, y=60)

# Блок лейбл/инпут
tab2_label_two = Label(tab2, text="Вставьте ссылку на файл для вставки значений:", font=("Arial Bold", 12))
tab2_label_two.place(x=X_FIRST_LANE, y=110)
tab2_result_file = Entry(tab2, width=90)
tab2_result_file.place(x=X_FIRST_LANE, y=140)

# Кнопка запуска
tab1_run_button = Button(tab2, text="Провести обработку", Format_2)
tab1_run_button.place(x=225, y=175)

# Вкладка 3
# Блок лейбл/инпут
tab3_label_one = Label(tab3, text="Вставьте ссылку на папку с архивами для обработки:", font=("Arial Bold", 12))
tab3_label_one.place(x=X_FIRST_LANE, y=30)
tab3_process_folder = Entry(tab3, width=90)
tab3_process_folder.place(x=X_FIRST_LANE, y=60)

# Блок лейбл/инпут
tab3_label_two = Label(tab3, text="Вставьте ссылку на файл для вставки значений:", font=("Arial Bold", 12))
tab3_label_two.place(x=X_FIRST_LANE, y=110)
tab3_result_file = Entry(tab3, width=90)
tab3_result_file.place(x=X_FIRST_LANE, y=140)

# Кнопка запуска
tab1_run_button = Button(tab3, text="Провести обработку", command=Format_3)
tab1_run_button.place(x=225, y=175)

# Вкладка 4
# Блок лейбл/инпут
tab4_label_one = Label(tab4, text="Вставьте ссылку на папку с архивами для обработки:", font=("Arial Bold", 12))
tab4_label_one.place(x=X_FIRST_LANE, y=30)
tab4_process_folder = Entry(tab4, width=90)
tab4_process_folder.place(x=X_FIRST_LANE, y=60)

# Блок лейбл/инпут
tab4_label_two = Label(tab4, text="Вставьте ссылку на файл для вставки значений:", font=("Arial Bold", 12))
tab4_label_two.place(x=X_FIRST_LANE, y=110)
tab4_result_file = Entry(tab4, width=90)
tab4_result_file.place(x=X_FIRST_LANE, y=140)

# Кнопка запуска
tab1_run_button = Button(tab4, text="Провести обработку", command=Format_4)
tab1_run_button.place(x=225, y=175)


# Вызов окна в работу
window.mainloop()
