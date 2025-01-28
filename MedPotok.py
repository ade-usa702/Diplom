#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import matplotlib
matplotlib.use('TkAgg')

import re
import tkinter as tk
from tkinter import PhotoImage, filedialog, messagebox, ttk

import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from openpyxl.chart import PieChart, Reference

selected_dataA = None
selected_data = None


class ExcelAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("MedPotok")
        # Устанавливаем фиксированный размер окна
        self.root.resizable(False, False)

        try:
            # Load the icon image
            self.icon = PhotoImage(file="C:\data\icon.png")
            self.root.iconphoto(False, self.icon)
        except Exception as e:
            print("Error loading icon image:", e)
        
        self.file_loaded = False  # Переменная для отслеживания загрузки файла
        self.file_path = None
        
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack_propagate(False)
        self.main_frame.pack(padx=20, pady=20)
        
        # Ширина всех кнопок
        button_width = 40   
        
        self.load_button = tk.Button(self.main_frame, text="Загрузить файл",width=button_width, command=self.load_file)
        self.load_button.grid(row=0, column=0, pady=10, ipady=10)
        

        
        self.stat_buttons = []
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика госпитализаций", width=button_width, command=self.show_hospitalization_stats, state=tk.DISABLED))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика переводов", width=button_width, command=self.show_transfer_stats, state=tk.DISABLED))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика амбулаторных поступлений",width=button_width, command=self.show_ambulatory_stats, state=tk.DISABLED))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Общий аналитический отчет",width=button_width, command=self.show_general_report, state=tk.DISABLED))

        for idx, button in enumerate(self.stat_buttons):
            button.grid(row=idx+1, column=0, pady=(10,0), ipady=10)

        self.return_button = tk.Button(self.main_frame, text="Возврат", width=button_width, command=self.clear_data, state=tk.DISABLED)
        self.return_button.grid(row=len(self.stat_buttons)+1, column=0, pady=10, ipady=10)
        
        self.exit_button = tk.Button(self.main_frame, text="Выход", width=button_width, command=self.root.destroy)
        self.exit_button.grid(row=len(self.stat_buttons)+2, column=0, pady=10, ipady=10)
        
        self.data = None
        self.department_data = {}  # Словарь для хранения данных о фреймах в зависимости от названия отделения
        
                # Рассчитываем размер окна
        self.root.update_idletasks()  # Обновляем отображение окна, чтобы получить актуальные размеры
        window_width = 450  # Ширина основного фрейма + запас
        window_height = 650  # Высота основного фрейма + запас

        # Рассчитываем координаты для центрирования окна в самом верху экрана
        screen_width = root.winfo_screenwidth()
        x = (screen_width - window_width) // 2
        y = 0

        # Устанавливаем геометрию окна
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")



    def reinit_main_frame(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(padx=20, pady=20)

        # Ширина всех кнопок
        button_width = 40   
        self.load_button = tk.Button(self.main_frame, text="Загрузить файл",width=button_width, command=self.load_file,state=tk.DISABLED)
        self.load_button.grid(row=0, column=0, pady=10, ipady=10)
        

        
        self.stat_buttons = []
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика госпитализаций", width=button_width, command=self.show_hospitalization_stats, state=tk.NORMAL))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика переводов", width=button_width, command=self.show_transfer_stats, state=tk.NORMAL))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Статистика амбулаторных поступлений",width=button_width, command=self.show_ambulatory_stats, state=tk.NORMAL))
        self.stat_buttons.append(tk.Button(self.main_frame, text="Общий аналитический отчет",width=button_width, command=self.show_general_report, state=tk.NORMAL))

        for idx, button in enumerate(self.stat_buttons):
            button.grid(row=idx+1, column=0, pady=(10,0), ipady=10)

        self.return_button = tk.Button(self.main_frame, text="Возврат", width=button_width, command=self.clear_data, state=tk.NORMAL)
        self.return_button.grid(row=len(self.stat_buttons)+1, column=0, pady=10, ipady=10)
        
        self.exit_button = tk.Button(self.main_frame, text="Выход", width=button_width, command=self.root.destroy)
        self.exit_button.grid(row=len(self.stat_buttons)+2, column=0, pady=10, ipady=10)

        
        
    def show_hospitalization_stats(self):
            self.clear_frame()
            if self.data is not None:
                # Create a histogram
                sns.set()  # Set seaborn defaults
                sns.set(style="whitegrid")  # Set seaborn style
                plt.figure(figsize=(6, 4))  # Set figure size
                # Группировка данных по столбцу "Месяц" и вычисление суммы значений столбца "Госп. ВСЕГО"
                monthly_hospitalizations = self.data.groupby('Месяц')['Госп. ВСЕГО'].sum().reset_index()

                # Ширина всех кнопок
                button_width = 25 
        
                # Создание гистограммы
                plt.figure(figsize=(6, 4))
                sns.barplot(x='Месяц', y='Госп. ВСЕГО', data=monthly_hospitalizations)
                plt.rc('axes', labelsize=10)
                plt.rc('axes', titlesize=10)
                plt.xlabel('Месяц')
                plt.ylabel('Количество пациентов')
                plt.title('Статистика госпитализаций пациентов по месяцам') 
                plt.tight_layout()


                # Создаем словарь с данными для каждого отделения
                departments = self.data["Наименование отделения"].unique()
                self.department_data = {}
                self.OtdelData = {}
                for i in departments:
                    self.department_data[i] = self.data[self.data["Наименование отделения"] == i]
                    self.OtdelData[i] = self.department_data[i].groupby('Месяц')['Госп. ВСЕГО'].sum()
                    
                # кнопка "Возврат"
                self.return_button2 = tk.Button(self.root, text="Возврат",width=25, command=self.back)
                self.return_button2.place(relx=.5, rely=.5, anchor="s")
                self.return_button2.pack(pady=(10, 0))                    

                # кнопка "Сохранить диаграмму"
                self.save_plot_button2 = tk.Button(self.root, text="Сохранить диаграмму",width=button_width, command=self.save_plot)
                self.save_plot_button2.place(relx=.5, rely=.3, anchor="s")
                self.save_plot_button2.pack(pady=5)

                # кнопка "Сохранить в текстовом файле"
                self.save_to_text_button2 = tk.Button(self.root, text="Сохранить в текстовом файле",width=button_width, command=self.save_to_text)
                self.save_to_text_button2.place(relx=.5, rely=.1, anchor="s")
                self.save_to_text_button2.pack(pady=5)
                
                # Кнопка "Сохранить в Excel"
                self.save_to_excel_buttonHosp = tk.Button(self.root, text="Сохранить в Excel", width=25,  command=self.save_to_excelHosp)
                self.save_to_excel_buttonHosp.pack(pady=5)
                
                # Создаем фрейм для выбора периода
                self.period_frame = tk.Frame(self.root)
                self.period_frame.pack(pady=5)

                # Создаем метку и комбо-бокс внутри фрейма
                self.period_label = tk.Label(self.period_frame, text="Выберите период:")
                self.period_label.grid(row=0, column=0, padx=(40, 0))

                periods = ["День", "Неделя", "Месяц", "3 месяца", "6 месяцев"]
                self.period_combo = ttk.Combobox(self.period_frame, values=periods, width=10)
                self.period_combo.grid(row=0, column=1, padx=(0, 45))
                self.period_combo.bind("<<ComboboxSelected>>", self.show_period_widgetsHosp)
                

                # Создаем дополнительные элементы для выбора отделения
                self.department_label = tk.Label(self.root, text="Выберите отделение:")
                self.department_label.pack()

                self.selected_department = tk.StringVar(self.root)
                self.selected_department.set(departments[0])

                self.department_menu = tk.OptionMenu(self.root, self.selected_department, *departments)
                self.department_menu.pack()

                self.plot_button = tk.Button(self.root, text="Применить", width=button_width, command=self.plot_hospitalization_stats)
                self.plot_button.pack()

                # Create a Tkinter canvas containing the histogram plot
                self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                self.histogram_canvas.draw()
                self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка", "Сначала загрузите данные.")

    def back(self):
        self.clear_frame()
        self.reinit_main_frame()
                
    def show_period_widgetsHosp(self, event):
        self.selected_period = self.period_combo.get()
        # Удаляем предыдущие виджеты периода, если они есть
        for widget in self.period_frame.winfo_children():
            widget.destroy()

        if self.selected_period == "День":
            # Создаем виджеты для ввода дня и месяца
            self.week_label = tk.Label(self.period_frame, text="Введите день и месяц:")
            self.week_label.grid(row=1, column=0)
            self.day_entry = tk.Entry(self.period_frame, width=5)
            self.day_entry.grid(row=1, column=1)
            self.month_entry = tk.Entry(self.period_frame, width=5)
            self.month_entry.grid(row=1, column=2)
            max_month = max(self.data["Месяц"])
            self.max_month_label = tk.Label(self.period_frame, text=f"Количество месяцев в данном файле {max_month}")
            self.max_month_label.grid(row=2, column=0, columnspan=3, pady=5, padx=35)            
        elif self.selected_period == "Неделя":
            # Создаем виджет для ввода недели
            self.week_label = tk.Label(self.period_frame, text="Введите неделю:")
            self.week_label.grid(row=1, column=0)
            self.week_entry = tk.Entry(self.period_frame, width=10)
            self.week_entry.grid(row=1, column=1)
            # Display max week label
            max_week = max(self.data["Неделя"])
            self.max_week_label = tk.Label(self.period_frame, text=f"Номера недель в данном файле от 1 до {max_week}")
            self.max_week_label.grid(row=2, column=0, columnspan=2, pady=5, padx=35)
        elif self.selected_period == "Месяц":
            # Создаем виджет для ввода месяца
            self.week_label = tk.Label(self.period_frame, text="Введите месяц:")
            self.week_label.grid(row=1, column=0)
            self.month_entry = tk.Entry(self.period_frame, width=10)
            self.month_entry.grid(row=1, column=1)
            max_month = max(self.data["Месяц"])
            self.max_month_label = tk.Label(self.period_frame, text=f"Количество месяцев в данном файле {max_month}")
            self.max_month_label.grid(row=2, column=0, columnspan=2, pady=5, padx=35)                
        elif self.selected_period == "3 месяца":
            # Создаем виджеты для ввода начала и конца месяца
            self.week_label = tk.Label(self.period_frame, text="Введите от и до месяц:")
            self.week_label.grid(row=1, column=0)
            self.start_month_entry = tk.Entry(self.period_frame, width=5)
            self.start_month_entry.grid(row=1, column=1)
            self.end_month_entry = tk.Entry(self.period_frame, width=5)
            self.end_month_entry.grid(row=1, column=2)
            max_month = max(self.data["Месяц"])
            self.max_month_label = tk.Label(self.period_frame, text=f"Количество месяцев в данном файле {max_month}")
            self.max_month_label.grid(row=2, column=0, columnspan=3, pady=5, padx=35)                
        # Выводим кнопку "Отменить" для всех случаев
        self.cancel_button = tk.Button(self.period_frame, text="Отменить", command=self.cancel_periodHosp)
        self.cancel_button.grid(row=1, column=3, padx=2)
          
    def cancel_periodHosp(self):
        # Сбрасываем выбранное значение периода на None
        self.selected_period = None
        # Уничтожаем таблицу и связанную с ней полосу прокрутки
        if hasattr(self, 'table'):
            self.table.destroy()
        if hasattr(self, 'hscrollbar'):
            self.hscrollbar.destroy()
        if hasattr(self, 'histogram_canvas'):
            self.histogram_canvas.get_tk_widget().destroy()
        # Уничтожаем фрейм выбора периода и пересоздаем виджеты
        self.period_frame.destroy()
        self.period_frame = tk.Frame(self.root)
        self.period_frame.pack(pady=5)
        self.period_label = tk.Label(self.period_frame, text="Выберите период:")
        self.period_label.grid(row=0, column=0, padx=(43, 0))

        periods = ["День", "Неделя", "Месяц", "3 месяца", "6 месяцев"]
        self.period_combo = ttk.Combobox(self.period_frame, values=periods, width=10)
        self.period_combo.grid(row=0, column=1, padx=(0, 45))
        self.period_combo.bind("<<ComboboxSelected>>", self.show_period_widgetsHosp)
        
    def plot_hospitalization_stats(self):
        plt.clf()  # Удаляем предыдущий график
        # Проверка, создан ли Combobox
        if self.period_combo is None:
            messagebox.showerror("Ошибка", "Сначала выберите период.")
            return
        # Получение выбранного периода
        selected_period = self.selected_period
        if selected_period == "День":
            self.plot_day_statsHosp()
        elif selected_period == "Неделя":
            self.plot_week_statsHosp()
        elif selected_period == "Месяц":
            self.plot_month_statsHosp()  
        elif selected_period == "3 месяца":
            self.plot_month3_statsHosp()
        elif selected_period == "6 месяцев":
            self.plot_month6_statsHosp() 
 

    def plot_day_statsHosp(self):
        self.save_plot_button2.config(state="disabled")
        self.save_to_excel_buttonHosp.config(state="disabled")
        self.histogram_canvas.get_tk_widget().destroy()
        # Получение выбранного пользователем месяца и дня
        selected_month = self.month_entry.get()
        selected_day = self.day_entry.get()

        # Преобразование выбранных месяцев в числовой формат 
        selected_month = int(selected_month)
        selected_day = int(selected_day)

        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()

        # Предположим, что self.data содержит DataFrame с данными
        # Фильтрация данных по выбранному месяцу, дню и отделению
        filtered_data = self.data[(self.data["Месяц"] == selected_month) & 
                                  (self.data["День"] == selected_day) & 
                                  (self.data["Наименование отделения"] == selected_department)]
        
        # Если нет данных для выбранного дня и месяца, вывести ошибку
        if filtered_data.empty:
            messagebox.showerror("Ошибка", "Такого дня/месяца нет. Пожалуйста проверьте данные.")
            return

        # Создание текстового виджета для отображения таблицы
        columns = list(filtered_data.columns)
        self.table = ttk.Treeview(self.root, columns=columns, show="headings", height=10)

        # Установка ширины столбцов
        for col in columns:
            self.table.column(col, width=90)  # Set the width as desired

        # Создание горизонтальной полосы прокрутки и сохранение ссылки на неё
        self.hscrollbar = ttk.Scrollbar(self.root, orient="horizontal", command=self.table.xview)
        self.hscrollbar.pack(side="bottom", fill="x")

        # Установка прокрутки для таблицы
        self.table.configure(xscrollcommand=self.hscrollbar.set)

        # Вставка заголовков столбцов
        for col in columns:
            self.table.heading(col, text=col)

        # Вставка данных в таблицу
        for index, row in filtered_data.iterrows():
            self.table.insert("", "end", values=row.tolist())

        # Упаковка таблицы
        self.table.pack(side="left", fill="both", expand=True)

        

    def plot_week_statsHosp(self):
        self.save_plot_button2.config(state=tk.NORMAL)
        self.save_to_excel_buttonHosp.config(state="disabled")

        # Получение выбранного пользователем номера недели
        selected_week = self.week_entry.get()
        selected_week = int(selected_week)

        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()

        # Фильтрация данных по выбранной неделе и отделению
        filtered_data = self.data[(self.data["Неделя"] == selected_week) & 
                                  (self.data["Наименование отделения"] == selected_department)]

        # Получение данных из столбца "Госп. ВСЕГО" по выбранной неделе и отделению
        selected_data = filtered_data["Госп. ВСЕГО"]
        
        if selected_department and selected_week is not None:
            if not selected_data.empty:
                # Проверка, все ли значения равны нулю       
                if selected_data.eq(0).all():
                    messagebox.showinfo("Информация", f"В отделении {selected_department} на неделе {selected_week} не было случаев.")
                else:
                    # Построение гистограммы
                    plt.figure(figsize=(6, 4))
                    # Построение гистограммы с помощью функции hist(), задав количество корзин (bins) как длину столбца данных
                    plt.hist(range(1, len(selected_data) + 1), weights=selected_data, bins=len(selected_data))
                    plt.rc('axes', labelsize= 10 )
                    plt.rc('axes', titlesize= 10 )

                    plt.xlabel(f'Дни {selected_week} недели')
                    plt.ylabel('Количество госпитализированных пациентов')
                    plt.title(f'Гистограмма госп. в отделении {selected_department} на неделе {selected_week}')
                    plt.tight_layout()

                    # Создаем данные для линейной регрессии
                    x = np.arange(1, len(selected_data) + 1)
                    y = np.array(selected_data)

                    # Вычисляем коэффициенты линейной регрессии
                    coefficients = np.polyfit(x, y, 1)
                    polynomial = np.poly1d(coefficients)
                    linear_regression_line = polynomial(x)

                    # Добавляем линейную регрессию
                    plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                             label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                    plt.legend()

                    # Update the Tkinter canvas containing the histogram plot
                    self.histogram_canvas.get_tk_widget().destroy()
                    self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                    self.histogram_canvas.draw()
                    self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка", f"Нет данных для недели {selected_week} в отделении {selected_department}.")  
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или неделю лечения.")

    def plot_month_statsHosp(self):
        self.save_plot_button2.config(state=tk.NORMAL)
        self.save_to_excel_buttonHosp.config(state=tk.NORMAL)
        # Получение выбранного пользователем месяца
        selected_month = self.month_entry.get()
        selected_month = int(selected_month)
        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()
        # Фильтрация данных по выбранной неделе и отделению
        filtered_data = self.data[(self.data["Месяц"] == selected_month) &
                                  (self.data["Наименование отделения"] == selected_department)]
        # Получение данных из столбца "Госп. ВСЕГО" по выбранной неделе и отделению
        selected_data = filtered_data["Госп. ВСЕГО"]
        # Разбиваем данные на группы по 7 значений и суммируем каждую группу
        grouped_data = [selected_data[i:i + 7].sum() for i in range(0, len(selected_data), 7)]

        if selected_department and selected_month is not None:
            if not selected_data.empty:
                # Проверка, все ли значения равны нулю
                if selected_data.eq(0).all():
                    messagebox.showinfo("Информация",
                                        f"В отделении {selected_department} на неделе {selected_month} не было случаев.")
                else:
                    # Построение гистограммы
                    plt.figure(figsize=(6, 4))
                    # Построение гистограммы с помощью функции bar(), задав количество корзин (bins) как длину столбца данных
                    plt.bar(range(1, len(grouped_data) + 1), grouped_data)  # Группы на оси x, сумма значений на оси y
                    plt.rc('axes', labelsize=10)
                    plt.rc('axes', titlesize=10)
                    plt.xlabel(f'Недели {selected_month} месяца')
                    plt.ylabel('Количество госпитализированных пациентов')
                    plt.title(
                        f'Гистограмма госп. пациентов в отделении {selected_department} в месяце {selected_month}')
                    plt.tight_layout()

                    # Создаем данные для линейной регрессии
                    x = np.arange(1, len(grouped_data) + 1)
                    y = np.array(grouped_data)

                    # Вычисляем коэффициенты линейной регрессии
                    coefficients = np.polyfit(x, y, 1)
                    polynomial = np.poly1d(coefficients)
                    linear_regression_line = polynomial(x)

                    # Добавляем линейную регрессию
                    plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                             label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                    plt.legend()

                    # Update the Tkinter canvas containing the histogram plot
                    self.histogram_canvas.get_tk_widget().destroy()
                    self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                    self.histogram_canvas.draw()
                    self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка",
                                     f"Нет данных для месяца {selected_month} в отделении {selected_department}.")
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или месяц лечения.")

    def plot_month3_statsHosp(self):
        self.save_plot_button2.config(state=tk.NORMAL)
        self.save_to_excel_buttonHosp.config(state=tk.NORMAL)
        # Получение выбранного пользователем начала-месяц
        selected_start_month = self.start_month_entry.get()
        # Получение выбранного пользователем конца-месяц
        selected_end_month = self.end_month_entry.get()
        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()
        # Преобразование выбранных месяцев в числовой формат 
        selected_start_month = int(selected_start_month)
        selected_end_month = int(selected_end_month)

        # Фильтрация данных по выбранным месяцам и отделению
        filtered_data = self.data[(self.data["Месяц"] >= selected_start_month) &
                                  (self.data["Месяц"] <= selected_end_month) &
                                  (self.data["Наименование отделения"] == selected_department)]
        # Получение данных из столбца "Госп. ВСЕГО" по выбранной неделе и отделению
        grouped_data = filtered_data.groupby("Месяц")["Госп. ВСЕГО"].sum()

        # Действия с данными - построение графика
        if selected_department and selected_start_month and selected_end_month is not None:
            if not grouped_data.empty:
                # Построение гистограммы
                plt.figure(figsize=(6, 4))
                plt.bar(grouped_data.index, grouped_data.values, width=0.4)
                plt.xticks(
                    grouped_data.index)  # Устанавливаем метки по оси x на каждый месяц от selected_start_month до selected_end_month
                plt.rc('axes', labelsize=10)
                plt.rc('axes', titlesize=10)
                plt.xlabel('Месяц')
                plt.ylabel('Количество госпитализированных пациентов')
                plt.title(f'Статистика госп. пациентов по указанным месяцам в отделении {selected_department}')
                plt.tight_layout()

                # Создаем данные для линейной регрессии
                x = np.arange(len(grouped_data))
                y = grouped_data.values

                # Вычисляем коэффициенты линейной регрессии
                coefficients = np.polyfit(x, y, 1)
                polynomial = np.poly1d(coefficients)
                linear_regression_line = polynomial(x)

                # Добавляем линейную регрессию
                plt.plot(grouped_data.index, linear_regression_line, color='red', linestyle='-', linewidth=2,
                         label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                plt.legend()

                # Update the Tkinter canvas containing the histogram plot
                self.histogram_canvas.get_tk_widget().destroy()
                self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                self.histogram_canvas.draw()
                self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка",
                                     f"Нет данных для выбранного периода с {selected_start_month} по {selected_end_month} в отделении {selected_department}.")
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или период лечения.")

    def plot_month6_statsHosp(self):
        self.save_plot_button2.config(state=tk.NORMAL)
        self.save_to_excel_buttonHosp.config(state=tk.NORMAL)
        selected_i = self.selected_department.get()
        selected_data = self.OtdelData.get(selected_i)

        if selected_data is not None and not selected_data.empty:
            plt.figure(figsize=(6, 4))  # Set figure size
            sns.barplot(x=selected_data.index, y=selected_data.values)
            plt.rc('axes', labelsize=10)
            plt.rc('axes', titlesize=10)
            plt.xlabel('Месяц')
            plt.ylabel('Количество госпитализированных пациентов')
            plt.title(f'Статистика госп. пациентов по месяцам в отделении {selected_i}')
            plt.tight_layout()

            # Создаем данные для линейной регрессии
            x = np.arange(len(selected_data))
            y = np.array(selected_data.values)

            # Вычисляем коэффициенты линейной регрессии
            coefficients = np.polyfit(x, y, 1)
            polynomial = np.poly1d(coefficients)
            linear_regression_line = polynomial(x)

            # Добавляем линейную регрессию
            plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                     label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
            plt.legend()

            # Настраиваем метки на оси x, чтобы соответствовать индексам месяцев
            plt.xticks(ticks=np.arange(1, len(selected_data) + 1), labels=selected_data.index)

            # Update the Tkinter canvas containing the histogram plot
            self.histogram_canvas.get_tk_widget().destroy()
            self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
            self.histogram_canvas.draw()
            self.histogram_canvas.get_tk_widget().pack()
        else:
            messagebox.showerror("Ошибка", "Выберите отделение.")



    def save_plot(self):
            # Сохранить диаграмму в файл
            file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
            if file_path:
                plt.savefig(file_path)

    def save_to_text(self):
            # Сохранить данные в текстовый файл
            selected_i = self.selected_department.get()
            selected_data = self.OtdelData.get(selected_i)
            if selected_data is not None:
                file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
                if file_path:
                    with open(file_path, "w") as file:
                        file.write(f'Статистика госпитализаций пациентов по месяцам в отделении {selected_i}\n')
                        file.write(selected_data.to_string())
                        file.write("\nГде: 1 - Январь, 2 - Февраль, 3 - Март, 4 - Апрель, 5 - Май, 6 - Июнь")
            else:
                messagebox.showerror("Ошибка", "Выберите отделение.")    
            pass  
        
    def save_to_excelHosp(self):
        #Сохранить диаграмму в Эксель
        # Выбираем путь сохранения файла
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            selected_i = self.selected_department.get()
            selected_data = self.OtdelData.get(selected_i)
            if selected_data is not None:
                    # Создаем DataFrame из данных
                    df = pd.DataFrame(selected_data)
                    df.reset_index(inplace=True)
                    df.columns = ['Месяц', 'Количество пациентов']

                    # Создаем новый Excel файл и записываем в него данные
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Госпитализация')

                        # Добавляем график в Excel файл
                        worksheet = writer.sheets['Госпитализация']
                        chart = writer.book.add_chart({'type': 'line'})

                        # Добавляем данные для графика
                        chart.add_series({
                            'categories': ['Госпитализация', 1, 0, len(df), 0],
                            'values': ['Госпитализация', 1, 1, len(df), 1],
                            'name': ['Госпитализация', 0, 1],
                        })

                        # Устанавливаем заголовок и метки осей
                        chart.set_title({'name': 'Статистика госпитализированных пациентов по месяцам'})
                        chart.set_x_axis({'name': 'Месяц'})
                        chart.set_y_axis({'name': 'Количество госпитализированных пациентов'})

                        # Вставляем график в лист Excel
                        worksheet.insert_chart('D2', chart)

                    messagebox.showinfo("Успешно", "Данные и график сохранены в файл.")
            else:
                    messagebox.showerror("Ошибка", "Выберите отделение.")  
                    
                    

    def show_transfer_stats(self):
        self.clear_frame()
        if self.data is not None:
            # Отображение статистики для столбца 'с/п амб.'
            # Преобразование столбца 'с/п амб.' в числовой формат
            self.data['с/п амб.'] = pd.to_numeric(self.data['с/п амб.'], errors='coerce')
            # Группировка данных по столбцу "Месяц" и вычисление суммы значений столбца "с/п амб."
            stat_perevod_amb = self.data.groupby('Месяц')['с/п амб.'].sum().reset_index()

            button_width = 40
            button_height = 2 
            
            # Создание круговой диаграммы для 'с/п амб.'
            fig, axes = plt.subplots(1, 2, figsize=(6, 3))  # Create a figure with two subplots
            plt.rc('axes', labelsize= 10 )
            plt.rc('axes', titlesize= 10 )
            # Построение первой диаграммы (с/п амб.)
            axes[0].pie(stat_perevod_amb['с/п амб.'], labels=stat_perevod_amb['Месяц'], autopct='%1.1f%%')
            axes[0].set_title('Статистика перевод. амб. и')

            # Отображение статистики для столбца 'с/п госп.'
            # Преобразование столбца 'с/п госп.' в числовой формат
            self.data['с/п госп.'] = pd.to_numeric(self.data['с/п госп.'], errors='coerce')
            # Группировка данных по столбцу "Месяц" и вычисление суммы значений столбца "с/п госп."
            stat_perevod_gosp = self.data.groupby('Месяц')['с/п госп.'].sum().reset_index()
            plt.rc('axes', labelsize= 10 )
            plt.rc('axes', titlesize= 10 )
            # Создание круговой диаграммы для 'с/п госп.'
            axes[1].pie(stat_perevod_gosp['с/п госп.'], labels=stat_perevod_gosp['Месяц'], autopct='%1.1f%%')
            axes[1].set_title('статистика перевод. госп. по месяцам')

            # Create a Tkinter canvas containing the Matplotlib figure
            self.histogram_canvas = FigureCanvasTkAgg(fig, master=self.root)
            self.histogram_canvas.draw()
            self.histogram_canvas.get_tk_widget().pack()

        else:
            messagebox.showerror("Ошибка", "Сначала загрузите данные.")
           
         # кнопка "Возврат"
        self.return_button3 = tk.Button(self.root, text="Возврат",width=button_width, height = button_height, command=self.back)
        self.return_button3.place(relx=.5, rely=.5, anchor="s")
        self.return_button3.pack(pady=10)

        # кнопка "Сохранить диаграмму"
        self.save_plot_button3 = tk.Button(self.root, text="Сохранить диаграмму",width=button_width,height = button_height, command=self.save_plot2)
        self.save_plot_button3.place(relx=.5, rely=.3, anchor="s")
        self.save_plot_button3.pack(pady=10)

        # кнопка "Сохранить в текстовом файле"
        self.save_to_text_button3 = tk.Button(self.root, text="Сохранить в текстовом файле",width=button_width,height = button_height, command=self.save_to_text2)
        self.save_to_text_button3.place(relx=.5, rely=.1, anchor="s")
        self.save_to_text_button3.pack(pady=10)
        
        # Кнопка "Сохранить в Excel"
        self.save_to_excel_buttonTr = tk.Button(self.root, text="Сохранить в Excel", width=button_width,height = button_height,  command=lambda: self.save_to_excelTr(stat_perevod_amb, stat_perevod_gosp))
        self.save_to_excel_buttonTr.pack(pady=10)


        

    def save_to_excelTr(self, stat_perevod_amb, stat_perevod_gosp):
        if self.data is not None:
            # Create a file dialog to get the file name and location from the user
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

            if file_path:
                # Create a new Excel workbook
                workbook = openpyxl.Workbook()

                # Create a new Excel writer object
                excel_writer = pd.ExcelWriter(file_path, engine='openpyxl')
                excel_writer.book = workbook

                # Write each DataFrame to a specific sheet
                stat_perevod_amb.to_excel(excel_writer, sheet_name='Ambulatory Transfer Stats', index=False)
                stat_perevod_gosp.to_excel(excel_writer, sheet_name='Hospital Transfer Stats', index=False)

                # Create pie charts
                amb_chart = PieChart()
                amb_chart.title = "Ambulatory Transfer Stats"
                amb_chart_data = Reference(worksheet=workbook['Ambulatory Transfer Stats'], min_col=2, min_row=1, max_row=stat_perevod_amb.shape[0], max_col=2)
                amb_chart.add_data(amb_chart_data, titles_from_data=True)
                amb_chart.set_categories(Reference(worksheet=workbook['Ambulatory Transfer Stats'], min_col=1, min_row=2, max_row=stat_perevod_amb.shape[0]))
                workbook['Ambulatory Transfer Stats'].add_chart(amb_chart, "D2")

                gosp_chart = PieChart()
                gosp_chart.title = "Hospital Transfer Stats"
                gosp_chart_data = Reference(worksheet=workbook['Hospital Transfer Stats'], min_col=2, min_row=1, max_row=stat_perevod_gosp.shape[0], max_col=2)
                gosp_chart.add_data(gosp_chart_data, titles_from_data=True)
                gosp_chart.set_categories(Reference(worksheet=workbook['Hospital Transfer Stats'], min_col=1, min_row=2, max_row=stat_perevod_gosp.shape[0]))
                workbook['Hospital Transfer Stats'].add_chart(gosp_chart, "D2")

                # Save the Excel file
                excel_writer.save()

                messagebox.showinfo("Успех", "Данные и графики успешно сохранены в Excel файл.")
        else:
            messagebox.showerror("Ошибка", "Сначала загрузите данные.")

        
    def save_plot2(self):
                # Сохранить диаграмму в файл
                file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
                if file_path:
                    plt.savefig(file_path)

    def save_to_text2(self):
                # Сохранить данные в текстовый файл
                if self.data is not None:
                    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
                    if file_path:
                        with open(file_path, "w") as file:
                            file.write(self.data.to_string())


    def show_ambulatory_stats(self):
            # Здесь логика для вывода статистики амбулаторных поступлений
            self.clear_frame()
            if self.data is not None:
                # Create a histogram
                sns.set()  # Set seaborn defaults
                sns.set(style="whitegrid")  # Set seaborn style
                plt.figure(figsize=(6, 4))  # Set figure size
                # Группировка данных по столбцу "Месяц" и вычисление суммы значений столбца "Госп. ВСЕГО"
                monthly_amb = self.data.groupby('Месяц')['Амб. ВСЕГО'].sum().reset_index()

                # Создание гистограммы
                plt.figure(figsize=(6, 4))
                sns.barplot(x='Месяц', y='Амб. ВСЕГО', data=monthly_amb)
                plt.rc('axes', labelsize= 10 )
                plt.rc('axes', titlesize= 10 )
                plt.xlabel('Месяц')
                plt.ylabel('Количество пациентов')
                plt.title('Статистика пациентов амбулаторно по месяцам') 
                plt.tight_layout()

                # Создаем словарь с данными для каждого отделения
                departments = self.data["Наименование отделения"].unique()
                self.department_dataA = {}
                self.OtdelDataA = {}
                for i in departments:
                    self.department_dataA[i] = self.data[self.data["Наименование отделения"] == i]
                    self.OtdelDataA[i] = self.department_dataA[i].groupby('Месяц')['Амб. ВСЕГО'].sum()
                    
                # кнопка "Возврат"
                self.return_button4 = tk.Button(self.root, text="Возврат",width=25, command=self.back)
                self.return_button4.place(relx=.5, rely=.5, anchor="s")
                self.return_button4.pack(pady=(10, 0))

                # кнопка "Сохранить диаграмму"
                self.save_plot_button4 = tk.Button(self.root, text="Сохранить диаграмму", width=25,command=self.save_plot4)
                self.save_plot_button4.place(relx=.5, rely=.3, anchor="s")
                self.save_plot_button4.pack(pady=5)

                # кнопка "Сохранить в текстовом файле"
                self.save_to_text_button4 = tk.Button(self.root, text="Сохранить в текстовом файле",width=25, command=self.save_to_text4)
                self.save_to_text_button4.place(relx=.5, rely=.1, anchor="s")
                self.save_to_text_button4.pack(pady=5)
                
                # Кнопка "Сохранить в Excel"
                self.save_to_excel_button = tk.Button(self.root, text="Сохранить в Excel", width=25,  command=self.save_to_excel)
                self.save_to_excel_button.pack(pady=5)
                
                # Создаем фрейм для выбора периода
                self.period_frame = tk.Frame(self.root)
                self.period_frame.pack(pady=5)

                # Создаем метку и комбо-бокс внутри фрейма
                self.period_label = tk.Label(self.period_frame, text="Выберите период:")
                self.period_label.grid(row=0, column=0, padx=(43, 0))

                periods = ["День", "Неделя", "Месяц", "3 месяца", "6 месяцев"]
                self.period_combo = ttk.Combobox(self.period_frame, values=periods, width=10)
                self.period_combo.grid(row=0, column=1, padx=(0, 45))
                self.period_combo.bind("<<ComboboxSelected>>", self.show_period_widgets)

                
                # Создаем дополнительные элементы для выбора отделения
                department_frame = tk.Frame(self.root)
                department_frame.pack(pady=5, padx=5, fill='both')

                self.department_label = tk.Label(department_frame, text="Выберите отделение:")
                self.department_label.grid(row=0, column=0, padx=(97,0))
                
                self.selected_department = tk.StringVar(self.root)
                self.selected_department.set(departments[0])

                self.department_menu = tk.OptionMenu(department_frame, self.selected_department, *departments)
                self.department_menu.grid(row=0, column=1, padx=(0,45))

                self.plot_button = tk.Button(self.root, text="Применить", width=25, command=self.plot_amb_stats)
                self.plot_button.pack(pady=5)

                # Create a Tkinter canvas containing the histogram plot
                self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                self.histogram_canvas.draw()
                self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка", "Сначала загрузите данные.")
                
    def show_period_widgets(self, event):
        self.selected_period = self.period_combo.get()
        # Удаляем предыдущие виджеты периода, если они есть
        for widget in self.period_frame.winfo_children():
            widget.destroy()

        if self.selected_period == "День":
            # Создаем виджеты для ввода дня и месяца
            self.week_label = tk.Label(self.period_frame, text="Введите день и месяц:")
            self.week_label.grid(row=1, column=0)
            self.day_entry = tk.Entry(self.period_frame, width=5)
            self.day_entry.grid(row=1, column=1)
            self.month_entry = tk.Entry(self.period_frame, width=5)
            self.month_entry.grid(row=1, column=2)
        elif self.selected_period == "Неделя":
            # Создаем виджет для ввода недели
            self.week_label = tk.Label(self.period_frame, text="Введите неделю:")
            self.week_label.grid(row=1, column=0)
            self.week_entry = tk.Entry(self.period_frame, width=10)
            self.week_entry.grid(row=1, column=1)
            # Display max week label
            max_week = max(self.data["Неделя"])
            self.max_week_label = tk.Label(self.period_frame, text=f"Номера недель в данном файле от 1 до {max_week}")
            self.max_week_label.grid(row=2, column=0, columnspan=2, pady=5, padx=35)
        elif self.selected_period == "Месяц":
            # Создаем виджет для ввода месяца
            self.week_label = tk.Label(self.period_frame, text="Введите месяц:")
            self.week_label.grid(row=1, column=0)
            self.month2_entry = tk.Entry(self.period_frame, width=10)
            self.month2_entry.grid(row=1, column=1)
            max_month = max(self.data["Месяц"])
            self.max_month_label = tk.Label(self.period_frame, text=f"Количество месяцев в данном файле {max_month}")
            self.max_month_label.grid(row=2, column=0, columnspan=2, pady=5, padx=35)
        elif self.selected_period == "3 месяца":
            # Создаем виджеты для ввода начала и конца месяца
            self.week_label = tk.Label(self.period_frame, text="Введите от и до месяц:")
            self.week_label.grid(row=1, column=0)
            self.start_month_entry = tk.Entry(self.period_frame, width=5)
            self.start_month_entry.grid(row=1, column=1)
            self.end_month_entry = tk.Entry(self.period_frame, width=5)
            self.end_month_entry.grid(row=1, column=2)
            max_month = max(self.data["Месяц"])
            self.max_month_label = tk.Label(self.period_frame, text=f"Количество месяцев в данном файле {max_month}")
            self.max_month_label.grid(row=2, column=0, columnspan=3, pady=5, padx=35)
        # Выводим кнопку "Отменить" для всех случаев
        self.cancel_button = tk.Button(self.period_frame, text="Отменить", command=self.cancel_period)
        self.cancel_button.grid(row=1, column=3, padx=2)
          
    def cancel_period(self):
        # Сбрасываем выбранное значение периода на None
        self.selected_period = None
        # Уничтожаем таблицу и связанную с ней полосу прокрутки
        if hasattr(self, 'table'):
            self.table.destroy()
        if hasattr(self, 'hscrollbar'):
            self.hscrollbar.destroy()
        if hasattr(self, 'histogram_canvas'):
            self.histogram_canvas.get_tk_widget().destroy()            
        # Уничтожаем фрейм выбора периода и пересоздаем виджеты
        self.period_frame.destroy()
        self.period_frame = tk.Frame(self.root)
        self.period_frame.pack(pady=5)
        self.period_label = tk.Label(self.period_frame, text="Выберите период:")
        self.period_label.grid(row=0, column=0, padx=(43, 0))

        periods = ["День", "Неделя", "Месяц", "3 месяца", "6 месяцев"]
        self.period_combo = ttk.Combobox(self.period_frame, values=periods, width=10)
        self.period_combo.grid(row=0, column=1, padx=(0, 45))
        self.period_combo.bind("<<ComboboxSelected>>", self.show_period_widgets)
        
    def plot_amb_stats(self):
        plt.clf()  # Удаляем предыдущий график
        # Проверка, создан ли Combobox
        if self.period_combo is None:
            messagebox.showerror("Ошибка", "Сначала выберите период.")
            return
        # Получение выбранного периода
        selected_period = self.selected_period
        if selected_period == "День":
            self.plot_day_stats()
        elif selected_period == "Неделя":
            self.plot_week_stats()
        elif selected_period == "Месяц":
            self.plot_month_stats()  
        elif selected_period == "3 месяца":
            self.plot_month3_stats()
        elif selected_period == "6 месяцев":
            self.plot_month6_stats() 
 

    def plot_day_stats(self):
        self.save_plot_button4.config(state="disabled")
        self.save_to_excel_button.config(state="disabled")
        self.histogram_canvas.get_tk_widget().destroy()
        # Получение выбранного пользователем месяца и дня
        selected_month = self.month_entry.get()
        selected_day = self.day_entry.get()

        # Преобразование выбранных месяцев в числовой формат 
        selected_month = int(selected_month)
        selected_day = int(selected_day)

        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()

        # Предположим, что self.data содержит DataFrame с данными
        # Фильтрация данных по выбранному месяцу, дню и отделению
        filtered_data = self.data[(self.data["Месяц"] == selected_month) & 
                                  (self.data["День"] == selected_day) & 
                                  (self.data["Наименование отделения"] == selected_department)]
        
        # Если нет данных для выбранного дня и месяца, вывести ошибку
        if filtered_data.empty:
            messagebox.showerror("Ошибка", "Такого дня/месяца нет. Пожалуйста проверьте данные.")
            return

        # Создание текстового виджета для отображения таблицы
        columns = list(filtered_data.columns)
        self.table = ttk.Treeview(self.root, columns=columns, show="headings", height=10)

        # Установка ширины столбцов
        for col in columns:
            self.table.column(col, width=90)  # Set the width as desired

        # Создание горизонтальной полосы прокрутки и сохранение ссылки на неё
        self.hscrollbar = ttk.Scrollbar(self.root, orient="horizontal", command=self.table.xview)
        self.hscrollbar.pack(side="bottom", fill="x")

        # Установка прокрутки для таблицы
        self.table.configure(xscrollcommand=self.hscrollbar.set)

        # Вставка заголовков столбцов
        for col in columns:
            self.table.heading(col, text=col)

        # Вставка данных в таблицу
        for index, row in filtered_data.iterrows():
            self.table.insert("", "end", values=row.tolist())

        # Упаковка таблицы
        self.table.pack(side="left", fill="both", expand=True)

        

    def plot_week_stats(self):
        self.save_plot_button4.config(state=tk.NORMAL)
        self.save_to_excel_button.config(state="disabled")

        # Получение выбранного пользователем номера недели
        selected_week = self.week_entry.get()
        selected_week = int(selected_week)

        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()

        # Фильтрация данных по выбранной неделе и отделению
        filtered_data = self.data[(self.data["Неделя"] == selected_week) & 
                                  (self.data["Наименование отделения"] == selected_department)]

        # Получение данных из столбца "Амб. ВСЕГО" по выбранной неделе и отделению
        selected_dataA = filtered_data["Амб. ВСЕГО"]
        
        if selected_department and selected_week is not None:
            if not selected_dataA.empty:
                # Проверка, все ли значения равны нулю       
                if selected_dataA.eq(0).all():
                    messagebox.showinfo("Информация", f"В отделении {selected_department} на неделе {selected_week} не было случаев.")
                else:
                    # Построение гистограммы
                    plt.figure(figsize=(6, 4))
                    # Построение гистограммы с помощью функции hist(), задав количество корзин (bins) как длину столбца данных
                    plt.hist(range(1, len(selected_dataA) + 1), weights=selected_dataA, bins=len(selected_dataA))
                    plt.rc('axes', labelsize= 10 )
                    plt.rc('axes', titlesize= 10 )
                    plt.xlabel(f'Дни {selected_week} недели')
                    plt.ylabel('Количество амбулаторно принятых')
                    plt.title(f'Гистограмма амб. принятых в отделении {selected_department} на неделе {selected_week}')
                    plt.tight_layout()

                    # Создаем данные для линейной регрессии
                    x = np.arange(1, len(selected_dataA) + 1)
                    y = np.array(selected_dataA)

                    # Вычисляем коэффициенты линейной регрессии
                    coefficients = np.polyfit(x, y, 1)
                    polynomial = np.poly1d(coefficients)
                    linear_regression_line = polynomial(x)

                    # Добавляем линейную регрессию
                    plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                             label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                    plt.legend()

                    # Update the Tkinter canvas containing the histogram plot
                    self.histogram_canvas.get_tk_widget().destroy()
                    self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                    self.histogram_canvas.draw()
                    self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка", f"Нет данных для недели {selected_week} в отделении {selected_department}.")  
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или неделю лечения.")


    def plot_month_stats(self):
        self.save_plot_button4.config(state=tk.NORMAL)
        self.save_to_excel_button.config(state=tk.NORMAL)
        # Получение выбранного пользователем месяца
        selected_month = self.month2_entry.get()
        selected_month = int(selected_month)
        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()
        # Фильтрация данных по выбранной неделе и отделению
        filtered_data = self.data[(self.data["Месяц"] == selected_month) &
                     (self.data["Наименование отделения"] == selected_department)]
        # Получение данных из столбца "Амб. ВСЕГО" по выбранной неделе и отделению
        selected_dataA = filtered_data["Амб. ВСЕГО"]
        # Разбиваем данные на группы по 7 значений и суммируем каждую группу
        grouped_data = [selected_dataA[i:i+7].sum() for i in range(0, len(selected_dataA), 7)]

        if selected_department and selected_month is not None:
            if not selected_dataA.empty:
                # Проверка, все ли значения равны нулю
                if selected_dataA.eq(0).all():
                    messagebox.showinfo("Информация", f"В отделении {selected_department} на неделе {selected_month} не было случаев.")
                else:
                    # Построение гистограммы
                    plt.figure(figsize=(6, 4))
                    # Построение гистограммы с помощью функции hist(), задав количество корзин (bins) как длину столбца данных
                    plt.bar(range(1, len(grouped_data) + 1), grouped_data)  # Группы на оси x, сумма значений на оси y
                    plt.rc('axes', labelsize= 10 )
                    plt.rc('axes', titlesize= 10 )
                    plt.xlabel(f'Недели {selected_month} месяца')
                    plt.ylabel('Количество амбулаторно принятых')
                    plt.title(f'Гистограмма амб. принятых в отделении {selected_department} в месяце {selected_month}')
                    plt.tight_layout()


                    # Создаем данные для линейной регрессии
                    x = np.arange(1, len(grouped_data) + 1)
                    y = np.array(grouped_data)

                    # Вычисляем коэффициенты линейной регрессии
                    coefficients = np.polyfit(x, y, 1)
                    polynomial = np.poly1d(coefficients)
                    linear_regression_line = polynomial(x)

                    # Добавляем линейную регрессию
                    plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                             label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                    plt.legend()

                    # Update the Tkinter canvas containing the histogram plot
                    self.histogram_canvas.get_tk_widget().destroy()
                    self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                    self.histogram_canvas.draw()
                    self.histogram_canvas.get_tk_widget().pack()
            else:
                messagebox.showerror("Ошибка", f"Нет данных для месяца {selected_month} в отделении {selected_department}.")
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или месяц лечения.")



    def plot_month3_stats(self):
        self.save_plot_button4.config(state=tk.NORMAL)
        self.save_to_excel_button.config(state=tk.NORMAL)
        # Получение выбранного пользователем начала-месяц
        selected_start_month = self.start_month_entry.get()
        # Получение выбранного пользователем конца-месяц
        selected_end_month = self.end_month_entry.get()
        # Получение выбранного пользователем отделения
        selected_department = self.selected_department.get()
        # Преобразование выбранных месяцев в числовой формат 
        selected_start_month = int(selected_start_month)
        selected_end_month = int(selected_end_month)

        # Фильтрация данных по выбранным месяцам и отделению
        filtered_data = self.data[(self.data["Месяц"] >= selected_start_month) & (self.data["Месяц"] <= selected_end_month) & (self.data["Наименование отделения"] == selected_department)]
        # Получение данных из столбца "Амб. ВСЕГО" по выбранной неделе и отделению
        grouped_data = filtered_data.groupby("Месяц")["Амб. ВСЕГО"].sum()


        # Действия с данными - построение графика
        if selected_department and selected_start_month and selected_end_month is not None:

                    # Построение гистограммы
                    plt.figure(figsize=(6, 4))
                    plt.bar(range(selected_start_month, selected_end_month+ 1), grouped_data, width=0.4)
                    plt.xticks(range(selected_start_month, selected_end_month + 1))  # Устанавливаем метки по оси x на каждый месяц от month1 до month2
                    plt.rc('axes', labelsize= 10 )
                    plt.rc('axes', titlesize= 10 )
                    plt.xlabel('Месяц')
                    plt.ylabel('Количество пациентов')
                    plt.title(f'Статистика амб. принятых по указанным месяцам в отделении {selected_department}')
                    plt.tight_layout()

                    # Создаем данные для линейной регрессии
                    x = np.arange(len(grouped_data))
                    y = grouped_data.values

                    # Вычисляем коэффициенты линейной регрессии
                    coefficients = np.polyfit(x, y, 1)
                    polynomial = np.poly1d(coefficients)
                    linear_regression_line = polynomial(x)

                    # Добавляем линейную регрессию
                    plt.plot(grouped_data.index, linear_regression_line, color='red', linestyle='-', linewidth=2,
                             label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
                    plt.legend()

                    # Update the Tkinter canvas containing the histogram plot
                    self.histogram_canvas.get_tk_widget().destroy()
                    self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
                    self.histogram_canvas.draw()
                    self.histogram_canvas.get_tk_widget().pack()
        else:
            messagebox.showerror("Ошибка", "Выберите отделение и/или период лечения.")

    def plot_month6_stats(self):
        self.save_plot_button4.config(state=tk.NORMAL)
        self.save_to_excel_button.config(state=tk.NORMAL)
        selected_iA = self.selected_department.get()
        selected_dataA = self.OtdelDataA.get(selected_iA)

        if selected_dataA is not None and not selected_dataA.empty:
            plt.figure(figsize=(6, 4))  # Set figure size
            sns.barplot(x=selected_dataA.index, y=selected_dataA.values)
            plt.rc('axes', labelsize=10)
            plt.rc('axes', titlesize=10)
            plt.xlabel('Месяц')
            plt.ylabel('Количество пациентов')
            plt.title(f'Статистика амб. принятых по месяцам в отделении {selected_iA}')
            plt.tight_layout()

            # Создаем данные для линейной регрессии
            x = np.arange(len(selected_dataA))
            y = np.array(selected_dataA.values)

            # Вычисляем коэффициенты линейной регрессии
            coefficients = np.polyfit(x, y, 1)
            polynomial = np.poly1d(coefficients)
            linear_regression_line = polynomial(x)

            # Добавляем линейную регрессию
            plt.plot(x, linear_regression_line, color='red', linestyle='-', linewidth=2,
                     label=f'Линейная регрессия (y = {coefficients[0]:.2f}x + {coefficients[1]:.2f})')
            plt.legend()

            # Update the Tkinter canvas containing the histogram plot
            self.histogram_canvas.get_tk_widget().destroy()
            self.histogram_canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
            self.histogram_canvas.draw()
            self.histogram_canvas.get_tk_widget().pack()
        else:
            messagebox.showerror("Ошибка", "Выберите отделение или нет данных.")
        
         
        
            
    def save_to_excel(self):
        #Сохранить диаграмму в Эксель
        # Выбираем путь сохранения файла
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            selected_iA = self.selected_department.get()
            selected_dataA = self.OtdelDataA.get(selected_iA)
            if selected_dataA is not None:
                    # Создаем DataFrame из данных
                    df = pd.DataFrame(selected_dataA)
                    df.reset_index(inplace=True)
                    df.columns = ['Месяц', 'Количество пациентов']

                    # Создаем новый Excel файл и записываем в него данные
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Амбулаторное лечение')

                        # Добавляем график в Excel файл
                        worksheet = writer.sheets['Амбулаторное лечение']
                        chart = writer.book.add_chart({'type': 'line'})

                        # Добавляем данные для графика
                        chart.add_series({
                            'categories': ['Амбулаторное лечение', 1, 0, len(df), 0],
                            'values': ['Амбулаторное лечение', 1, 1, len(df), 1],
                            'name': ['Амбулаторное лечение', 0, 1],
                        })

                        # Устанавливаем заголовок и метки осей
                        chart.set_title({'name': 'Статистика амб. принятых по месяцам'})
                        chart.set_x_axis({'name': 'Месяц'})
                        chart.set_y_axis({'name': 'Количество пациентов'})

                        # Вставляем график в лист Excel
                        worksheet.insert_chart('D2', chart)

                    messagebox.showinfo("Успешно", "Данные и график сохранены в файл.")
            else:
                    messagebox.showerror("Ошибка", "Выберите отделение.")
        
    def save_plot4(self):
            # Сохранить диаграмму в файл
            file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
            if file_path:
                plt.savefig(file_path)

    def save_to_text4(self):
            # Сохранить данные в текстовый файл
            selected_iA = self.selected_department.get()
            selected_dataA = self.OtdelDataA.get(selected_iA)
            if selected_dataA is not None:
                file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
                if file_path:
                    with open(file_path, "w") as file:
                        file.write(f'Статистика госпитализаций пациентов по месяцам в отделении {selected_iA}')
                        file.write(selected_dataA.to_string())
                        file.write("Где: 1 - Январь, 2 - Февраль, 3 - Март, 4 - Апрель, 5 - Май, 6 - Июнь")
            else:
                messagebox.showerror("Ошибка", "Выберите отделение.")    
            pass  

        
    def create_text_widget(self):
        # Создание текстового виджета
        self.text_widget = tk.Text(self.root, wrap="none")
        self.text_widget.pack(expand=True, fill="both")

    def update_text_widget(self):
        # Обновление содержимого текстового виджета в соответствии с выбранным отделением
        selected_department = self.selected_department.get()
        self.clear_text_widget()  # Очищаем текстовый виджет перед обновлением
        if selected_department in self.department_data:
            data_to_display = self.department_data[selected_department].to_string()
            self.text_widget.insert(tk.END, data_to_display)
        else:
            self.text_widget.insert(tk.END, "Выберите отделение")

    def show_general_report(self):
        # Здесь логика для вывода общего аналитического отчета
        self.clear_frame()
        if self.data is not None:
            # Создаем словарь с данными для каждого отделения
            departments = self.data["Наименование отделения"].unique()
            self.department_data = {}
            numeric_columns = self.data.select_dtypes(include=['number']).columns[4:]  # Исключаем первые четыре столбца
            for department in departments:
                self.department_data[department] = self.data[self.data["Наименование отделения"] == department][numeric_columns].sum()

            button_width = 25
            
            # кнопка "Возврат"
            self.return_button5 = tk.Button(self.root, text="Возврат", width=button_width, command=self.back)
            self.return_button5.pack(pady=5)
            
            # Создаем дополнительные элементы для выбора отделения
            self.department_label = tk.Label(self.root, text="Выберите отделение:")
            self.department_label.pack()

            self.selected_department = tk.StringVar(self.root)
            self.selected_department.set(departments[0])

            self.department_menu = tk.OptionMenu(self.root, self.selected_department, *departments)
            self.department_menu.pack()

            # Создаем кнопку "Выбрать"
            self.select_button = tk.Button(self.root, text="Выбрать", width=button_width, command=self.update_text_widget)
            self.select_button.pack(pady=5)

            # кнопка "Сохранить в текстовом файле"
            self.save_to_text_button5 = tk.Button(self.root, text="Сохранить в текстовом файле", width=button_width, command=self.save_to_text5)
            self.save_to_text_button5.pack(pady=5)

            # Кнопка "Сохранить в Excel"
            self.save_to_excel_button4 = tk.Button(self.root, text="Сохранить в Excel", width=25, command=lambda: self.save_to_excel4(self.selected_department.get()))
            self.save_to_excel_button4.pack(pady=5)

            # Кнопка "Сохранить ВСЕ в Excel"
            self.save_to_excel_button5 = tk.Button(self.root, text="Сохранить ВСЕ в Excel", width=25, command=self.save_to_excel5)
            self.save_to_excel_button5.pack(pady=5)

            # Создаем текстовый виджет
            self.create_text_widget()

            # Отображаем данные выбранного отделения в текстовом виджете
            self.update_text_widget()


    def save_to_excel4(self, selected_department):
        if selected_department in self.department_data:
            # Create a file dialog to get the file name and location from the user
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

            if file_path:
                # Create a new Excel writer object
                excel_writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

                # Remove invalid characters from the department name for sheet name
                sheet_name = re.sub(r'[^\w\s-]', '', selected_department)[:31]

                # Write selected department's data to a separate sheet
                self.department_data[selected_department].to_excel(excel_writer, sheet_name=sheet_name, index=True)

                # Save the Excel file
                excel_writer.save()

                messagebox.showinfo("Успех", "Данные успешно сохранены в Excel файл.")
        else:
            messagebox.showerror("Ошибка", "Выбранное отделение не содержит данных.")
            
            
    def save_to_excel5(self):
        if self.department_data:
            # Create a file dialog to get the file name and location from the user
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

            if file_path:
                # Create a new Excel writer object
                excel_writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

                # Write each department's data to a separate sheet
                for department, data in self.department_data.items():
                    # Remove invalid characters from the department name
                    sheet_name = re.sub(r'[^\w\s-]', '', department)[:31]
                    data.to_excel(excel_writer, sheet_name=sheet_name, index=True)

                # Save the Excel file
                excel_writer.save()

                messagebox.showinfo("Успех", "Данные успешно сохранены в Excel файл.")
        else:
            messagebox.showerror("Ошибка", "Нет данных для сохранения.")

            
    def save_to_text5(self):
        # Сохранить данные в текстовый файл
        selected_i = self.selected_department.get()
        selected_data = self.department_data.get(selected_i)
        if selected_data is not None:
            file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
            if file_path:
                with open(file_path, "w") as file:
                    file.write(f'Общая статистика в отделении {selected_i}\n')
                    file.write(selected_data.to_string())
        else:
            messagebox.showerror("Ошибка", "Выберите отделение.")    
        pass
    
    def clear_text_widget(self):
            self.text_widget.delete("1.0", tk.END)
        
    def load_file(self):
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
            if file_path:
                try:
                    self.data = pd.read_excel(file_path)
                    self.load_button.config(state=tk.DISABLED)
                    self.enable_buttons()
                    self.return_button.config(state=tk.NORMAL)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Ошибка загрузки файла: {e}")


    def enable_buttons(self):
            for button in self.stat_buttons:
                button.config(state=tk.NORMAL)           

    def clear_data(self):
        self.file_path = None  # Удаление пути к файлу
        self.file_loaded = False
        self.data = None
        self.load_button.config(state=tk.NORMAL)  # Включение кнопки "Загрузить файл"
        self.disable_buttons()  # Отключение остальных кнопок статистики
        self.return_button.config(state=tk.DISABLED) 
        
    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def disable_buttons(self):
        for button in self.stat_buttons:
            button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAnalyzer(root)
    root.mainloop()
  
