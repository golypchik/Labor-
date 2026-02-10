#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Фрейм ключевых элементов
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageTk


class KeyElementsFrame:
    """Фрейм для ключевых элементов"""

    def __init__(self, parent, session_manager):
        self.parent = parent
        self.session_manager = session_manager

        # Переменные для хранения данных
        self.object_name = tk.StringVar()
        self.organization_name = tk.StringVar()
        self.temp_mode = tk.StringVar()
        self.humidity_mode = tk.StringVar()
        self.mapping_date = tk.StringVar()
        self.mapping_datetime = tk.StringVar()
        self.mapping_type = tk.StringVar()
        self.signature_date = tk.StringVar()
        self.employee_position = tk.StringVar()
        self.employee_name = tk.StringVar()
        self.area = tk.StringVar()
        self.photos = []  # Список путей к фотографиям
        self.certificate_continuation = tk.StringVar()
        self.certificate_continuation_copy = tk.StringVar()
        self.certificate_number = tk.StringVar()
        self.research_time = tk.StringVar()
        self.repeated_mapping_date = tk.StringVar()
        self.interval = tk.StringVar()

        # Словарь подсказок для полей
        self.field_hints = {
            "Отчет по картированию написать продолжение:": "ПОМЕЩЕНИЕ ХРАНЕНИЯ ЛЕКАРСТВЕННЫХ СРЕДСТВ №6",
            "Наименование объекта картирования:": "Два шкафа и сейф для хранения лекарственных средств, расположенные в кабинете старшей медсестры реабилитационного отделения №1",
            "Наименование организации заявителя:": "Учреждение здравоохранения «11-я городская клиническая больница», 223028, г. Минск, ул. Корженевского, дом 4",
            "Температурный режим:": "+15℃…+25℃",
            "Влажностный режим:": "не более 60%",
            "Дата проведения картирования:": "22.01.2026 – 26.01.2026",
            "Дата и время проведения картирования:": "22.01.2026 10:00 – 26.01.2026 10:00",
            "Вид картирования:": "Первичное",
            "Дата подписания:": "22.01.2026",
            "Должность сотрудника фирмы:": "Главный врач УЗ «11-я городская клиническая больница»",
            "ФИО сотрудника:": "Часнойть А.Ч.",
            "Площадь помещения:": "14,5",
            "Упрощенное название:": "УЗ «11 городская клиническая больница»",
            "Время проведения исследования:": "3 дня 12 часов",
            "Дата проведения повторного картирования:": "30.10.2028",
            "Интервал:": "1 минута",
        }

        self.create_widgets()
    
    def create_widgets(self):
        """Создание виджетов фрейма"""
        # Основной контейнер с прокруткой
        canvas = tk.Canvas(self.parent, bg="#ecf0f1")
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Привязка прокрутки колесиком мыши и тачпадом
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)

        # Для Linux с Button-4 и Button-5
        def _on_button4(event):
            canvas.yview_scroll(-1, "units")
        def _on_button5(event):
            canvas.yview_scroll(1, "units")
        canvas.bind("<Button-4>", _on_button4)
        canvas.bind("<Button-5>", _on_button5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Заголовок
        title_label = tk.Label(
            scrollable_frame,
            text="Ключевые элементы",
            font=("Arial", 18, "bold"),
            bg="#ecf0f1"
        )
        title_label.pack(pady=20)
        
        # Основной контейнер для полей
        main_frame = tk.Frame(scrollable_frame, bg="#ecf0f1")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Поля ввода
        fields_frame = tk.LabelFrame(
            main_frame,
            text="Основная информация",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        fields_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Список полей для создания
        fields = [
            ("Отчет по картированию написать продолжение:", self.certificate_continuation),
            ("Наименование объекта картирования:", self.object_name),
            ("Наименование организации заявителя:", self.organization_name),
            ("Температурный режим:", self.temp_mode),
            ("Влажностный режим:", self.humidity_mode),
            ("Дата проведения картирования:", self.mapping_date),
            ("Дата и время проведения картирования:", self.mapping_datetime),
            ("Вид картирования:", self.mapping_type),
            ("Дата подписания:", self.signature_date),
            ("Должность сотрудника фирмы:", self.employee_position),
            ("ФИО сотрудника:", self.employee_name),
            ("Площадь помещения:", self.area),
            ("Упрощенное название:", self.certificate_continuation_copy),
            ("Время проведения исследования:", self.research_time),
            ("Дата проведения повторного картирования:", self.repeated_mapping_date),
            ("Интервал:", self.interval),
        ]
        
        self.entry_widgets = {}
        for i, (label_text, var) in enumerate(fields):
            field_frame = tk.Frame(fields_frame, bg="#ecf0f1")
            field_frame.pack(fill=tk.X, pady=5)

            label = tk.Label(
                field_frame,
                text=label_text,
                font=("Arial", 10),
                bg="#ecf0f1",
                width=35,
                anchor=tk.W
            )
            label.pack(side=tk.LEFT, padx=10)

            # Определяем тип поля и его параметры
            is_multiline = label_text in ["Отчет по картированию написать продолжение:", "Наименование объекта картирования:", "Наименование организации заявителя:"]

            if "Дата" in label_text:
                # Контейнер для поля даты с горизонтальной прокруткой
                entry_container = tk.Frame(field_frame, bg="#ecf0f1")
                entry_container.pack(side=tk.LEFT, padx=5, pady=0)

                # Для дат используем Entry с горизонтальной прокруткой (ширина как у других полей)
                entry = tk.Entry(
                    entry_container,
                    textvariable=var,
                    width=28,
                    font=("Arial", 10)
                )

                # Горизонтальная полоса прокрутки
                h_scrollbar = tk.Scrollbar(entry_container, orient=tk.HORIZONTAL, command=entry.xview)
                entry.configure(xscrollcommand=h_scrollbar.set)

                # Привязка прокрутки колесиком мыши
                def _on_mousewheel(event):
                    # Обрабатываем разные значения delta для мыши и тачпада
                    delta = event.delta
                    if abs(delta) > 120:  # Тачпад часто дает большие значения
                        delta = delta // 10  # Нормализуем
                    entry.xview_scroll(int(-1 * (delta / 120)), "units")
                entry.bind("<MouseWheel>", _on_mousewheel)

                # Для Linux с Button-4 и Button-5 (тачпад)
                def _on_button4(event):
                    entry.xview_scroll(-1, "units")
                def _on_button5(event):
                    entry.xview_scroll(1, "units")
                entry.bind("<Button-4>", _on_button4)
                entry.bind("<Button-5>", _on_button5)

                # Дополнительные события для тачпадов
                entry.bind("<ButtonPress-4>", lambda e: entry.xview_scroll(-1, "units"))
                entry.bind("<ButtonPress-5>", lambda e: entry.xview_scroll(1, "units"))

                entry.pack(side=tk.TOP)
                h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
                self.entry_widgets[label_text] = entry
            elif is_multiline:
                # Контейнер для многострочного поля с вертикальной прокруткой
                text_container = tk.Frame(field_frame, bg="#ecf0f1")
                text_container.pack(side=tk.LEFT, padx=5, pady=2)

                # Многострочное поле для длинных названий с переносом слов
                text_widget = tk.Text(
                    text_container,
                    height=5,
                    width=45,
                    font=("Arial", 10),
                    wrap=tk.WORD  # Включаем автоматический перенос слов
                )

                # Вертикальная полоса прокрутки
                v_scrollbar = tk.Scrollbar(text_container, command=text_widget.yview)
                text_widget.configure(yscrollcommand=v_scrollbar.set)

                # Располагаем виджеты
                text_widget.pack(side=tk.LEFT, fill=tk.Y)
                v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

                # Привязка прокрутки колесиком мыши и тачпадом
                def _on_mousewheel(event):
                    # Обрабатываем разные значения delta для мыши и тачпада
                    delta = event.delta
                    if abs(delta) > 120:  # Тачпад часто дает большие значения
                        delta = delta // 10  # Нормализуем
                    text_widget.yview_scroll(int(-1 * (delta / 120)), "units")
                text_widget.bind("<MouseWheel>", _on_mousewheel)

                # Для Linux с Button-4 и Button-5 (тачпад)
                def _on_button4(event):
                    text_widget.yview_scroll(-1, "units")
                def _on_button5(event):
                    text_widget.yview_scroll(1, "units")
                text_widget.bind("<Button-4>", _on_button4)
                text_widget.bind("<Button-5>", _on_button5)

                # Дополнительные события для тачпадов
                text_widget.bind("<ButtonPress-4>", lambda e: text_widget.yview_scroll(-1, "units"))
                text_widget.bind("<ButtonPress-5>", lambda e: text_widget.yview_scroll(1, "units"))

                # Создаем уникальные функции для каждого текстового виджета
                def update_var(widget=text_widget, variable=var):
                    def func(*args):
                        variable.set(widget.get("1.0", "end-1c"))
                    return func

                def update_text(widget=text_widget, variable=var):
                    def func(*args):
                        current = widget.get("1.0", "end-1c")
                        if current != variable.get():
                            widget.delete("1.0", tk.END)
                            widget.insert("1.0", variable.get())
                    return func

                text_widget.bind("<KeyRelease>", update_var())
                var.trace_add("write", update_text())
                self.entry_widgets[label_text] = text_widget
            elif label_text == "Вид картирования:":
                # Выпадающий список с возможностью свободного ввода
                combobox = ttk.Combobox(
                    field_frame,
                    textvariable=var,
                    width=26,
                    font=("Arial", 10),
                    state="normal"  # Разрешаем свободный ввод
                )
                # Устанавливаем варианты для выпадающего списка
                combobox['values'] = ("Первичное", "Повторное")
                combobox.pack(side=tk.LEFT, padx=5)
                self.entry_widgets[label_text] = combobox
            else:
                # Контейнер для однострочного поля с горизонтальной прокруткой
                entry_container = tk.Frame(field_frame, bg="#ecf0f1")
                entry_container.pack(side=tk.LEFT, padx=5, pady=0)

                # Обычное однострочное поле
                entry = tk.Entry(
                    entry_container,
                    textvariable=var,
                    width=28,
                    font=("Arial", 10)
                )

                # Горизонтальная полоса прокрутки
                h_scrollbar = tk.Scrollbar(entry_container, orient=tk.HORIZONTAL, command=entry.xview)
                entry.configure(xscrollcommand=h_scrollbar.set)

                # Привязка прокрутки колесиком мыши
                def _on_mousewheel(event):
                    # Обрабатываем разные значения delta для мыши и тачпада
                    delta = event.delta
                    if abs(delta) > 120:  # Тачпад часто дает большие значения
                        delta = delta // 10  # Нормализуем
                    entry.xview_scroll(int(-1 * (delta / 120)), "units")
                entry.bind("<MouseWheel>", _on_mousewheel)

                # Для Linux с Button-4 и Button-5 (тачпад)
                def _on_button4(event):
                    entry.xview_scroll(-1, "units")
                def _on_button5(event):
                    entry.xview_scroll(1, "units")
                entry.bind("<Button-4>", _on_button4)
                entry.bind("<Button-5>", _on_button5)

                # Дополнительные события для тачпадов
                entry.bind("<ButtonPress-4>", lambda e: entry.xview_scroll(-1, "units"))
                entry.bind("<ButtonPress-5>", lambda e: entry.xview_scroll(1, "units"))

                entry.pack(side=tk.TOP)
                h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
                self.entry_widgets[label_text] = entry

            # Кнопка очистки поля (ставим перед подсказкой)
            if label_text in self.field_hints:
                clear_btn = tk.Button(
                    field_frame,
                    text="✕",
                    command=lambda lt=label_text: self.clear_field(lt),
                    bg="#e74c3c",
                    fg="white",
                    font=("Arial", 10, "bold"),
                    width=2,
                    cursor="hand2"
                )
                clear_btn.pack(side=tk.LEFT, padx=5)

                # Текст подсказки справа от поля (полностью, выбираемый)
                hint_text = self.field_hints[label_text]

                hint_text_widget = tk.Text(
                    field_frame,
                    height=5 if is_multiline else 1,
                    width=40,
                    font=("Arial", 8),
                    wrap=tk.WORD,
                    bg="#f8f9fa",
                    fg="#495057",
                    relief=tk.FLAT,
                    borderwidth=1
                )
                hint_text_widget.insert("1.0", hint_text)
                hint_text_widget.config(state=tk.DISABLED)  # Только для чтения, но можно выделять
                hint_text_widget.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # Загрузка фото
        photo_frame = tk.LabelFrame(
            main_frame,
            text="Фото",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        photo_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Контейнер для фото с двумя колонками
        photos_container = tk.Frame(photo_frame, bg="#ecf0f1")
        photos_container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Левая колонка - кнопки управления
        left_column = tk.Frame(photos_container, bg="#ecf0f1", width=300)
        left_column.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        left_column.pack_propagate(False)
        
        photo_btn_frame = tk.Frame(left_column, bg="#ecf0f1")
        photo_btn_frame.pack(fill=tk.X, pady=5)
        
        load_photo_btn = tk.Button(
            photo_btn_frame,
            text="+ Добавить фото",
            command=self.load_photo,
            bg="#3498db",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        load_photo_btn.pack(side=tk.LEFT, padx=5)
        
        remove_photo_btn = tk.Button(
            photo_btn_frame,
            text="Удалить выбранное",
            command=self.remove_photo,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        remove_photo_btn.pack(side=tk.LEFT, padx=5)
        
        # Кнопки перемещения
        move_btn_frame = tk.Frame(left_column, bg="#ecf0f1")
        move_btn_frame.pack(fill=tk.X, pady=5)
        
        move_up_btn = tk.Button(
            move_btn_frame,
            text="↑ Вверх",
            command=self.move_photo_up,
            bg="#95a5a6",
            fg="white",
            font=("Arial", 9),
            padx=10,
            pady=3,
            cursor="hand2"
        )
        move_up_btn.pack(side=tk.LEFT, padx=5)
        
        move_down_btn = tk.Button(
            move_btn_frame,
            text="↓ Вниз",
            command=self.move_photo_down,
            bg="#95a5a6",
            fg="white",
            font=("Arial", 9),
            padx=10,
            pady=3,
            cursor="hand2"
        )
        move_down_btn.pack(side=tk.LEFT, padx=5)
        
        # Список загруженных фото
        photos_list_frame = tk.Frame(left_column, bg="#ecf0f1")
        photos_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        tk.Label(
            photos_list_frame,
            text="Список фото:",
            font=("Arial", 10, "bold"),
            bg="#ecf0f1"
        ).pack(anchor=tk.W, pady=5)
        
        self.photos_listbox = tk.Listbox(photos_list_frame, height=10, font=("Arial", 9))
        photos_scrollbar = ttk.Scrollbar(photos_list_frame, orient=tk.VERTICAL, command=self.photos_listbox.yview)
        self.photos_listbox.configure(yscrollcommand=photos_scrollbar.set)
        
        self.photos_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        photos_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Правая колонка - превью
        right_column = tk.Frame(photos_container, bg="#ecf0f1")
        right_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)
        
        tk.Label(
            right_column,
            text="Превью фото:",
            font=("Arial", 10, "bold"),
            bg="#ecf0f1"
        ).pack(anchor=tk.W, pady=5)
        
        # Фрейм для превью с прокруткой
        preview_canvas = tk.Canvas(right_column, bg="#ecf0f1", highlightthickness=1, highlightbackground="#bdc3c7")
        preview_scrollbar = ttk.Scrollbar(right_column, orient="vertical", command=preview_canvas.yview)
        preview_scrollable = tk.Frame(preview_canvas, bg="#ecf0f1")
        
        preview_scrollable.bind(
            "<Configure>",
            lambda e: preview_canvas.configure(scrollregion=preview_canvas.bbox("all"))
        )
        
        preview_canvas.create_window((0, 0), window=preview_scrollable, anchor="nw")
        preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
        
        preview_canvas.pack(side="left", fill="both", expand=True)
        preview_scrollbar.pack(side="right", fill="y")
        
        # Привязка прокрутки колесиком мыши
        def _on_mousewheel(event):
            preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        preview_canvas.bind("<MouseWheel>", _on_mousewheel)
        
        self.preview_frame = preview_scrollable
        self.preview_canvas = preview_canvas
        
        # Привязка выбора в списке к отображению превью
        self.photos_listbox.bind('<<ListboxSelect>>', self.on_photo_select)
        
        # Кнопка сохранения
        save_btn = tk.Button(
            main_frame,
            text="Сохранить",
            command=self.save_data,
            bg="#27ae60",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=10,
            cursor="hand2"
        )
        save_btn.pack(pady=20)
    
    def load_photo(self):
        """Загрузка фото"""
        file_paths = filedialog.askopenfilenames(
            title="Выберите фото",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
        )
        
        for file_path in file_paths:
            try:
                # Проверяем, что это изображение
                img = Image.open(file_path)
                self.photos.append(file_path)
                self.photos_listbox.insert(tk.END, Path(file_path).name)
                self.update_preview()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить фото {Path(file_path).name}:\n{e}")
    
    def remove_photo(self):
        """Удаление выбранного фото"""
        selection = self.photos_listbox.curselection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите фото для удаления")
            return
        
        index = selection[0]
        self.photos_listbox.delete(index)
        del self.photos[index]
        self.update_preview()
    
    def move_photo_up(self):
        """Перемещение фото вверх"""
        selection = self.photos_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        
        index = selection[0]
        # Меняем местами в списке
        self.photos[index], self.photos[index - 1] = self.photos[index - 1], self.photos[index]
        
        # Обновляем отображение
        self.refresh_photos_list()
        self.photos_listbox.selection_set(index - 1)
        self.update_preview()
    
    def move_photo_down(self):
        """Перемещение фото вниз"""
        selection = self.photos_listbox.curselection()
        if not selection or selection[0] >= len(self.photos) - 1:
            return
        
        index = selection[0]
        # Меняем местами в списке
        self.photos[index], self.photos[index + 1] = self.photos[index + 1], self.photos[index]
        
        # Обновляем отображение
        self.refresh_photos_list()
        self.photos_listbox.selection_set(index + 1)
        self.update_preview()
    
    def refresh_photos_list(self):
        """Обновление списка фото"""
        self.photos_listbox.delete(0, tk.END)
        for photo_path in self.photos:
            self.photos_listbox.insert(tk.END, Path(photo_path).name)
    
    def on_photo_select(self, event):
        """Обработчик выбора фото в списке"""
        self.update_preview()
    
    def update_preview(self):
        """Обновление превью фото"""
        # Очищаем превью
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        
        selection = self.photos_listbox.curselection()
        if selection and selection[0] < len(self.photos):
            photo_path = self.photos[selection[0]]
            try:
                img = Image.open(photo_path)
                # Масштабируем для превью (максимальная ширина 400px)
                max_width = 400
                width, height = img.size
                if width > max_width:
                    ratio = max_width / width
                    new_width = max_width
                    new_height = int(height * ratio)
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(img)
                
                # Отображаем превью
                preview_label = tk.Label(
                    self.preview_frame,
                    image=photo,
                    bg="#ecf0f1"
                )
                preview_label.image = photo  # Сохраняем ссылку
                preview_label.pack(pady=10)
                
                # Имя файла
                name_label = tk.Label(
                    self.preview_frame,
                    text=Path(photo_path).name,
                    font=("Arial", 9),
                    bg="#ecf0f1"
                )
                name_label.pack(pady=5)
            except Exception as e:
                error_label = tk.Label(
                    self.preview_frame,
                    text=f"Ошибка загрузки превью: {e}",
                    font=("Arial", 9),
                    bg="#ecf0f1",
                    fg="red"
                )
                error_label.pack(pady=10)
    
    def save_data(self):
        """Сохранение данных"""
        # Проверка обязательных полей
        # Для многострочных полей (Text) напрямую проверяем содержимое
        object_name_widget = self.entry_widgets["Наименование объекта картирования:"]
        object_name_text = object_name_widget.get("1.0", "end-1c").strip()
        
        if not object_name_text:
            messagebox.showwarning("Предупреждение", "Заполните поле 'Наименование объекта картирования'")
            return
        
        # Обновляем StringVar из содержимого Text виджетов (чтобы гарантировать синхронизацию)
        for label_text, widget in self.entry_widgets.items():
            if isinstance(widget, tk.Text):
                text = widget.get("1.0", "end-1c")
                # Находим соответствующую StringVar
                for field_label, var in [
                    ("Отчет по картированию написать продолжение:", self.certificate_continuation),
                    ("Наименование объекта картирования:", self.object_name),
                    ("Наименование организации заявителя:", self.organization_name),
                    ("Температурный режим:", self.temp_mode),
                    ("Влажностный режим:", self.humidity_mode),
                    ("Дата проведения картирования:", self.mapping_date),
                    ("Дата и время проведения картирования:", self.mapping_datetime),
                    ("Вид картирования:", self.mapping_type),
                    ("Дата подписания:", self.signature_date),
                    ("Должность сотрудника фирмы:", self.employee_position),
                    ("ФИО сотрудника:", self.employee_name),
                    ("Площадь помещения:", self.area),
                    ("Упрощенное название:", self.certificate_continuation_copy),
                    ("Время проведения исследования:", self.research_time),
                    ("Дата проведения повторного картирования:", self.repeated_mapping_date),
                    ("Интервал:", self.interval),
                ]:
                    if label_text == field_label:
                        var.set(text)
                        break
        
        # Сохраняем данные в БД настроек
        import sqlite3
        settings_db_path = self.session_manager.get_settings_db_path()
        conn = sqlite3.connect(settings_db_path)
        cursor = conn.cursor()
        
        data = {
            'object_name': self.object_name.get(),
            'organization_name': self.organization_name.get(),
            'temp_mode': self.temp_mode.get(),
            'humidity_mode': self.humidity_mode.get(),
            'mapping_date': self.mapping_date.get(),
            'mapping_datetime': self.mapping_datetime.get(),
            'mapping_type': self.mapping_type.get(),
            'signature_date': self.signature_date.get(),
            'employee_position': self.employee_position.get(),
            'employee_name': self.employee_name.get(),
            'area': self.area.get(),
            'photo_paths': ','.join(self.photos),
            'certificate_continuation': self.certificate_continuation.get(),
            'certificate_continuation_copy': self.certificate_continuation_copy.get(),
            'certificate_number': self.certificate_number.get(),
            'research_time': self.research_time.get(),
            'repeated_mapping_date': self.repeated_mapping_date.get(),
            'interval': self.interval.get()
        }
        
        for key, value in data.items():
            cursor.execute("""
                INSERT OR REPLACE INTO settings (key, value)
                VALUES (?, ?)
            """, (key, value))
        
        conn.commit()
        conn.close()
        
        # Кастомное модальное окно без системного звука
        self.show_custom_message("Сохранение", "Данные сохранены успешно!")

    def clear_data(self):
        """Очистка данных фрейма"""
        from datetime import datetime
        # Очищаем переменные StringVar
        self.object_name.set('')
        self.organization_name.set('')
        self.temp_mode.set('')
        self.humidity_mode.set('')
        self.mapping_date.set('')
        self.mapping_type.set('')
        self.signature_date.set('')
        self.employee_position.set('')
        self.employee_name.set('')
        self.area.set('')
        self.photos = []
        self.certificate_continuation.set('')
        self.certificate_continuation_copy.set('')
        self.research_time.set('')
        self.repeated_mapping_date.set('')
        self.interval.set('')
        # Очищаем текстовые виджеты напрямую (так как они не всегда синхронизируются с StringVar)
        for label_text, widget in self.entry_widgets.items():
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
            elif isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, ttk.Combobox):
                widget.set('')
        # Очищаем список фото
        self.photos_listbox.delete(0, tk.END)
        self.update_preview()

    def get_data(self):
        """Получение данных"""
        return {
            'object_name': self.object_name.get(),
            'organization_name': self.organization_name.get(),
            'temp_mode': self.temp_mode.get(),
            'humidity_mode': self.humidity_mode.get(),
            'mapping_date': self.mapping_date.get(),
            'mapping_datetime': self.mapping_datetime.get(),
            'mapping_type': self.mapping_type.get(),
            'signature_date': self.signature_date.get(),
            'employee_position': self.employee_position.get(),
            'employee_name': self.employee_name.get(),
            'area': self.area.get(),
            'photo_paths': self.photos,
            'certificate_continuation': self.certificate_continuation.get(),
            'certificate_continuation_copy': self.certificate_continuation_copy.get(),
            'research_time': self.research_time.get(),
            'repeated_mapping_date': self.repeated_mapping_date.get(),
            'interval': self.interval.get()
        }

    def pack(self, **kwargs):
        """Упаковка фрейма"""
        self.parent.pack(**kwargs)



    def clear_field(self, field_label):
        """Очистка поля по нажатию кнопки ✕"""
        if field_label in self.entry_widgets:
            widget = self.entry_widgets[field_label]
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                # Обновляем связанную StringVar
                for label_text, var in [
                    ("Отчет по картированию написать продолжение:", self.certificate_continuation),
                    ("Наименование объекта картирования:", self.object_name),
                    ("Наименование организации заявителя:", self.organization_name),
                    ("Температурный режим:", self.temp_mode),
                    ("Влажностный режим:", self.humidity_mode),
                    ("Дата проведения картирования:", self.mapping_date),
                    ("Дата и время проведения картирования:", self.mapping_datetime),
                    ("Вид картирования:", self.mapping_type),
                    ("Дата подписания:", self.signature_date),
                    ("Должность сотрудника фирмы:", self.employee_position),
                    ("ФИО сотрудника:", self.employee_name),
                    ("Площадь помещения:", self.area),
                    ("Упрощенное название:", self.certificate_continuation_copy),
                    ("Время проведения исследования:", self.research_time),
                    ("Дата проведения повторного картирования:", self.repeated_mapping_date),
                    ("Интервал:", self.interval),
                ]:
                    if label_text == field_label:
                        var.set('')
                        break
            elif isinstance(widget, ttk.Combobox):
                widget.set('')
                # Обновляем связанную StringVar
                for label_text, var in [
                    ("Вид картирования:", self.mapping_type),
                ]:
                    if label_text == field_label:
                        var.set('')
                        break

    def show_custom_message(self, title, message):
        """Кастомное модальное окно без системного звука"""
        # Создаем топ-level окно
        top = tk.Toplevel(self.parent)
        top.title(title)
        top.geometry("300x150")
        top.resizable(False, False)
        top.attributes("-topmost", True)  # всегда наверху
        
        # Центрируем окно
        top.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (300 // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (150 // 2)
        top.geometry(f"+{x}+{y}")
        
        # Добавляем сообщение
        label = tk.Label(top, text=message, font=("Arial", 12), padx=20, pady=30)
        label.pack()
        
        # Кнопка закрытия
        def close_window():
            top.destroy()
        
        ok_btn = tk.Button(top, text="ОК", command=close_window, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), padx=20, pady=5)
        ok_btn.pack(pady=5)
        
        # Делаем окно модальным
        top.grab_set()
        top.wait_window()

    def pack_forget(self):
        """Скрытие фрейма"""
        self.parent.pack_forget()
