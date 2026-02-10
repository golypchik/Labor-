#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Фрейм управления проектом
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import shutil
import pandas as pd
from datetime import datetime
from itertools import combinations


class ProjectManagementFrame:
    """Фрейм для управления проектом"""
    
    def __init__(self, parent, session_manager):
        self.parent = parent
        self.session_manager = session_manager
        self.selected_files = []
        self.report_type = tk.StringVar(value="Объект хранения")
        self.use_humidity = tk.BooleanVar(value=False)
        self.logger_screenshots = []  # [(номер_логгера, путь), ...] — скриншоты для Приложения 5
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
            text="Управление проектом",
            font=("Arial", 18, "bold"),
            bg="#ecf0f1"
        )
        title_label.pack(pady=20)
        
        # 1. Подгрузка данных
        data_frame = tk.LabelFrame(
            scrollable_frame,
            text="1. Подгрузить данные",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        data_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Контейнер для кнопок (вертикальная раскладка)
        buttons_frame = tk.Frame(data_frame, bg="#ecf0f1")
        buttons_frame.pack(side=tk.LEFT, padx=10, pady=5)

        load_btn = tk.Button(
            buttons_frame,
            text="Выбрать\nExcel файлы",
            command=self.load_excel_files,
            bg="#3498db",
            fg="white",
            font=("Arial", 11),
            width=12,
            padx=10,
            pady=8,
            cursor="hand2"
        )
        load_btn.pack(pady=5)

        add_btn = tk.Button(
            buttons_frame,
            text="Добавить\nфайлы",
            command=self.add_excel_files,
            bg="#27ae60",
            fg="white",
            font=("Arial", 11),
            width=12,
            padx=10,
            pady=8,
            cursor="hand2"
        )
        add_btn.pack(pady=5)

        clear_btn = tk.Button(
            buttons_frame,
            text="Очистить\nвсе",
            command=self.clear_files,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 11),
            width=12,
            padx=10,
            pady=8,
            cursor="hand2"
        )
        clear_btn.pack(pady=5)
        
        # Список выбранных файлов с прокруткой
        files_list_frame = tk.Frame(data_frame, bg="#ecf0f1")
        files_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Создаем контейнер для Treeview с горизонтальной прокруткой
        tree_container = tk.Frame(files_list_frame, bg="#ecf0f1")
        tree_container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview для отображения файлов с временными диапазонами и временем исследования
        columns = ("Файл", "Путь", "Начало записи", "Конец записи", "Время исследования")
        self.files_tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=8)
        self.files_tree.heading("Файл", text="Имя файла")
        self.files_tree.heading("Путь", text="Путь")
        self.files_tree.heading("Начало записи", text="Начало записи")
        self.files_tree.heading("Конец записи", text="Конец записи")
        self.files_tree.heading("Время исследования", text="Время исследования")
        self.files_tree.column("Файл", width=150)
        self.files_tree.column("Путь", width=250)
        self.files_tree.column("Начало записи", width=130)
        self.files_tree.column("Конец записи", width=130)
        self.files_tree.column("Время исследования", width=150)

        # Вертикальная полоса прокрутки
        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=v_scrollbar.set)

        # Горизонтальная полоса прокрутки
        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.files_tree.xview)
        self.files_tree.configure(xscrollcommand=h_scrollbar.set)

        # Располагаем элементы
        self.files_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Настройка сетки
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Кнопка удаления выбранного файла
        remove_btn = tk.Button(
            data_frame,
            text="Удалить выбранный",
            command=self.remove_selected_file,
            bg="#e67e22",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        remove_btn.pack(pady=5)

        # Общие временные диапазоны логгеров
        ranges_frame = tk.LabelFrame(data_frame, text="Общие временные диапазоны", font=("Arial", 10, "bold"), bg="#ecf0f1")
        ranges_frame.pack(fill=tk.X, pady=10)

        self.ranges_text = tk.Text(ranges_frame, height=4, font=("Arial", 9), bg="#f8f9fa")
        ranges_scrollbar = ttk.Scrollbar(ranges_frame, orient=tk.VERTICAL, command=self.ranges_text.yview)
        self.ranges_text.configure(yscrollcommand=ranges_scrollbar.set)

        self.ranges_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        ranges_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.ranges_text.insert("1.0", "Загрузите Excel файлы для отображения общих временных диапазонов логгеров")
        self.ranges_text.config(state=tk.DISABLED)

        # 2. Тип отчета
        report_frame = tk.LabelFrame(
            scrollable_frame,
            text="2. Тип отчета",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        report_frame.pack(fill=tk.X, padx=20, pady=10)
        
        report_types = ["Объект хранения", "Зона хранения", "Холодильник/Морозильник"]
        for report_type in report_types:
            rb = tk.Radiobutton(
                report_frame,
                text=report_type,
                variable=self.report_type,
                value=report_type,
                font=("Arial", 11),
                bg="#ecf0f1",
                activebackground="#ecf0f1"
            )
            rb.pack(anchor=tk.W, padx=20, pady=5)
        
        # 3. Учитывать влажность
        humidity_frame = tk.LabelFrame(
            scrollable_frame,
            text="3. Учитывать влажность",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        humidity_frame.pack(fill=tk.X, padx=20, pady=10)
        
        humidity_check = tk.Checkbutton(
            humidity_frame,
            text="Учитывать влажность в отчете",
            variable=self.use_humidity,
            font=("Arial", 11),
            bg="#ecf0f1",
            activebackground="#ecf0f1"
        )
        humidity_check.pack(anchor=tk.W, padx=20, pady=5)

        # 4. Скриншоты графиков по логгерам (Приложение 5)
        screenshots_frame = tk.LabelFrame(
            scrollable_frame,
            text="4. Скриншоты графиков по логгерам (Приложение 5)",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        screenshots_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        tk.Label(screenshots_frame, text="Добавляйте скриншоты в порядке логгеров. Номер логгера — из названия Excel-файла.",
                 font=("Arial", 9), bg="#ecf0f1", fg="#666").pack(anchor=tk.W, pady=(0, 5))

        ss_container = tk.Frame(screenshots_frame, bg="#ecf0f1")
        ss_container.pack(fill=tk.BOTH, expand=True, pady=5)

        ss_list_frame = tk.Frame(ss_container, bg="#ecf0f1")
        ss_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.screenshots_listbox = tk.Listbox(ss_list_frame, height=6, font=("Arial", 10))
        ss_scrollbar = ttk.Scrollbar(ss_list_frame, command=self.screenshots_listbox.yview)
        self.screenshots_listbox.configure(yscrollcommand=ss_scrollbar.set)
        self.screenshots_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ss_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.screenshots_listbox.bind("<<ListboxSelect>>", self._on_screenshot_select)

        ss_btns = tk.Frame(ss_container, bg="#ecf0f1", width=160)
        ss_btns.pack(side=tk.RIGHT, fill=tk.Y)
        ss_btns.pack_propagate(False)

        tk.Button(ss_btns, text="+ Добавить", command=self._add_logger_screenshot,
                  bg="#27ae60", fg="white", font=("Arial", 10), padx=8, pady=5, cursor="hand2").pack(pady=3, fill=tk.X)
        tk.Button(ss_btns, text="Удалить", command=self._remove_logger_screenshot,
                  bg="#e74c3c", fg="white", font=("Arial", 10), padx=8, pady=5, cursor="hand2").pack(pady=3, fill=tk.X)
       
        preview_frame = tk.LabelFrame(screenshots_frame, text="Превью", font=("Arial", 10, "bold"), bg="#ecf0f1")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.screenshot_preview_label = tk.Label(preview_frame, text="Выберите скриншот", bg="#f8f9fa", fg="#666")
        self.screenshot_preview_label.pack(pady=20, padx=20)

        # Кнопка сохранения
        save_btn = tk.Button(
            scrollable_frame,
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
    
    def _refresh_screenshots_listbox(self):
        self.screenshots_listbox.delete(0, tk.END)
        for num, path in self.logger_screenshots:
            name = Path(path).name if path else "—"
            self.screenshots_listbox.insert(tk.END, f"Логгер №{num} — {name}")

    def _on_screenshot_select(self, event):
        sel = self.screenshots_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.logger_screenshots):
            return
        num, path = self.logger_screenshots[idx]
        self.screenshot_preview_label.config(text=f"Логгер №{num}")
        if path and Path(path).exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(path)
                img.thumbnail((300, 300))
                photo = ImageTk.PhotoImage(img)
                self.screenshot_preview_label.config(image=photo, text="")
                self.screenshot_preview_label.image = photo
            except Exception:
                self.screenshot_preview_label.config(image="", text=f"Логгер №{num}")
        else:
            self.screenshot_preview_label.config(image="", text=f"Логгер №{num}")

    def _add_logger_screenshot(self):
        """Добавление одного или нескольких скриншотов. Номер логгера берется из названия изображения. Пользователь может изменить номер."""
        # Получаем количество загруженных Excel файлов
        excel_count = len(self.files_tree.get_children())
        if excel_count == 0:
            messagebox.showwarning("Предупреждение", "Сначала загрузите Excel файлы для анализа")
            return

        # Получаем текущее количество уже добавленных скриншотов
        current_screenshots_count = len(self.logger_screenshots)
        remaining_slots = excel_count - current_screenshots_count
        
        if remaining_slots <= 0:
            messagebox.showwarning("Предупреждение", 
                f"Вы уже добавили скриншоты для всех {excel_count} Excel файлов.\n"
                f"Удалите некоторые скриншоты, если хотите добавить новые.")
            return

        paths = filedialog.askopenfilenames(
            title=f"Выберите скриншоты (максимум {remaining_slots} изображений, осталось мест для {remaining_slots} логгеров)",
            filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp"), ("Все файлы", "*.*")]
        )
        if not paths:
            return

        # Проверяем количество выбранных изображений с учетом уже добавленных
        if len(paths) > remaining_slots:
            messagebox.showwarning("Предупреждение", 
                f"Вы выбрали {len(paths)} изображений, но осталось только {remaining_slots} свободных мест для логгеров.\n"
                f"Пожалуйста, выберите не более {remaining_slots} изображений.")
            return

        # Сортируем пути по имени файла для предсказуемого порядка
        paths = sorted(paths, key=lambda p: Path(p).name)

        # Получаем номера логгеров из Excel файлов
        excel_logger_nums = self._get_excel_logger_numbers()

        # Создаем диалог для ввода номеров логгеров
        d = tk.Toplevel(self.parent)
        d.title("Номера логгеров")
        d.geometry("600x300")  # Увеличено в два раза
        d.transient(self.parent)
        d.grab_set()

        tk.Label(d, text="Введите номера логгеров для выбранных изображений:", font=("Arial", 10, "bold")).pack(pady=5)

        # Отображаем подсказку с номерами из Excel файлов (один раз)
        if excel_logger_nums:
            hint_text = f"Номера логгеров из Excel файлов: {', '.join(excel_logger_nums)}"
            tk.Label(d, text=hint_text, font=("Arial", 9), fg="#666").pack(pady=2)

        # Контейнер для списка изображений и полей ввода
        list_frame = tk.Frame(d)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Создаем список для хранения виджетов
        entries = []
        for i, path in enumerate(paths):
            fn = Path(path).name
            import re
            m = re.search(r'\d+', fn)
            suggested_num = m.group(0) if m else str(len(self.logger_screenshots) + i + 1)

            row = tk.Frame(list_frame)
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=f"{i+1}. {fn}", width=25, anchor="w").pack(side=tk.LEFT)
            ent = tk.Entry(row, width=10, font=("Arial", 10))
            ent.insert(0, suggested_num)
            ent.pack(side=tk.LEFT, padx=5)

            entries.append((path, ent))

        # Кнопки управления
        btn_frame = tk.Frame(d)
        btn_frame.pack(pady=10)

        def ok():
            for path, ent in entries:
                num = ent.get().strip()
                if num:
                    self.logger_screenshots.append((num, path))
            self._refresh_screenshots_listbox()
            d.destroy()

        def cancel():
            d.destroy()

        tk.Button(btn_frame, text="ОК", command=ok, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Отмена", command=cancel, bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)

        # Фокус на первое поле ввода
        if entries:
            entries[0][1].focus()
        d.lift()
        d.focus_force()

    def _get_excel_logger_numbers(self):
        """Получение номеров логгеров из загруженных Excel файлов."""
        logger_nums = []
        for item_id in self.files_tree.get_children():
            values = self.files_tree.item(item_id, "values")
            fn = Path(values[0]).stem if values else ""
            import re
            m = re.search(r'\d+', fn)
            if m:
                logger_nums.append(m.group(0))
            elif fn:
                logger_nums.append(fn)
        return logger_nums

    def _add_multiple_logger_screenshots(self):
        """Добавление нескольких скриншотов за раз. Номер логгера берется из названия изображения. Пользователь может изменить номер."""
        paths = filedialog.askopenfilenames(
            title="Выберите скриншоты (можно несколько)",
            filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp"), ("Все файлы", "*.*")]
        )
        if not paths:
            return

        # Сортируем пути по имени файла для предсказуемого порядка
        paths = sorted(paths, key=lambda p: Path(p).name)

        # Создаем диалог для ввода номеров логгеров
        d = tk.Toplevel(self.parent)
        d.title("Номера логгеров")
        d.geometry("400x300")
        d.transient(self.parent)
        d.grab_set()

        tk.Label(d, text="Введите номера логгеров для выбранных изображений:", font=("Arial", 10, "bold")).pack(pady=5)

        # Контейнер для списка изображений и полей ввода
        list_frame = tk.Frame(d)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Создаем список для хранения виджетов
        entries = []
        for i, path in enumerate(paths):
            fn = Path(path).name
            import re
            m = re.search(r'\d+', fn)
            suggested_num = m.group(0) if m else str(len(self.logger_screenshots) + i + 1)

            row = tk.Frame(list_frame)
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=f"{i+1}. {fn}", width=25, anchor="w").pack(side=tk.LEFT)
            ent = tk.Entry(row, width=10, font=("Arial", 10))
            ent.insert(0, suggested_num)
            ent.pack(side=tk.LEFT, padx=5)
            entries.append((path, ent))

        # Кнопки управления
        btn_frame = tk.Frame(d)
        btn_frame.pack(pady=10)

        def ok():
            for path, ent in entries:
                num = ent.get().strip()
                if num:
                    self.logger_screenshots.append((num, path))
            self._refresh_screenshots_listbox()
            d.destroy()

        def cancel():
            d.destroy()

        tk.Button(btn_frame, text="ОК", command=ok, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Отмена", command=cancel, bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)

        # Фокус на первое поле ввода
        if entries:
            entries[0][1].focus()
        d.lift()
        d.focus_force()

    def _remove_logger_screenshot(self):
        sel = self.screenshots_listbox.curselection()
        if sel:
            idx = sel[0]
            del self.logger_screenshots[idx]
            self._refresh_screenshots_listbox()
            self.screenshot_preview_label.config(image="", text="Выберите скриншот")
    
    def load_excel_files(self):
        """Загрузка Excel файлов"""
        files = filedialog.askopenfilenames(
            title="Выберите Excel файлы",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if files:
            self.selected_files = []
            self.files_tree.delete(*self.files_tree.get_children())
            self.copy_files_to_inform(files)
    
    def add_excel_files(self):
        """Добавление дополнительных Excel файлов"""
        files = filedialog.askopenfilenames(
            title="Выберите дополнительные Excel файлы",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if files:
            self.copy_files_to_inform(files)
    
    def copy_files_to_inform(self, files):
        """Копирование файлов в папку inform и извлечение временных диапазонов"""
        inform_dir = self.session_manager.inform_dir
        inform_dir.mkdir(exist_ok=True)

        for file_path in files:
            src = Path(file_path)
            dst = inform_dir / src.name

            try:
                shutil.copy2(src, dst)
                self.selected_files.append(str(dst))

                # Извлекаем временной диапазон из Excel файла
                start_time, end_time = self.extract_time_range(str(dst))

                # Вычисляем время исследования
                research_time = self.calculate_research_time(start_time, end_time)

                # Добавляем в список с временными диапазонами и временем исследования
                self.files_tree.insert("", tk.END, values=(src.name, str(dst), start_time, end_time, research_time))

            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось скопировать файл {src.name}:\n{e}")

        # Обновляем отображение общих временных диапазонов
        self.update_common_ranges_display()

    def calculate_research_time(self, start_time_str, end_time_str):
        """Вычисление времени исследования в формате 'X дней Y часов Z минут'"""
        try:
            if start_time_str in ["Не найдено", "Нет данных", "Ошибка"] or \
               end_time_str in ["Не найдено", "Нет данных", "Ошибка"]:
                return "Недоступно"

            start_time = datetime.strptime(start_time_str, "%d.%m.%Y %H:%M")
            end_time = datetime.strptime(end_time_str, "%d.%m.%Y %H:%M")

            # Вычисляем разницу
            delta = end_time - start_time

            # Разбиваем на дни, часы, минуты
            days = delta.days
            hours = delta.seconds // 3600
            minutes = (delta.seconds % 3600) // 60

            # Формируем строку
            parts = []
            if days > 0:
                parts.append(f"{days} {'день' if days == 1 else 'дня' if days < 5 else 'дней'}")
            if hours > 0:
                parts.append(f"{hours} {'час' if hours == 1 else 'часа' if hours < 5 else 'часов'}")
            if minutes > 0 or (days == 0 and hours == 0):
                parts.append(f"{minutes} {'минута' if minutes == 1 else 'минуты' if minutes < 5 else 'минут'}")

            return " ".join(parts) if parts else "0 минут"

        except Exception:
            return "Ошибка расчета"

    def extract_time_range(self, file_path):
        """Извлечение временного диапазона из Excel файла"""
        try:
            # Читаем Excel файл
            df = pd.read_excel(file_path)

            # Ищем колонку с датой/временем (обычно называется 'Date' или 'Time' или 'Дата')
            time_column = None
            for col in df.columns:
                if any(keyword in str(col).lower() for keyword in ['date', 'time', 'дата', 'время']):
                    time_column = col
                    break

            if time_column is None:
                return "Не найдено", "Не найдено"

            # Преобразуем в datetime
            time_values = pd.to_datetime(df[time_column], errors='coerce').dropna()

            if len(time_values) == 0:
                return "Нет данных", "Нет данных"

            # Находим минимальное и максимальное время
            start_time = time_values.min().strftime("%d.%m.%Y %H:%M")
            end_time = time_values.max().strftime("%d.%m.%Y %H:%M")

            return start_time, end_time

        except Exception as e:
            return f"Ошибка: {str(e)[:20]}", f"Ошибка: {str(e)[:20]}"
    
    def remove_selected_file(self):
        """Удаление выбранного файла"""
        selection = self.files_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите файл для удаления")
            return

        for item in selection:
            values = self.files_tree.item(item, "values")
            file_path = values[1]

            # Удаляем файл
            try:
                Path(file_path).unlink()
                self.selected_files.remove(file_path)
                self.files_tree.delete(item)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось удалить файл:\n{e}")

        # Обновляем отображение общих диапазонов
        self.update_common_ranges_display()
    
    def clear_files(self):
        """Очистка всех файлов"""
        if not self.selected_files:
            return

        if messagebox.askyesno("Подтверждение", "Удалить все загруженные файлы?"):
            for file_path in self.selected_files[:]:
                try:
                    Path(file_path).unlink()
                except Exception as e:
                    print(f"Ошибка удаления файла {file_path}: {e}")

            self.selected_files = []
            self.files_tree.delete(*self.files_tree.get_children())

            # Очищаем отображение диапазонов
            self.update_common_ranges_display()

    def update_common_ranges_display(self):
        """Обновление отображения общих временных диапазонов логгеров.
        Логика: 1) если все логгеры имеют общий диапазон — указываем его;
        2) иначе если 50%+ логгеров имеют пересечение — берём максимальную такую группу;
        3) логгеры, не входящие в диапазон, перечисляем отдельно;
        4) логгеры без диапазона (ошибка/нет данных) перечисляем отдельно."""
        self.ranges_text.config(state=tk.NORMAL)
        self.ranges_text.delete("1.0", tk.END)

        if not self.selected_files:
            self.ranges_text.insert("1.0", "Загрузите Excel файлы для отображения общих временных диапазонов логгеров")
            self.ranges_text.config(state=tk.DISABLED)
            return

        try:
            valid_files = []
            invalid_files = []

            for item_id in self.files_tree.get_children():
                values = self.files_tree.item(item_id, "values")
                file_name = values[0]
                start_str = values[2]
                end_str = values[3]

                invalid_reason = None
                if start_str in ("Не найдено", "Нет данных", "Ошибка") or start_str.startswith("Ошибка"):
                    invalid_reason = start_str
                elif end_str in ("Не найдено", "Нет данных", "Ошибка") or end_str.startswith("Ошибка"):
                    invalid_reason = end_str

                if invalid_reason:
                    invalid_files.append((file_name, invalid_reason))
                    continue
                try:
                    start_dt = datetime.strptime(start_str, "%d.%m.%Y %H:%M")
                    end_dt = datetime.strptime(end_str, "%d.%m.%Y %H:%M")
                    valid_files.append((file_name, start_dt, end_dt))
                except ValueError:
                    invalid_files.append((file_name, "Неверный формат"))

            total_loggers = len(valid_files) + len(invalid_files)
            if not valid_files:
                msg = "Не удалось извлечь временные диапазоны из файлов"
                if invalid_files:
                    msg += f"\n\nЛоггеры без диапазона: {', '.join(f[0] for f in invalid_files)}"
                self.ranges_text.insert("1.0", msg)
                self.ranges_text.config(state=tk.DISABLED)
                return

            valid_count = len(valid_files)
            
            # Находим общий диапазон для всех логгеров (пересечение всех диапазонов)
            all_starts = [f[1] for f in valid_files]
            all_ends = [f[2] for f in valid_files]
            common_start_all = max(all_starts)
            common_end_all = min(all_ends)
            
            # Находим максимальный общий диапазон (наибольшая группа логгеров с пересечением)
            best_group = []
            best_range = None
            best_group_size = 0
            
            # Перебираем все возможные группы логгеров, начиная с наибольших
            for group_size in range(valid_count, 0, -1):
                # Генерируем все комбинации размера group_size
                for group_indices in combinations(range(valid_count), group_size):
                    group = [valid_files[i] for i in group_indices]
                    cstart = max(f[1] for f in group)
                    cend = min(f[2] for f in group)
                    if cstart < cend:  # Есть пересечение
                        if len(group) > best_group_size:
                            best_group = group
                            best_range = (cstart, cend)
                            best_group_size = len(group)
                            break  # Нашли максимальную группу для этого размера
                if best_group:
                    break  # Нашли максимальную группу
            
            common_range_text = ""
            
            # Выводим максимальный общий диапазон
            if best_group and best_range:
                cstart, cend = best_range
                common_range_text = f"Максимальный общий временной диапазон (на основе {len(best_group)} из {total_loggers} логгеров):\n"
                common_range_text += f"{cstart.strftime('%d.%m.%Y %H:%M')} - {cend.strftime('%d.%m.%Y %H:%M')}\n\n"
                
                # Находим логгеры, не вошедшие в максимальный общий диапазон
                excluded = [f[0] for f in valid_files if f not in best_group]
                if excluded:
                    common_range_text += f"Логгеры, не вошедшие в максимальный общий диапазон: {', '.join(excluded)}\n\n"
            else:
                common_range_text = f"Максимальный общий временной диапазон отсутствует\n\n"
            
            # Выводим общий диапазон для всех логгеров
            if common_start_all < common_end_all:
                common_range_text += f"Общий временной диапазон (на основе {total_loggers} из {total_loggers} логгеров):\n"
                common_range_text += f"{common_start_all.strftime('%d.%m.%Y %H:%M')} - {common_end_all.strftime('%d.%m.%Y %H:%M')}\n"
            else:
                common_range_text += f"Общий временной диапазон (на основе {total_loggers} из {total_loggers} логгеров): отсутствует\n"

            if invalid_files:
                common_range_text += f"\nЛоггеры без диапазона: {', '.join(f[0] for f in invalid_files)}\n"

            self.ranges_text.insert("1.0", common_range_text)

        except Exception as e:
            self.ranges_text.insert("1.0", f"Ошибка анализа диапазонов: {str(e)}")

        self.ranges_text.config(state=tk.DISABLED)

    def save_data(self):
        """Сохранение данных"""
        import sqlite3
        import json
        settings_db_path = self.session_manager.get_settings_db_path()
        conn = sqlite3.connect(settings_db_path)
        cursor = conn.cursor()
        
        # Сохраняем тип отчета
        cursor.execute("""
            INSERT OR REPLACE INTO settings (key, value)
            VALUES (?, ?)
        """, ('report_type', self.report_type.get()))
        
        # Сохраняем флаг использования влажности
        cursor.execute("""
            INSERT OR REPLACE INTO settings (key, value)
            VALUES (?, ?)
        """, ('use_humidity', str(self.use_humidity.get())))
        
        # Сохраняем список файлов
        files_str = ','.join(self.selected_files)
        cursor.execute("""
            INSERT OR REPLACE INTO settings (key, value)
            VALUES (?, ?)
        """, ('selected_files', files_str))
        
        # Сохраняем скриншоты логгеров (сериализуем список кортежей)
        screenshots_str = json.dumps(self.logger_screenshots)
        cursor.execute("""
            INSERT OR REPLACE INTO settings (key, value)
            VALUES (?, ?)
        """, ('logger_screenshots', screenshots_str))
        
        conn.commit()
        conn.close()
        
        # Показываем сообщение об успешном сохранении
        self.show_custom_message("Сохранение", "Данные сохранены успешно!")

    def show_custom_message(self, title, message):
        """Кастомное модальное окно без системного звука"""
        top = tk.Toplevel(self.parent)
        top.title(title)
        top.geometry("300x150")
        top.resizable(False, False)
        top.attributes("-topmost", True)
        
        top.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (300 // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (150 // 2)
        top.geometry(f"+{x}+{y}")
        
        label = tk.Label(top, text=message, font=("Arial", 12), padx=20, pady=30)
        label.pack()
        
        def close_window():
            top.destroy()
        
        ok_btn = tk.Button(top, text="ОК", command=close_window, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), padx=20, pady=5)
        ok_btn.pack(pady=5)
        
        top.grab_set()
        top.wait_window()

    def clear_data(self):
        """Очистка данных фрейма"""
        self.selected_files = []
        self.files_tree.delete(*self.files_tree.get_children())
        self.report_type.set("Объект хранения")
        self.use_humidity.set(False)
        self.logger_screenshots = []
        if hasattr(self, 'screenshots_listbox'):
            self._refresh_screenshots_listbox()
        if hasattr(self, 'screenshot_preview_label'):
            self.screenshot_preview_label.config(image="", text="Выберите скриншот")

    def pack(self, **kwargs):
        """Упаковка фрейма"""
        self.parent.pack(**kwargs)

    def pack_forget(self):
        """Скрытие фрейма"""
        self.parent.pack_forget()
