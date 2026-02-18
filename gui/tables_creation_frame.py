#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Фрейм создания таблиц
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import sqlite3
from collections import defaultdict
from pathlib import Path

from report_generation.report_generator import ReportGenerator
from data_processing.excel_processor import ExcelProcessor
from gui.clipboard_manager import setup_clipboard_manager


class TablesCreationFrame:
    """Фрейм для создания таблиц"""
    
    def __init__(self, parent, session_manager, main_window=None):
        self.parent = parent
        self.session_manager = session_manager
        self.main_window = main_window  # Ссылка на главное окно
        
        # Диапазоны
        self.temp_min = tk.StringVar()
        self.temp_max = tk.StringVar()
        self.humidity_min = tk.StringVar()
        self.humidity_max = tk.StringVar()
        
        # Периоды
        self.periods = []
        
        self.create_widgets()
    
    def create_widgets(self):
        """Создание виджетов фрейма"""
        # Инициализируем менеджер буфера обмена для этого фрейма
        self.clipboard_manager = setup_clipboard_manager(self.parent)
        
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
            return "break"  # Останавливаем распространение события
        canvas.bind("<MouseWheel>", _on_mousewheel)

        # Для Linux с Button-4 и Button-5
        def _on_button4(event):
            canvas.yview_scroll(-1, "units")
            return "break"  # Останавливаем распространение события
        def _on_button5(event):
            canvas.yview_scroll(1, "units")
            return "break"  # Останавливаем распространение события
        canvas.bind("<Button-4>", _on_button4)
        canvas.bind("<Button-5>", _on_button5)
        
        # Привязка глобальной прокрутки к canvas
        def _on_canvas_mousewheel(event):
            # Если это вертикальная прокрутка, передаем управление глобальной системе
            if hasattr(self, 'parent') and hasattr(self.parent, 'master'):
                # Имитируем событие на главном окне
                self.parent.master.event_generate("<MouseWheel>", delta=event.delta, when="tail")
            return "break"
        canvas.bind("<MouseWheel>", _on_canvas_mousewheel, add="+")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Заголовок
        title_label = tk.Label(
            scrollable_frame,
            text="Создание таблиц",
            font=("Arial", 18, "bold"),
            bg="#ecf0f1"
        )
        title_label.pack(pady=20)
        
        # Подсказка о назначении полей
        hint_label = tk.Label(
            scrollable_frame,
            text="Внимание: температуру и влажность вводите для создания таблиц:\nТаблица 2 – Результаты картирования по температуре в зоне хранения лекарственных средств\nТаблица 3 – Результаты картирования по относительной влажности в зоне хранения лекарственных средств",
            font=("Arial", 10),
            bg="#ecf0f1",
            fg="#2c3e50",
            justify=tk.CENTER,
            wraplength=600
        )
        hint_label.pack(pady=5)
        
        # Диапазоны температур и влажности
        ranges_frame = tk.LabelFrame(
            scrollable_frame,
            text="Диапазоны температур и влажности",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        ranges_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Температура
        temp_frame = tk.Frame(ranges_frame, bg="#ecf0f1")
        temp_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            temp_frame,
            text="Температура:",
            font=("Arial", 11),
            bg="#ecf0f1",
            width=15,
            anchor=tk.W
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(temp_frame, text="от", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        temp_min_entry = tk.Entry(temp_frame, textvariable=self.temp_min, width=10, font=("Arial", 10))
        temp_min_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(temp_min_entry)
        
        tk.Label(temp_frame, text="до", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        temp_max_entry = tk.Entry(temp_frame, textvariable=self.temp_max, width=10, font=("Arial", 10))
        temp_max_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(temp_max_entry)
        
        tk.Label(temp_frame, text="°C", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        
        # Влажность
        humidity_frame = tk.Frame(ranges_frame, bg="#ecf0f1")
        humidity_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            humidity_frame,
            text="Влажность:",
            font=("Arial", 11),
            bg="#ecf0f1",
            width=15,
            anchor=tk.W
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(humidity_frame, text="от", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        humidity_min_entry = tk.Entry(humidity_frame, textvariable=self.humidity_min, width=10, font=("Arial", 10))
        humidity_min_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(humidity_min_entry)
        
        tk.Label(humidity_frame, text="до", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        humidity_max_entry = tk.Entry(humidity_frame, textvariable=self.humidity_max, width=10, font=("Arial", 10))
        humidity_max_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(humidity_max_entry)
        
        tk.Label(humidity_frame, text="%", font=("Arial", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        
        # Периоды
        periods_frame = tk.LabelFrame(
            scrollable_frame,
            text="Периоды",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        periods_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Кнопка добавления периода
        add_period_btn = tk.Button(
            periods_frame,
            text="+ Добавить период",
            command=self.add_period,
            bg="#3498db",
            fg="white",
            font=("Arial", 11),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        add_period_btn.pack(pady=10)
        
        # Список периодов
        periods_list_frame = tk.Frame(periods_frame, bg="#ecf0f1")
        periods_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        columns = ("Название", "Начало", "Окончание")
        self.periods_tree = ttk.Treeview(periods_list_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.periods_tree.heading(col, text=col)
            self.periods_tree.column(col, width=200)
        
        scrollbar_periods = ttk.Scrollbar(periods_list_frame, orient=tk.VERTICAL, command=self.periods_tree.yview)
        self.periods_tree.configure(yscrollcommand=scrollbar_periods.set)
        
        # Привязка прокрутки колесиком мыши и тачпадом к Treeview
        def _on_tree_mousewheel(event):
            # Если Treeview активен и в нем идет работа - обрабатываем прокрутку внутри поля
            if self.periods_tree.focus_get() == self.periods_tree:
                # Обрабатываем разные значения delta для мыши и тачпада
                delta = event.delta
                if abs(delta) > 120:  # Тачпад часто дает большие значения
                    delta = delta // 10  # Нормализуем
                self.periods_tree.yview_scroll(int(-1 * (delta / 120)), "units")
                return "break"  # Останавливаем распространение события
            # Если Treeview не активен - не перехватываем прокрутку, пусть всплывает к глобальному обработчику
        self.periods_tree.bind("<MouseWheel>", _on_tree_mousewheel)

        # Для Linux с Button-4 и Button-5 (тачпад)
        def _on_button4(event):
            if self.periods_tree.focus_get() == self.periods_tree:
                self.periods_tree.yview_scroll(-1, "units")
                return "break"  # Останавливаем распространение события
        def _on_button5(event):
            if self.periods_tree.focus_get() == self.periods_tree:
                self.periods_tree.yview_scroll(1, "units")
                return "break"  # Останавливаем распространение события
        self.periods_tree.bind("<Button-4>", _on_button4)
        self.periods_tree.bind("<Button-5>", _on_button5)

        # Дополнительные события для тачпадов
        def _on_button_press_4(event):
            if self.periods_tree.focus_get() == self.periods_tree:
                self.periods_tree.yview_scroll(-1, "units")
        def _on_button_press_5(event):
            if self.periods_tree.focus_get() == self.periods_tree:
                self.periods_tree.yview_scroll(1, "units")
        self.periods_tree.bind("<ButtonPress-4>", _on_button_press_4)
        self.periods_tree.bind("<ButtonPress-5>", _on_button_press_5)
        
        self.periods_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_periods.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки управления периодами
        buttons_frame = tk.Frame(periods_frame, bg="#ecf0f1")
        buttons_frame.pack(pady=5)

        # Кнопка редактирования периода
        edit_period_btn = tk.Button(
            buttons_frame,
            text="Редактировать выбранный",
            command=self.edit_period,
            bg="#f39c12",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        edit_period_btn.pack(side=tk.LEFT, padx=5)

        # Кнопка удаления периода
        remove_period_btn = tk.Button(
            buttons_frame,
            text="Удалить выбранный",
            command=self.remove_period,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        remove_period_btn.pack(side=tk.LEFT, padx=5)
        
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
        save_btn.pack(pady=10)
        
        # Кнопка генерации отчета
        generate_btn = tk.Button(
            scrollable_frame,
            text="Сгенерировать отчет",
            command=self.generate_report,
            bg="#9b59b6",
            fg="white",
            font=("Arial", 14, "bold"),
            padx=40,
            pady=15,
            cursor="hand2"
        )
        generate_btn.pack(pady=20)
    
    def add_period(self, edit_index=None):
        """Добавление или редактирование периода"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("Добавить период" if edit_index is None else "Редактировать период")
        dialog.geometry("700x320")  # Еще меньше высота
        dialog.transient(self.parent)
        dialog.grab_set()

        # Получаем данные для редактирования
        edit_data = None
        if edit_index is not None:
            edit_data = self.periods[edit_index]

        # Название периода
        name_frame = tk.Frame(dialog, bg="#ecf0f1")
        name_frame.pack(fill=tk.X, padx=20, pady=2)

        tk.Label(name_frame, text="Название периода:", font=("Arial", 11), bg="#ecf0f1", width=25, anchor=tk.W).pack(side=tk.LEFT)
        name_entry = tk.Entry(name_frame, width=40, font=("Arial", 10))
        if edit_data:
            name_entry.insert(0, edit_data['name'])
        name_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(name_entry)

        # Дата и время начала
        start_frame = tk.Frame(dialog, bg="#ecf0f1")
        start_frame.pack(fill=tk.X, padx=20, pady=2)

        tk.Label(start_frame, text="Дата и время начала:", font=("Arial", 11), bg="#ecf0f1", width=25, anchor=tk.W).pack(side=tk.LEFT)
        start_entry = tk.Entry(start_frame, width=40, font=("Arial", 10))
        if edit_data:
            start_dt = datetime.strptime(edit_data['start'], "%Y-%m-%d %H:%M:%S")
            start_entry.insert(0, start_dt.strftime("%d.%m.%Y %H:%M"))
        start_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(start_entry)

        # Кнопка очистки
        clear_start_btn = tk.Button(
            start_frame,
            text="✕",
            command=lambda: start_entry.delete(0, tk.END),
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            width=2,
            cursor="hand2"
        )
        clear_start_btn.pack(side=tk.LEFT, padx=5)

        # Текст подсказки справа
        start_hint_text = tk.Text(
            start_frame,
            height=1,
            width=18,
            font=("Arial", 8),
            wrap=tk.WORD,
            bg="#f8f9fa",
            fg="#495057",
            relief=tk.FLAT,
            borderwidth=1
        )
        start_hint_text.insert("1.0", "22.01.2026 10:00")
        start_hint_text.config(state=tk.DISABLED)
        start_hint_text.pack(side=tk.LEFT, padx=5)

        # Дата и время окончания
        end_frame = tk.Frame(dialog, bg="#ecf0f1")
        end_frame.pack(fill=tk.X, padx=20, pady=2)

        tk.Label(end_frame, text="Дата и время окончания:", font=("Arial", 11), bg="#ecf0f1", width=25, anchor=tk.W).pack(side=tk.LEFT)
        end_entry = tk.Entry(end_frame, width=40, font=("Arial", 10))
        if edit_data:
            end_dt = datetime.strptime(edit_data['end'], "%Y-%m-%d %H:%M:%S")
            end_entry.insert(0, end_dt.strftime("%d.%m.%Y %H:%M"))
        end_entry.pack(side=tk.LEFT, padx=5)
        
        # Добавляем контекстное меню для виджетов ввода
        self.clipboard_manager.create_context_menu(end_entry)

        # Кнопка очистки
        clear_end_btn = tk.Button(
            end_frame,
            text="✕",
            command=lambda: end_entry.delete(0, tk.END),
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            width=2,
            cursor="hand2"
        )
        clear_end_btn.pack(side=tk.LEFT, padx=5)

        # Текст подсказки справа
        end_hint_text = tk.Text(
            end_frame,
            height=1,
            width=18,
            font=("Arial", 8),
            wrap=tk.WORD,
            bg="#f8f9fa",
            fg="#495057",
            relief=tk.FLAT,
            borderwidth=1
        )
        end_hint_text.insert("1.0", "22.01.2026 18:00")
        end_hint_text.config(state=tk.DISABLED)
        end_hint_text.pack(side=tk.LEFT, padx=5)



        def save_period():
            name = name_entry.get().strip()
            start = start_entry.get().strip()
            end = end_entry.get().strip()

            if not all([name, start, end]):
                messagebox.showwarning("Предупреждение", "Заполните обязательные поля")
                return

            # Проверяем формат даты
            try:
                start_dt = datetime.strptime(start, "%d.%m.%Y %H:%M")
                end_dt = datetime.strptime(end, "%d.%m.%Y %H:%M")

                if end_dt <= start_dt:
                    messagebox.showerror("Ошибка", "Дата окончания должна быть позже даты начала")
                    return

                period_data = {
                    'name': name,
                    'start': start_dt.strftime("%Y-%m-%d %H:%M:%S"),
                    'end': end_dt.strftime("%Y-%m-%d %H:%M:%S")
                }

                if edit_index is not None:
                    # Редактирование существующего периода
                    self.periods[edit_index] = period_data
                    # Обновляем дерево
                    items = self.periods_tree.get_children()
                    self.periods_tree.delete(items[edit_index])
                    self.periods_tree.insert("", edit_index, values=(name, start, end))
                else:
                    # Добавление нового периода
                    self.periods.append(period_data)
                    self.periods_tree.insert("", tk.END, values=(name, start, end))

                dialog.destroy()

            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты. Используйте: дд.мм.гггг чч:мм")

        tk.Button(
            dialog,
            text="Сохранить",
            command=save_period,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5,
            cursor="hand2"
        ).pack(pady=15)

    def edit_period(self):
        """Редактирование выбранного периода"""
        selection = self.periods_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите период для редактирования")
            return

        index = self.periods_tree.index(selection[0])
        self.add_period(edit_index=index)
    
    def remove_period(self):
        """Удаление периода"""
        selection = self.periods_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите период для удаления")
            return
        
        for item in selection:
            index = self.periods_tree.index(item)
            self.periods_tree.delete(item)
            del self.periods[index]
    
    def save_data(self, silent: bool = False):
        """Сохранение данных"""
        try:
            temp_min_val = float(self.temp_min.get()) if self.temp_min.get() else None
            temp_max_val = float(self.temp_max.get()) if self.temp_max.get() else None
            humidity_min_val = float(self.humidity_min.get()) if self.humidity_min.get() else None
            humidity_max_val = float(self.humidity_max.get()) if self.humidity_max.get() else None
        except ValueError:
            messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений")
            return
        
        if not self.periods:
            messagebox.showwarning("Предупреждение", "Добавьте хотя бы один период")
            return
        
        # Сохраняем периоды в БД
        periods_db_path = self.session_manager.get_periods_db_path()
        conn = sqlite3.connect(periods_db_path)
        cursor = conn.cursor()
        
        # Очищаем старые данные
        cursor.execute("DELETE FROM periods")
        
        # Добавляем новые периоды
        for period in self.periods:
            cursor.execute("""
                INSERT INTO periods (start_time, end_time, name, loggers, required_mode_from, required_mode_to)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                period['start'],
                period['end'],
                period['name'],
                period.get('excluded_loggers', ''),
                temp_min_val,
                temp_max_val
            ))
        
        conn.commit()
        conn.close()
        
        if not silent:
            self.show_custom_message("Сохранение", "Данные сохранены успешно!")

    def clear_data(self):
        """Очистка данных фрейма"""
        self.temp_min.set('')
        self.temp_max.set('')
        self.humidity_min.set('')
        self.humidity_max.set('')
        self.periods = []
        self.periods_tree.delete(*self.periods_tree.get_children())

    def generate_report(self):
        """Генерация отчета"""
        # Проверяем наличие периодов
        if not self.periods:
            messagebox.showwarning("Предупреждение", "Добавьте хотя бы один период")
            return
        
        # Проверяем наличие загруженных Excel файлов
        project_mgmt = self.get_project_management_frame()
        if not project_mgmt or not project_mgmt.selected_files:
            messagebox.showwarning("Предупреждение", "Загрузите Excel файлы в разделе 'Управление проектом'")
            return
        
        # Запрашиваем название файла
        dialog = tk.Toplevel(self.parent)
        dialog.title("Генерация отчета")
        dialog.geometry("500x150")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        tk.Label(
            dialog,
            text="Введите название файла отчета:",
            font=("Arial", 11)
        ).pack(pady=20)
        
        filename_entry = tk.Entry(dialog, width=40, font=("Arial", 11))
        filename_entry.pack(pady=10)
        filename_entry.insert(0, "Отчет_по_картированию")
        filename_entry.focus()
        
        def generate():
            filename = filename_entry.get().strip()
            if not filename:
                messagebox.showwarning("Предупреждение", "Введите название файла")
                return
            
            # Убираем расширение если есть
            if filename.endswith('.docx'):
                filename = filename[:-5]
            
            # Сохраняем текущие данные (в том числе диапазоны температур в БД)
            # без показа окна "Данные сохранены успешно"
            self.save_data(silent=True)
            
            # Получаем данные из других разделов
            other_info_frame = self.get_other_info_frame()
            key_elements_frame = self.get_key_elements_frame()
            
            # Диапазоны температур и влажности из полей ввода
            try:
                temp_min_val = float(self.temp_min.get()) if self.temp_min.get() else None
                temp_max_val = float(self.temp_max.get()) if self.temp_max.get() else None
                humidity_min_val = float(self.humidity_min.get()) if self.humidity_min.get() else None
                humidity_max_val = float(self.humidity_max.get()) if self.humidity_max.get() else None
            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений диапазонов")
                return

            # Подготавливаем данные для отчета
            other_info = {}
            if other_info_frame:
                other_info = {
                    'mapping_results': other_info_frame.mapping_results_widget.get("1.0", tk.END).strip(),
                    'conclusion': other_info_frame.conclusion.get(),
                    'risk_areas': other_info_frame.risk_areas,
                    'layout_image': other_info_frame.images.get('layout'),
                    'loggers_image': other_info_frame.images.get('loggers'),
                    'temp_map_image': other_info_frame.images.get('temp_map'),
                    'humidity_map_image': other_info_frame.images.get('humidity_map'),
                    'contract_date': datetime.now().strftime("%d.%m.%Y"),
                    'humidity_min': humidity_min_val,
                    'humidity_max': humidity_max_val,
                    'selected_recommendations': other_info_frame.get_selected_recommendations(),
                    'image_orientations': {
                        'layout': other_info_frame.get_image_orientation('layout'),
                        'loggers': other_info_frame.get_image_orientation('loggers'),
                        'temp_map': other_info_frame.get_image_orientation('temp_map'),
                        'humidity_map': other_info_frame.get_image_orientation('humidity_map')
                    }
                }
            if project_mgmt:
                other_info['use_humidity'] = project_mgmt.use_humidity.get()
                other_info['logger_screenshots'] = getattr(project_mgmt, 'logger_screenshots', [])
            
            # Определяем тип отчета
            report_type = "Объект хранения"
            if project_mgmt:
                report_type = project_mgmt.report_type.get()
            
            use_humidity = False
            if project_mgmt:
                use_humidity = project_mgmt.use_humidity.get()
            
            # Определяем путь к шаблону
            templates_dir = Path(self.session_manager.project_root) / "temp"
            template_mapping = {
                "Объект хранения": "template3.docx",
                "Зона хранения": "template4.docx",
                "Холодильник/Морозильник": "template5.docx"
            }
            
            template_name = template_mapping.get(report_type, "template3.docx")
            template_path = templates_dir / template_name
            
            if not template_path.exists():
                messagebox.showerror(
                    "Ошибка",
                    f"Шаблон {template_name} не найден в папке temp/"
                )
                dialog.destroy()
                return
            
            # Запрашиваем место сохранения
            output_path = filedialog.asksaveasfilename(
                title="Сохранить отчет",
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                initialfile=f"{filename}.docx"
            )
            
            if not output_path:
                dialog.destroy()
                return
            
            dialog.destroy()
            
            # Обрабатываем Excel файлы и сохраняем статистику
            try:
                excel_processor = ExcelProcessor(self.session_manager)
                logger_data = excel_processor.process_excel_files(project_mgmt.selected_files)

                # Расчёт однородности температуры и влажности во времени
                def compute_homogeneity(data_dict, value_key):
                    time_values = defaultdict(list)
                    for data in data_dict.values():
                        times = data.get('times') or []
                        values = data.get(value_key) or []
                        n = min(len(times), len(values))
                        for t, v in zip(times[:n], values[:n]):
                            time_values[t].append(v)

                    max_diff = None
                    max_times = []
                    for t, values in time_values.items():
                        if len(values) < 2:
                            continue
                        diff = max(values) - min(values)
                        if max_diff is None or diff > max_diff + 1e-9:
                            max_diff = diff
                            max_times = [t]
                        elif abs(diff - max_diff) <= 1e-9:
                            max_times.append(t)
                    return max_diff, max_times

                def format_time_value(t):
                    if isinstance(t, datetime):
                        return t.strftime("%d.%m.%Y %H:%M")
                    return str(t)

                temp_hom_value, temp_hom_times_raw = compute_homogeneity(logger_data, 'temperatures')
                hum_hom_value, hum_hom_times_raw = compute_homogeneity(logger_data, 'humidities')

                if temp_hom_value is not None:
                    temp_hom_text = f"{temp_hom_value:.2f} (" + ", ".join(
                        format_time_value(t) for t in temp_hom_times_raw
                    ) + ")"
                else:
                    temp_hom_text = "—"

                if hum_hom_value is not None:
                    hum_hom_text = f"{hum_hom_value:.2f} (" + ", ".join(
                        format_time_value(t) for t in hum_hom_times_raw
                    ) + ")"
                else:
                    hum_hom_text = "—"

                # Сохраняем значения однородности в other_info
                other_info['temp_homogeneity_text'] = temp_hom_text
                other_info['hum_homogeneity_text'] = hum_hom_text

                # Сохраняем статистику логгеров для каждого периода
                periods_db_path = self.session_manager.get_periods_db_path()
                conn = sqlite3.connect(periods_db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM periods")
                period_ids = [row[0] for row in cursor.fetchall()]
                conn.close()
                
                for period_id in period_ids:
                    excel_processor.save_logger_stats(logger_data, period_id)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка обработки Excel файлов:\n{e}")
                import traceback
                traceback.print_exc()
                return
            
            # Генерируем отчет
            generator = ReportGenerator(self.session_manager)
            success = generator.generate_report(
                report_type=report_type,
                template_path=str(template_path),
                output_path=output_path,
                use_humidity=use_humidity,
                other_info=other_info,
                periods=None
            )
            
            if success:
                messagebox.showinfo("Успех", f"Отчет успешно создан:\n{output_path}")
            else:
                messagebox.showerror("Ошибка", "Не удалось создать отчет. Проверьте консоль для подробностей.")
        
        tk.Button(
            dialog,
            text="Сгенерировать",
            command=generate,
            bg="#9b59b6",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=5,
            cursor="hand2"
        ).pack(pady=10)
    
    def get_project_management_frame(self):
        """Получить ссылку на фрейм управления проектом"""
        if self.main_window and hasattr(self.main_window, 'frames'):
            return self.main_window.frames.get("project_management")
        return None
    
    def get_other_info_frame(self):
        """Получить ссылку на фрейм прочей информации"""
        if self.main_window and hasattr(self.main_window, 'frames'):
            return self.main_window.frames.get("other_info")
        return None
    
    def get_key_elements_frame(self):
        """Получить ссылку на фрейм ключевых элементов"""
        if self.main_window and hasattr(self.main_window, 'frames'):
            return self.main_window.frames.get("key_elements")
        return None
    
    def pack(self, **kwargs):
        """Упаковка фрейма"""
        self.parent.pack(**kwargs)
    
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
