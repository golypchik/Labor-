#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Фрейм прочей информации
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from PIL import Image, ImageTk


class OtherInfoFrame:
    """Фрейм для прочей информации"""

    def __init__(self, parent, session_manager):
        self.parent = parent
        self.session_manager = session_manager

        # Данные
        self.mapping_results = tk.StringVar()
        self.conclusion = tk.StringVar()
        self.risk_areas = []

        # Изображения
        self.images = {
            'layout': None,
            'loggers': None,
            'temp_map': None,
            'humidity_map': None
        }

        # Подсказки для выпадающих списков
        self.mapping_hints = [
            "Температура и относительная влажность воздуха не выходили за установленные пределы во всех зонах хранения лекарственных средств.",
            "Температура вышла за установленные пределы во всех зонах хранения лекарственных средств. Относительная влажность воздуха не выходила за установленные пределы во всех зонах хранения лекарственных средств.",
            "Температура вышла за установленные пределы в зонах расположения логгеров ... Относительная влажность воздуха не выходила за установленные пределы во всех зонах хранения лекарственных средств.",
            "Температура вышла за установленные пределы в зоне расположения логгера ... Относительная влажность воздуха не выходила за установленные пределы во всех зонах хранения лекарственных средств."
        ]


        # Рекомендации по типам отчетов
        self.recommendations = {
            "Объект хранения": [
                "Для снижения уровня относительной влажности воздуха рекомендуется установить осушитель с автоматическим контролем влажности, проверить герметичность окон, дверей и вентиляции для исключения поступления влажного воздуха.",
                "Организовать регулярный мониторинг влажности с фиксацией показателей и разработать план корректирующих мероприятий при отклонении от установленных значений.",
                "До устранения выявленных отклонений использование зоны хранения для лекарственных средств не рекомендуется. Влагочувствительную продукцию следует временно разместить в помещениях, соответствующих установленным критериям влажностного режима (не более 60%).",
                "Провести повторное картирование для оценки эффективности принятых корректирующих мер.",
                "Установить датчики температуры и (или) влажности для рутинного мониторинга в зонах выявленных холодной и горячей точек (см. Приложение 3).",
                "Температурно-влажностное картирование объекта хранения лекарственных средств следует проводить в летний и зимний периоды с периодичностью не реже одного раза в три года.",
                "В случае внесения инженерно-конструкторских изменений, необходимо проведение внепланового картирования."
            ],
            "Зона хранения": [
                "Для стабилизации температурного режима рекомендуется установить кондиционер с функцией охлаждения и автоматическим контролем температуры, провести проверку вентиляции и исключить влияние внутренних источников тепла.",
                "Организовать регулярный мониторинг температуры с фиксацией показателей и разработать план корректирующих мероприятий при отклонении от установленных значений.",
                "До устранения выявленных отклонений использование зоны хранения для лекарственных средств не рекомендуется. Термолабильную продукцию следует временно разместить в помещениях, соответствующих установленным критериям температурного режима (+15℃…+25℃).",
                "Провести повторное картирование для оценки эффективности принятых корректирующих мер.",
                "Температурно-влажностное картирование объекта хранения лекарственных средств следует проводить в летний и зимний периоды с периодичностью не реже одного раза в три года.",
                "В случае внесения инженерно-конструкторских изменений, необходимо проведение внепланового картирования."
            ],
            "Холодильник/Морозильник": [
                "Не хранить термолабильную продукцию в зонах выхода температуры.",
                "Провести стандартные процедуры технического обслуживания: разморозку, санитарную обработку внутренних поверхностей, очистку системы отвода талой воды и другие процедуры в соответствии с инструкцией по эксплуатации.",
                "Проверить исправность терморегулятора, уплотнителей и системы охлаждения, убедиться в отсутствии перегрузки камеры и нарушений условий эксплуатации. При необходимости обратиться в сервисную службу для диагностики и устранения неисправностей.",
                "Провести повторное температурное картирование для оценки эффективности принятых корректирующих мер.",
                "При размещении продукции обеспечивать зазор между стенками камеры и самой продукцией для свободной циркуляции воздуха.",
                "Не рекомендуется размещать продукцию на дне холодильной камеры.",
                "В случае невозможности стабилизации температурного режима, а также при достижении предельного срока службы оборудования – рассмотреть его замену на новое, соответствующее требованиям хранения.",
                "Температурное картирование объекта хранения лекарственных средств следует проводить с периодичностью не реже одного раза в три года.",
                "В случае внесения инженерно-конструкторских изменений, необходимо провести внеплановое температурное картирование.",
                "Вести журнал мониторинга температурных показателей с фиксацией отклонений и принятых мер."
            ]
        }

        # Рабочий список рекомендаций (программа + пользовательские, порядок задаётся пользователем)
        self.recommendations_list = []  # Хранит текущий список для отчёта

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
            text="Прочая информация",
            font=("Arial", 18, "bold"),
            bg="#ecf0f1"
        )
        title_label.pack(pady=20)
        
        # Текстовые поля
        text_fields_frame = tk.LabelFrame(
            scrollable_frame,
            text="Общая информация",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        text_fields_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Итоги картирования
        mapping_container = tk.Frame(text_fields_frame, bg="#ecf0f1")
        mapping_container.pack(fill=tk.X, pady=5)

        tk.Label(
            mapping_container,
            text="Итоги картирования:",
            font=("Arial", 11),
            bg="#ecf0f1",
            width=30,
            anchor=tk.W
        ).pack(side=tk.LEFT, padx=10)

        # Многострочное поле ввода (сначала)
        mapping_text_frame = tk.Frame(mapping_container, bg="#ecf0f1")
        mapping_text_frame.pack(side=tk.LEFT, padx=5)

        tk.Label(
            mapping_text_frame,
            text="Текст:",
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#666"
        ).pack(anchor=tk.W)

        # Контейнер для текста с прокруткой
        text_container = tk.Frame(mapping_text_frame, bg="#ecf0f1")
        text_container.pack()

        mapping_text = tk.Text(
            text_container,
            height=5,
            width=58,
            font=("Arial", 10),
            wrap=tk.WORD
        )

        # Вертикальная полоса прокрутки
        text_scrollbar = tk.Scrollbar(text_container, command=mapping_text.yview)
        mapping_text.configure(yscrollcommand=text_scrollbar.set)

        # Привязка прокрутки колесиком мыши
        def _on_text_mousewheel(event):
            mapping_text.yview_scroll(int(-1 * (event.delta / 120)), "units")
        mapping_text.bind("<MouseWheel>", _on_text_mousewheel)

        mapping_text.pack(side=tk.LEFT)
        text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Синхронизируем с StringVar
        def update_mapping_var(*args):
            self.mapping_results.set(mapping_text.get("1.0", "end-1c"))

        def update_mapping_text(*args):
            current = mapping_text.get("1.0", "end-1c")
            if current != self.mapping_results.get():
                mapping_text.delete("1.0", tk.END)
                mapping_text.insert("1.0", self.mapping_results.get())

        mapping_text.bind("<KeyRelease>", update_mapping_var)
        self.mapping_results.trace_add("write", update_mapping_text)

        self.mapping_results_widget = mapping_text

        # Кнопка очистки (ставим перед подсказками)
        clear_mapping_btn = tk.Button(
            mapping_container,
            text="✕",
            command=lambda: self.clear_field(self.mapping_results, mapping_text),
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            width=2,
            cursor="hand2"
        )
        clear_mapping_btn.pack(side=tk.LEFT, padx=5)

        # Выпадающий список с подсказками (после кнопки очистки)
        mapping_combo_frame = tk.Frame(mapping_container, bg="#ecf0f1")
        mapping_combo_frame.pack(side=tk.LEFT, padx=5)

        tk.Label(
            mapping_combo_frame,
            text="Подсказки:",
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#666"
        ).pack(anchor=tk.W)

        mapping_combo = ttk.Combobox(
            mapping_combo_frame,
            values=self.mapping_hints,
            font=("Arial", 9),
            width=60,
            state="readonly"  # Только выбор из списка
        )
        mapping_combo.pack()
        mapping_combo.bind("<<ComboboxSelected>>", lambda e: self.insert_hint(self.mapping_results_widget, mapping_combo.get()))
        
        # Заключение
        conclusion_container = tk.Frame(text_fields_frame, bg="#ecf0f1")
        conclusion_container.pack(fill=tk.X, pady=5)

        tk.Label(
            conclusion_container,
            text="Заключение:",
            font=("Arial", 11),
            bg="#ecf0f1",
            width=30,
            anchor=tk.W
        ).pack(side=tk.LEFT, padx=10)

        # Многострочное поле ввода (5 строк с прокруткой)
        conclusion_text_frame = tk.Frame(conclusion_container, bg="#ecf0f1")
        conclusion_text_frame.pack(side=tk.LEFT, padx=5)

        tk.Label(
            conclusion_text_frame,
            text="Текст:",
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#666"
        ).pack(anchor=tk.W)

        # Контейнер для текста с прокруткой
        conclusion_text_container = tk.Frame(conclusion_text_frame, bg="#ecf0f1")
        conclusion_text_container.pack()

        self.conclusion_text = tk.Text(
            conclusion_text_container,
            height=5,
            width=58,
            font=("Arial", 10),
            wrap=tk.WORD
        )

        # Вертикальная полоса прокрутки
        conclusion_text_scrollbar = tk.Scrollbar(conclusion_text_container, command=self.conclusion_text.yview)
        self.conclusion_text.configure(yscrollcommand=conclusion_text_scrollbar.set)

        # Привязка прокрутки колесиком мыши
        def _on_conclusion_text_mousewheel(event):
            self.conclusion_text.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.conclusion_text.bind("<MouseWheel>", _on_conclusion_text_mousewheel)

        self.conclusion_text.pack(side=tk.LEFT)
        conclusion_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Синхронизируем с StringVar
        def update_conclusion_var(*args):
            self.conclusion.set(self.conclusion_text.get("1.0", "end-1c"))

        def update_conclusion_text(*args):
            current = self.conclusion_text.get("1.0", "end-1c")
            if current != self.conclusion.get():
                self.conclusion_text.delete("1.0", tk.END)
                self.conclusion_text.insert("1.0", self.conclusion.get())

        self.conclusion_text.bind("<KeyRelease>", update_conclusion_var)
        self.conclusion.trace_add("write", update_conclusion_text)

        # Кнопка очистки (справа от поля)
        clear_conclusion_btn = tk.Button(
            conclusion_container,
            text="✕",
            command=lambda: self.clear_field(self.conclusion, self.conclusion_text),
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            width=2,
            cursor="hand2"
        )
        clear_conclusion_btn.pack(side=tk.LEFT, padx=5)

        # Подсказка справа от поля
        hint_frame = tk.Frame(conclusion_container, bg="#ecf0f1")
        hint_frame.pack(side=tk.LEFT, padx=10)

        tk.Label(
            hint_frame,
            text="Подсказка:",
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#666"
        ).pack(anchor=tk.W)

        hint_text = tk.Text(
            hint_frame,
            height=5,
            width=60,
            font=("Arial", 9),
            wrap=tk.WORD,
            bg="#f8f9fa",
            fg="#495057",
            relief=tk.FLAT,
            borderwidth=1
        )
        
        default_hint = "По результатам картирования зоны хранения лекарственных средств в кабинете старшей медсестры реабилитационного отделения №1 температурный режим превышал установленный диапазон (+15 ℃…+25 ℃), что свидетельствует о несоответствии данного параметра установленным критериям. Показатель относительной влажности воздуха на протяжении всего периода исследования находился в пределах допустимого значения (не более 60 %)."
        
        hint_text.insert("1.0", default_hint)
        hint_text.config(state=tk.DISABLED)
        
        hint_text.pack()
        
        # Загрузка изображений
        images_frame = tk.LabelFrame(
            scrollable_frame,
            text="Загрузка изображений",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        images_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Основной контейнер для изображений (слева - список, справа - превью)
        images_container = tk.Frame(images_frame, bg="#ecf0f1")
        images_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Левая колонка - список изображений
        left_column = tk.Frame(images_container, bg="#ecf0f1", width=400)
        left_column.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        left_column.pack_propagate(False)

        tk.Label(
            left_column,
            text="Изображения:",
            font=("Arial", 10, "bold"),
            bg="#ecf0f1"
        ).pack(anchor=tk.W, pady=5)

        # Список изображений
        self.images_listbox = tk.Listbox(left_column, height=8, font=("Arial", 9))
        images_scrollbar = ttk.Scrollbar(left_column, orient=tk.VERTICAL, command=self.images_listbox.yview)
        self.images_listbox.configure(yscrollcommand=images_scrollbar.set)

        self.images_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        images_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Заполняем список начальными элементами
        image_buttons = [
            ("Планировка зоны хранения", "layout"),
            ("Расстановка логгеров", "loggers"),
            ("Температурная карта", "temp_map"),
            ("Влажностная карта", "humidity_map")
        ]

        for text, key in image_buttons:
            self.images_listbox.insert(tk.END, f"{text} - Не загружено")

        # Привязываем выбор к показу превью
        self.images_listbox.bind('<<ListboxSelect>>', self.on_image_select)

        # Кнопки управления
        buttons_frame = tk.Frame(left_column, bg="#ecf0f1")
        buttons_frame.pack(fill=tk.X, pady=10)

        for text, key in image_buttons:
            btn_container = tk.Frame(buttons_frame, bg="#ecf0f1")
            btn_container.pack(fill=tk.X, pady=2)

            # Кнопка загрузки
            load_btn = tk.Button(
                btn_container,
                text="Загрузить",
                command=lambda k=key: self.load_image(k),
                bg="#3498db",
                fg="white",
                font=("Arial", 9),
                padx=8,
                pady=2,
                cursor="hand2"
            )
            load_btn.pack(side=tk.LEFT, padx=2)

            self.images[f"{key}_btn"] = load_btn

            # Кнопка удаления (✕)
            delete_btn = tk.Button(
                btn_container,
                text="✕",
                command=lambda k=key: self.delete_image(k),
                bg="#e74c3c",
                fg="white",
                font=("Arial", 10, "bold"),
                width=2,
                cursor="hand2"
            )
            delete_btn.pack(side=tk.RIGHT, padx=2)

            self.images[f"{key}_delete_btn"] = delete_btn

        # Правая колонка - превью
        right_column = tk.Frame(images_container, bg="#ecf0f1")
        right_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        tk.Label(
            right_column,
            text="Превью:",
            font=("Arial", 10, "bold"),
            bg="#ecf0f1"
        ).pack(anchor=tk.W, pady=5)

        # Область превью
        preview_canvas = tk.Canvas(right_column, bg="#f8f9fa", highlightthickness=1, highlightbackground="#bdc3c7")
        preview_scrollbar = ttk.Scrollbar(right_column, orient="vertical", command=preview_canvas.yview)
        preview_scrollable = tk.Frame(preview_canvas, bg="#f8f9fa")

        preview_scrollable.bind(
            "<Configure>",
            lambda e: preview_canvas.configure(scrollregion=preview_canvas.bbox("all"))
        )

        preview_canvas.create_window((0, 0), window=preview_scrollable, anchor="nw")
        preview_canvas.configure(yscrollcommand=preview_scrollbar.set)

        preview_canvas.pack(side="left", fill="both", expand=True)
        preview_scrollbar.pack(side="right", fill="y")

        # Привязка прокрутки
        def _on_mousewheel(event):
            preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        preview_canvas.bind("<MouseWheel>", _on_mousewheel)

        self.preview_frame = preview_scrollable
        self.preview_canvas = preview_canvas

        # Места рисков
        risks_frame = tk.LabelFrame(
            scrollable_frame,
            text="Места рисков",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        risks_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Список мест рисков
        risks_list_frame = tk.Frame(risks_frame, bg="#ecf0f1")
        risks_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.risks_listbox = tk.Listbox(risks_list_frame, height=6, font=("Arial", 10))
        scrollbar_risks = ttk.Scrollbar(risks_list_frame, orient=tk.VERTICAL, command=self.risks_listbox.yview)
        self.risks_listbox.configure(yscrollcommand=scrollbar_risks.set)

        self.risks_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_risks.pack(side=tk.RIGHT, fill=tk.Y)

        # Кнопки управления местами рисков
        risks_buttons_frame = tk.Frame(risks_frame, bg="#ecf0f1")
        risks_buttons_frame.pack(fill=tk.X, pady=10)

        add_risk_btn = tk.Button(
            risks_buttons_frame,
            text="+ Добавить место риска",
            command=self.add_risk_area,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        add_risk_btn.pack(side=tk.LEFT, padx=5)

        remove_risk_btn = tk.Button(
            risks_buttons_frame,
            text="Удалить выбранное",
            command=self.remove_risk_area,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        remove_risk_btn.pack(side=tk.LEFT, padx=5)

        # Рекомендации заказчику (единый список: удалять, добавлять свои, менять порядок)
        recommendations_frame = tk.LabelFrame(
            scrollable_frame,
            text="Рекомендации заказчику",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            padx=20,
            pady=15
        )
        recommendations_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        tk.Label(
            recommendations_frame,
            text="Редактируйте список: удаляйте ненужные, добавляйте свои, задайте порядок кнопками Вверх/Вниз.",
            font=("Arial", 9),
            bg="#ecf0f1",
            fg="#666"
        ).pack(anchor=tk.W, pady=(0, 5))

        # Контейнер: список с вертикальной и горизонтальной прокруткой
        rec_list_container = tk.Frame(recommendations_frame, bg="#ecf0f1")
        rec_list_container.pack(fill=tk.BOTH, expand=True, pady=5)

        rec_scrollbar_y = ttk.Scrollbar(rec_list_container)
        rec_scrollbar_x = ttk.Scrollbar(rec_list_container, orient=tk.HORIZONTAL)
        self.recommendations_listbox = tk.Listbox(
            rec_list_container, height=8, font=("Arial", 10),
            yscrollcommand=rec_scrollbar_y.set,
            xscrollcommand=rec_scrollbar_x.set
        )
        rec_scrollbar_y.config(command=self.recommendations_listbox.yview)
        rec_scrollbar_x.config(command=self.recommendations_listbox.xview)
        self.recommendations_listbox.grid(row=0, column=0, sticky="nsew")
        rec_scrollbar_y.grid(row=0, column=1, sticky="ns")
        rec_scrollbar_x.grid(row=1, column=0, sticky="ew")
        rec_list_container.grid_rowconfigure(0, weight=1)
        rec_list_container.grid_columnconfigure(0, weight=1)

        # Кнопки управления (под списком, всегда видны)
        rec_buttons_frame = tk.Frame(recommendations_frame, bg="#ecf0f1")
        rec_buttons_frame.pack(fill=tk.X, pady=(10, 0))

        update_rec_btn = tk.Button(
            rec_buttons_frame,
            text="Обновить рекомендации",
            command=self.update_recommendations,
            bg="#f39c12",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        update_rec_btn.pack(side=tk.LEFT, padx=3)

        add_rec_btn = tk.Button(
            rec_buttons_frame,
            text="+ Добавить",
            command=self.add_custom_recommendation,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        add_rec_btn.pack(side=tk.LEFT, padx=3)

        remove_rec_btn = tk.Button(
            rec_buttons_frame,
            text="Удалить",
            command=self.remove_recommendation,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        remove_rec_btn.pack(side=tk.LEFT, padx=3)

        move_up_btn = tk.Button(
            rec_buttons_frame,
            text="↑ Вверх",
            command=lambda: self.move_recommendation(-1),
            bg="#3498db",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        move_up_btn.pack(side=tk.LEFT, padx=3)

        move_down_btn = tk.Button(
            rec_buttons_frame,
            text="↓ Вниз",
            command=lambda: self.move_recommendation(1),
            bg="#3498db",
            fg="white",
            font=("Arial", 10),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        move_down_btn.pack(side=tk.LEFT, padx=3)

        # Кнопка сохранения (внизу, под рекомендациями)
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

    def _refresh_listbox_from_list(self):
        """Синхронизирует listbox с recommendations_list (полный текст)"""
        self.recommendations_listbox.delete(0, tk.END)
        for rec in self.recommendations_list:
            self.recommendations_listbox.insert(tk.END, rec)

    def move_recommendation(self, direction):
        """Перемещение рекомендации вверх или вниз"""
        selection = self.recommendations_listbox.curselection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите рекомендацию для перемещения")
            return
        current_index = selection[0]
        new_index = current_index + direction
        if 0 <= new_index < len(self.recommendations_list):
            self.recommendations_list[current_index], self.recommendations_list[new_index] = \
                self.recommendations_list[new_index], self.recommendations_list[current_index]
            self._refresh_listbox_from_list()
            self.recommendations_listbox.selection_set(new_index)

    def get_project_type(self):
        """Получение типа отчета из фрейма управления проектом"""
        project_mgmt = self.get_project_management_frame()
        if project_mgmt:
            return project_mgmt.report_type.get()
        return "Объект хранения"

    def get_project_management_frame(self):
        """Получить ссылку на фрейм управления проектом"""
        if hasattr(self, 'main_window') and self.main_window:
            return self.main_window.frames.get("project_management")
        return None

    def update_recommendations(self):
        """Загрузка рекомендаций программы по типу отчёта (заменяет текущий список)"""
        report_type = self.get_project_type()
        self.recommendations_list = self.recommendations.get(report_type, [])[:]
        self._refresh_listbox_from_list()

    def add_custom_recommendation(self):
        """Добавление своей рекомендации в список"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("Добавить рекомендацию")
        dialog.geometry("500x200")
        dialog.transient(self.parent)
        dialog.grab_set()
        tk.Label(dialog, text="Текст рекомендации:", font=("Arial", 11)).pack(pady=10)
        text_area = tk.Text(dialog, height=6, width=50, font=("Arial", 10))
        text_area.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        def save_recommendation():
            text = text_area.get("1.0", tk.END).strip()
            if text:
                self.recommendations_list.append(text)
                self._refresh_listbox_from_list()
                dialog.destroy()
            else:
                messagebox.showwarning("Предупреждение", "Введите текст рекомендации")
        tk.Button(dialog, text="Сохранить", command=save_recommendation,
                  bg="#27ae60", fg="white", font=("Arial", 10), padx=20, pady=5, cursor="hand2").pack(pady=10)

    def remove_recommendation(self):
        """Удаление выбранной рекомендации (можно удалять любую, в т.ч. от программы)"""
        selection = self.recommendations_listbox.curselection()
        if selection:
            index = selection[0]
            if 0 <= index < len(self.recommendations_list):
                del self.recommendations_list[index]
                self._refresh_listbox_from_list()
        else:
            messagebox.showwarning("Предупреждение", "Выберите рекомендацию для удаления")

    def on_image_select(self, event):
        """Обработчик выбора изображения в списке"""
        selection = self.images_listbox.curselection()
        if selection:
            index = selection[0]
            # Определяем ключ изображения по индексу
            image_keys = ['layout', 'loggers', 'temp_map', 'humidity_map']
            if index < len(image_keys):
                key = image_keys[index]
                self.show_image_preview(key)

    def show_image_preview(self, key):
        """Показ превью выбранного изображения"""
        # Очищаем превью
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        if self.images[key]:
            try:
                # Создаем миниатюру большего размера для превью
                img = Image.open(self.images[key])
                # Сохраняем пропорции, но ограничиваем размер
                img.thumbnail((350, 350), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)

                # Отображаем превью
                preview_label = tk.Label(
                    self.preview_frame,
                    image=photo,
                    bg="#f8f9fa"
                )
                preview_label.image = photo  # Сохраняем ссылку на изображение
                preview_label.pack(pady=10)

                # Название файла
                name_label = tk.Label(
                    self.preview_frame,
                    text=Path(self.images[key]).name,
                    font=("Arial", 9),
                    bg="#f8f9fa",
                    fg="#333"
                )
                name_label.pack(pady=5)

                # Сохраняем ссылку на фото в словаре изображений
                self.images[f"{key}_preview_photo"] = photo

            except Exception as e:
                # Показываем сообщение об ошибке
                error_label = tk.Label(
                    self.preview_frame,
                    text=f"Ошибка загрузки превью:\n{str(e)}",
                    font=("Arial", 9),
                    bg="#f8f9fa",
                    fg="red",
                    justify=tk.LEFT
                )
                error_label.pack(pady=20)
        else:
            # Показываем сообщение, что изображение не загружено
            no_image_label = tk.Label(
                self.preview_frame,
                text="Изображение не загружено",
                font=("Arial", 10),
                bg="#f8f9fa",
                fg="#666"
            )
            no_image_label.pack(pady=20)


    
    def load_image(self, key):
        """Загрузка изображения"""
        file_path = filedialog.askopenfilename(
            title=f"Выберите изображение для {key}",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
        )

        if file_path:
            try:
                # Проверяем, что это изображение
                img = Image.open(file_path)
                self.images[key] = file_path

                # Обновляем список изображений
                image_titles = {
                    'layout': 'Планировка зоны хранения',
                    'loggers': 'Расстановка логгеров',
                    'temp_map': 'Температурная карта',
                    'humidity_map': 'Влажностная карта'
                }

                title = image_titles.get(key, key)
                # Находим индекс в списке
                for i in range(self.images_listbox.size()):
                    if title in self.images_listbox.get(i):
                        self.images_listbox.delete(i)
                        self.images_listbox.insert(i, f"{title} - Загружено")
                        break

                # Показываем кнопку удаления
                if f"{key}_delete_btn" in self.images:
                    self.images[f"{key}_delete_btn"].pack(side=tk.RIGHT, padx=2)

            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить изображение:\n{e}")
    
    def add_risk_area(self):
        """Добавление места риска"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("Добавить место риска")
        dialog.geometry("500x300")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        tk.Label(
            dialog,
            text="Описание места риска:",
            font=("Arial", 11)
        ).pack(pady=10)
        
        text_area = tk.Text(dialog, height=10, width=50, font=("Arial", 10))
        text_area.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        def save_risk():
            description = text_area.get("1.0", tk.END).strip()
            if description:
                self.risk_areas.append(description)
                self.risks_listbox.insert(tk.END, description[:50] + ("..." if len(description) > 50 else ""))
                dialog.destroy()
            else:
                messagebox.showwarning("Предупреждение", "Введите описание места риска")
        
        tk.Button(
            dialog,
            text="Сохранить",
            command=save_risk,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5,
            cursor="hand2"
        ).pack(pady=10)
    
    def remove_risk_area(self):
        """Удаление места риска"""
        selection = self.risks_listbox.curselection()
        if selection:
            index = selection[0]
            self.risks_listbox.delete(index)
            del self.risk_areas[index]
    
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

    def save_data(self):
        """Сохранение данных"""
        # Данные уже синхронизированы через StringVar
        self.show_custom_message("Сохранение", "Данные сохранены успешно!")

    def get_selected_recommendations(self):
        """Получение списка рекомендаций для отчёта (в порядке отображения)"""
        return self.recommendations_list.copy()

    def clear_data(self):
        """Очистка данных фрейма"""
        self.mapping_results.set('')
        self.conclusion.set('')
        self.risk_areas = []
        self.recommendations_list = []
        self._refresh_listbox_from_list()
        self.images['layout'] = None
        self.images['loggers'] = None
        self.images['temp_map'] = None
        self.images['humidity_map'] = None
        self.mapping_results_widget.delete("1.0", tk.END)
        self.conclusion_text.delete("1.0", tk.END)
        self.risks_listbox.delete(0, tk.END)
        # Обновляем список изображений
        image_titles = {
            'layout': 'Планировка зоны хранения',
            'loggers': 'Расстановка логгеров',
            'temp_map': 'Температурная карта',
            'humidity_map': 'Влажностная карта'
        }
        self.images_listbox.delete(0, tk.END)
        for text, key in image_titles.items():
            self.images_listbox.insert(tk.END, f"{text} - Не загружено")

    def pack(self, **kwargs):
        """Упаковка фрейма"""
        self.parent.pack(**kwargs)

    def insert_hint(self, text_widget, hint_text):
        """Вставка подсказки в текстовое поле"""
        text_widget.delete("1.0", tk.END)
        text_widget.insert("1.0", hint_text)

    def clear_field(self, string_var, widget):
        """Очистка поля"""
        string_var.set('')
        if hasattr(widget, 'set'):
            widget.set('')

    def delete_image(self, key):
        """Удаление загруженного изображения"""
        self.images[key] = None

        # Обновляем список изображений
        image_titles = {
            'layout': 'Планировка зоны хранения',
            'loggers': 'Расстановка логгеров',
            'temp_map': 'Температурная карта',
            'humidity_map': 'Влажностная карта'
        }

        title = image_titles.get(key, key)
        # Находим индекс в списке
        for i in range(self.images_listbox.size()):
            if title in self.images_listbox.get(i):
                self.images_listbox.delete(i)
                self.images_listbox.insert(i, f"{title} - Не загружено")
                break

        # Кнопка удаления остается видимой для визуальной целостности

    def pack_forget(self):
        """Скрытие фрейма"""
        self.parent.pack_forget()
