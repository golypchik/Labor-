#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Главное окно приложения
"""

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import os
from pathlib import Path

from gui.project_management_frame import ProjectManagementFrame
from gui.key_elements_frame import KeyElementsFrame
from gui.other_info_frame import OtherInfoFrame
from gui.tables_creation_frame import TablesCreationFrame


class MainWindow:
    """Главное окно приложения"""
    
    def __init__(self, root, session_manager):
        self.root = root
        self.session_manager = session_manager
        self.current_section = "project_management"
        
        # Получаем путь к логотипу
        self.project_root = Path(__file__).parent.parent
        self.logo_path = self.project_root / "image" / "logo.png"
        
        self.create_widgets()
    
    def create_widgets(self):
        """Создание виджетов главного окна"""
        # Левая панель навигации
        self.create_sidebar()
        
        # Правая область содержимого
        self.create_content_area()
        
        # Загрузка и отображение логотипа
        self.load_logo()
    
    def create_sidebar(self):
        """Создание боковой панели"""
        # Фрейм для боковой панели
        sidebar_frame = tk.Frame(self.root, bg="#2c3e50", width=250)
        sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)
        sidebar_frame.pack_propagate(False)
        
        # Область для логотипа (будет добавлена позже)
        self.logo_frame = tk.Frame(sidebar_frame, bg="#2c3e50", height=150)
        self.logo_frame.pack(fill=tk.X, pady=10)
        
        # Кнопки навигации
        nav_buttons = [
            ("Управление проектом", "project_management"),
            ("Ключевые элементы", "key_elements"),
            ("Прочая информация", "other_info"),
            ("Создание таблиц", "tables_creation")
        ]
        
        for text, section in nav_buttons:
            btn = tk.Button(
                sidebar_frame,
                text=text,
                command=lambda s=section: self.switch_section(s),
                bg="#34495e",
                fg="white",
                font=("Arial", 11),
                relief=tk.FLAT,
                padx=20,
                pady=10,
                anchor=tk.W,
                cursor="hand2"
            )
            btn.pack(fill=tk.X, padx=10, pady=5)
        
        # Кнопка "Завершить сессию"
        end_session_btn = tk.Button(
            sidebar_frame,
            text="Завершить сессию",
            command=self.end_session,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 11, "bold"),
            relief=tk.FLAT,
            padx=20,
            pady=15,
            cursor="hand2"
        )
        end_session_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
    
    def create_content_area(self):
        """Создание области содержимого"""
        # Фрейм для содержимого
        self.content_frame = tk.Frame(self.root, bg="#ecf0f1")
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Создание фреймов для разных разделов
        self.frames = {}
        self.frame_containers = {}
        
        # Фрейм управления проектом
        container_pm = tk.Frame(self.content_frame, bg="#ecf0f1")
        self.frames["project_management"] = ProjectManagementFrame(
            container_pm, self.session_manager
        )
        self.frame_containers["project_management"] = container_pm
        
        # Фрейм ключевых элементов
        container_ke = tk.Frame(self.content_frame, bg="#ecf0f1")
        self.frames["key_elements"] = KeyElementsFrame(
            container_ke, self.session_manager
        )
        self.frame_containers["key_elements"] = container_ke
        
        # Фрейм прочей информации
        container_oi = tk.Frame(self.content_frame, bg="#ecf0f1")
        self.frames["other_info"] = OtherInfoFrame(
            container_oi, self.session_manager
        )
        self.frame_containers["other_info"] = container_oi
        
        # Сохраняем ссылку на главное окно в фрейме
        self.frames["other_info"].main_window = self
        
        # Фрейм создания таблиц
        container_tc = tk.Frame(self.content_frame, bg="#ecf0f1")
        self.frames["tables_creation"] = TablesCreationFrame(
            container_tc, self.session_manager, main_window=self
        )
        self.frame_containers["tables_creation"] = container_tc
        
        # Показываем начальный фрейм
        self.switch_section("project_management")
    
    def load_logo(self):
        """Загрузка и отображение логотипа"""
        if self.logo_path.exists():
            try:
                # Загружаем изображение
                img = Image.open(self.logo_path)
                # Масштабируем до нужного размера (уменьшаем ширину на 10%)
                width = int(200 * 0.72)  # 72% от исходной ширины
                height = 120
                img = img.resize((width, height), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                # Создаем лейбл для логотипа
                logo_label = tk.Label(
                    self.logo_frame,
                    image=photo,
                    bg="#2c3e50"
                )
                logo_label.image = photo  # Сохраняем ссылку
                logo_label.pack(pady=10)
            except Exception as e:
                print(f"Ошибка загрузки логотипа: {e}")
        else:
            # Если логотипа нет, создаем заглушку
            placeholder = tk.Label(
                self.logo_frame,
                text="Логотип\n(logo.png)",
                bg="#2c3e50",
                fg="white",
                font=("Arial", 10)
            )
            placeholder.pack(pady=20)
    
    def switch_section(self, section):
        """Переключение между разделами"""
        # Скрываем все контейнеры
        for container in self.frame_containers.values():
            container.pack_forget()
        
        # Показываем выбранный контейнер
        self.frame_containers[section].pack(fill=tk.BOTH, expand=True)
        self.current_section = section
    
    def show_custom_askyesno(self, title, message):
        """Кастомный диалог для вопроса с выбором Yes/No"""
        # Создаем топ-level окно
        top = tk.Toplevel(self.root)
        top.title(title)
        top.geometry("400x180")
        top.resizable(False, False)
        top.attributes("-topmost", True)
        
        # Центрируем окно
        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (400 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (180 // 2)
        top.geometry(f"+{x}+{y}")
        
        # Добавляем сообщение
        label = tk.Label(top, text=message, font=("Arial", 12), padx=20, pady=20)
        label.pack()
        
        # Переменная для хранения ответа
        self.answer = None
        
        # Кнопки Yes и No
        buttons_frame = tk.Frame(top, bg="#ecf0f1")
        buttons_frame.pack(pady=10)
        
        def yes_callback():
            self.answer = True
            top.destroy()
        
        def no_callback():
            self.answer = False
            top.destroy()
        
        yes_btn = tk.Button(buttons_frame, text="Да", command=yes_callback, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), padx=20, pady=5)
        yes_btn.pack(side=tk.LEFT, padx=10)
        
        no_btn = tk.Button(buttons_frame, text="Нет", command=no_callback, bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), padx=20, pady=5)
        no_btn.pack(side=tk.LEFT, padx=10)
        
        # Делаем окно модальным
        top.grab_set()
        top.wait_window()
        
        return self.answer

    def show_custom_message(self, title, message):
        """Кастомное модальное окно без системного звука"""
        # Создаем топ-level окно
        top = tk.Toplevel(self.root)
        top.title(title)
        top.geometry("300x150")
        top.resizable(False, False)
        top.attributes("-topmost", True)
        
        # Центрируем окно
        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (300 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (150 // 2)
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

    def end_session(self):
        """Завершение сессии"""
        if self.show_custom_askyesno(
            "Подтверждение",
            "Вы уверены, что хотите завершить сессию?\nВсе данные будут удалены."
        ):
            # Очищаем временную папку с фото
            temp_dir = Path("temp_photos")
            if temp_dir.exists():
                import shutil
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    print(f"Ошибка очистки временной папки с фото: {e}")
            # Очищаем данные сессии
            self.session_manager.cleanup()
            # Очищаем данные во всех фреймах
            for frame in self.frames.values():
                if hasattr(frame, 'clear_data'):
                    frame.clear_data()
            # Переключаемся на управление проектом
            self.switch_section("project_management")
            self.show_custom_message("Информация", "Сессия завершена. Данные очищены.")
