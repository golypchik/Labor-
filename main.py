#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Лабор.Картирование - Приложение для обработки данных картирования
"""

import tkinter as tk
from tkinter import messagebox
import os
import sys

# Добавляем корневую директорию в путь
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import MainWindow
from utils.session_manager import SessionManager


def on_closing(root, session_manager):
    """Обработчик закрытия окна"""
    if messagebox.askyesno(
        "Подтверждение",
        "Вы уверены, что хотите закрыть приложение?\nВсе данные будут удалены."
    ):
        session_manager.cleanup()
        root.destroy()


def main():
    """Главная функция приложения"""
    root = tk.Tk()
    root.title("Лабор.Картирование")
    root.geometry("1200x700")
    
    # Инициализация менеджера сессии
    session_manager = SessionManager()
    
    # Создание главного окна
    app = MainWindow(root, session_manager)
    
    # Установка обработчика закрытия окна
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root, session_manager))
    
    # Запуск главного цикла
    root.mainloop()


if __name__ == "__main__":
    main()
