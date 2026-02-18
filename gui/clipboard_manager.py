#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Менеджер буфера обмена для приложения
Предоставляет универсальные функции копирования, вставки и контекстных меню
"""

import tkinter as tk
from tkinter import Menu
import tkinter.ttk as ttk
import pyperclip


class ClipboardManager:
    """Менеджер буфера обмена"""
    
    def __init__(self, root):
        self.root = root
        self.context_menus = {}  # Хранит контекстные меню для виджетов
        
        # Устанавливаем глобальные привязки
        self.setup_global_bindings()
    
    def setup_global_bindings(self):
        """Настройка глобальных привязок клавиш для всех раскладок"""
        # Привязываем события клавиш к главному окну
        self.root.bind("<KeyPress>", self.on_key_press)
        
        # Для Linux с Button-4 и Button-5 (тачпад)
        self.root.bind("<Button-4>", self.on_button4)
        self.root.bind("<Button-5>", self.on_button5)
    
    def on_key_press(self, event):
        """Обработка нажатий клавиш для copy-paste независимо от раскладки"""
        # Проверяем, активен ли виджет ввода
        widget = self.root.focus_get()
        if not self.is_input_widget(widget):
            return
        
        # Проверяем нажатие Ctrl (state & 4)
        ctrl_pressed = event.state & 0x04
        
        if not ctrl_pressed:
            return
        
        # Используем виртуальные коды клавиш для надежной работы с любой раскладкой
        # VK_C = 0x43, VK_V = 0x56, VK_X = 0x58, VK_A = 0x41
        key_code = event.keycode
        
        # Обрабатываем Ctrl+C (копирование) - VK_C = 67
        if key_code == 67:  # VK_C
            self.copy_to_clipboard(widget)
            return "break"  # Останавливаем распространение события
        
        # Обрабатываем Ctrl+V (вставка) - VK_V = 86
        elif key_code == 86:  # VK_V
            self.paste_from_clipboard(widget)
            return "break"  # Останавливаем распространение события
        
        # Обрабатываем Ctrl+X (вырезание) - VK_X = 88
        elif key_code == 88:  # VK_X
            self.cut_from_clipboard(widget)
            return "break"  # Останавливаем распространение события
        
        # Обрабатываем Ctrl+A (выделение всего) - VK_A = 65
        elif key_code == 65:  # VK_A
            self.select_all_text(widget)
            return "break"  # Останавливаем распространение события
    
    def on_button4(self, event):
        """Обработка Button-4 (тачпад вверх)"""
        # Если это прокрутка тачпада, не обрабатываем как контекстное меню
        if hasattr(event, 'delta') and event.delta < 0:
            return
        # Иначе - возможно, это долгое нажатие тачпада
        self.show_context_menu(event)
    
    def on_button5(self, event):
        """Обработка Button-5 (тачпад вниз)"""
        # Если это прокрутка тачпада, не обрабатываем как контекстное меню
        if hasattr(event, 'delta') and event.delta > 0:
            return
        # Иначе - возможно, это долгое нажатие тачпада
        self.show_context_menu(event)
    
    def is_input_widget(self, widget):
        """Проверка, является ли виджет полем ввода"""
        if widget is None:
            return False
        
        # Список типов виджетов, для которых доступен copy-paste
        input_widget_types = (
            tk.Entry, tk.Text, tk.Listbox, tk.Spinbox,
            ttk.Entry, ttk.Combobox, ttk.Spinbox
        )
        
        return isinstance(widget, input_widget_types)
    
    def copy_to_clipboard(self, widget):
        """Копирование текста в буфер обмена"""
        try:
            if isinstance(widget, tk.Entry) or isinstance(widget, ttk.Entry):
                # Для Entry копируем выделенный текст
                try:
                    selected_text = widget.selection_get()
                    pyperclip.copy(selected_text)
                except tk.TclError:
                    # Нет выделения - копируем весь текст
                    pyperclip.copy(widget.get())
            
            elif isinstance(widget, tk.Text):
                # Для Text копируем выделенный текст
                try:
                    selected_text = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
                    pyperclip.copy(selected_text)
                except tk.TclError:
                    # Нет выделения - копируем весь текст
                    pyperclip.copy(widget.get("1.0", tk.END).rstrip())
            
            elif isinstance(widget, tk.Listbox):
                # Для Listbox копируем выделенные элементы
                try:
                    selected_indices = widget.curselection()
                    if selected_indices:
                        selected_items = [widget.get(i) for i in selected_indices]
                        pyperclip.copy("\n".join(selected_items))
                except tk.TclError:
                    pass
            
            elif isinstance(widget, ttk.Combobox):
                # Для Combobox копируем текущее значение
                pyperclip.copy(widget.get())
        
        except Exception as e:
            print(f"Ошибка при копировании: {e}")
    
    def paste_from_clipboard(self, widget):
        """Вставка текста из буфера обмена"""
        try:
            clipboard_text = pyperclip.paste()
            if not clipboard_text:
                return
            
            if isinstance(widget, tk.Entry) or isinstance(widget, ttk.Entry):
                # Для Entry вставляем в позицию курсора
                try:
                    # Удаляем выделенный текст, если есть
                    widget.selection_clear()
                except tk.TclError:
                    pass
                
                # Вставляем текст
                cursor_pos = widget.index(tk.INSERT)
                widget.insert(cursor_pos, clipboard_text)
            
            elif isinstance(widget, tk.Text):
                # Для Text вставляем в позицию курсора
                try:
                    # Удаляем выделенный текст, если есть
                    widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError:
                    pass
                
                # Вставляем текст
                cursor_pos = widget.index(tk.INSERT)
                widget.insert(cursor_pos, clipboard_text)
            
            elif isinstance(widget, ttk.Combobox):
                # Для Combobox вставляем в поле ввода
                current_text = widget.get()
                cursor_pos = widget.index(tk.INSERT)
                new_text = current_text[:cursor_pos] + clipboard_text + current_text[cursor_pos:]
                widget.set(new_text)
        
        except Exception as e:
            print(f"Ошибка при вставке: {e}")
    
    def cut_from_clipboard(self, widget):
        """Вырезание текста в буфер обмена"""
        try:
            if isinstance(widget, tk.Entry) or isinstance(widget, ttk.Entry):
                # Для Entry вырезаем выделенный текст
                try:
                    selected_text = widget.selection_get()
                    pyperclip.copy(selected_text)
                    widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError:
                    # Нет выделения - вырезаем весь текст
                    all_text = widget.get()
                    pyperclip.copy(all_text)
                    widget.delete(0, tk.END)
            
            elif isinstance(widget, tk.Text):
                # Для Text вырезаем выделенный текст
                try:
                    selected_text = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
                    pyperclip.copy(selected_text)
                    widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError:
                    # Нет выделения - вырезаем весь текст
                    all_text = widget.get("1.0", tk.END).rstrip()
                    pyperclip.copy(all_text)
                    widget.delete("1.0", tk.END)
        
        except Exception as e:
            print(f"Ошибка при вырезании: {e}")
    
    def select_all_text(self, widget):
        """Выделение всего текста"""
        try:
            if isinstance(widget, tk.Entry) or isinstance(widget, ttk.Entry):
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
            
            elif isinstance(widget, tk.Text):
                widget.tag_add(tk.SEL, "1.0", tk.END)
                widget.mark_set(tk.INSERT, "1.0")
                widget.see(tk.INSERT)
            
            elif isinstance(widget, ttk.Combobox):
                widget.selection_range(0, tk.END)
        
        except Exception as e:
            print(f"Ошибка при выделении: {e}")
    
    def create_context_menu(self, widget):
        """Создание контекстного меню для виджета"""
        if not self.is_input_widget(widget):
            return
        
        # Создаем меню
        menu = Menu(widget, tearoff=0)
        
        # Добавляем пункты меню
        menu.add_command(
            label="Копировать",
            command=lambda: self.copy_to_clipboard(widget),
            accelerator="Ctrl+C"
        )
        
        menu.add_command(
            label="Вставить",
            command=lambda: self.paste_from_clipboard(widget),
            accelerator="Ctrl+V"
        )
        
        menu.add_command(
            label="Вырезать",
            command=lambda: self.cut_from_clipboard(widget),
            accelerator="Ctrl+X"
        )
        
        menu.add_separator()
        
        menu.add_command(
            label="Выделить все",
            command=lambda: self.select_all_text(widget),
            accelerator="Ctrl+A"
        )
        
        # Сохраняем меню
        self.context_menus[id(widget)] = menu
        
        # Привязываем события
        widget.bind("<Button-3>", lambda event: self.show_context_menu_at(event, widget))
        
        # Для Linux с Button-2 (средняя кнопка мыши) - вставка
        widget.bind("<Button-2>", lambda event: self.paste_from_clipboard(widget))
        
        # Для долгого нажатия тачпада (Button-4/Button-5 с задержкой)
        widget.bind("<ButtonPress-1>", lambda event: self.on_mouse_press(event, widget))
        widget.bind("<ButtonRelease-1>", lambda event: self.on_mouse_release(event, widget))
    
    def show_context_menu_at(self, event, widget):
        """Показ контекстного меню в указанной позиции"""
        menu = self.context_menus.get(id(widget))
        if menu:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
    
    def show_context_menu(self, event):
        """Показ контекстного меню для активного виджета"""
        widget = self.root.focus_get()
        if widget and self.is_input_widget(widget):
            menu = self.context_menus.get(id(widget))
            if menu:
                try:
                    menu.tk_popup(event.x_root, event.y_root)
                finally:
                    menu.grab_release()
    
    def on_mouse_press(self, event, widget):
        """Обработка нажатия мыши для долгого нажатия тачпада"""
        # Запускаем таймер для долгого нажатия
        widget._long_press_timer = self.root.after(500, lambda: self.show_context_menu_at(event, widget))
    
    def on_mouse_release(self, event, widget):
        """Обработка отпускания мыши"""
        # Отменяем таймер долгого нажатия
        if hasattr(widget, '_long_press_timer'):
            self.root.after_cancel(widget._long_press_timer)
            delattr(widget, '_long_press_timer')


def setup_clipboard_manager(root):
    """Функция для быстрой настройки менеджера буфера обмена"""
    return ClipboardManager(root)