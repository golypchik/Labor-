#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Менеджер сессии - управление данными приложения
"""

import os
import shutil
import sqlite3
from pathlib import Path


class SessionManager:
    """Класс для управления сессией приложения"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.inform_dir = self.project_root / "inform"
        self.db_dir = self.project_root / "database"
        
        # Создаем необходимые директории
        self.inform_dir.mkdir(exist_ok=True)
        self.db_dir.mkdir(exist_ok=True)
        
        # Инициализация баз данных
        self.init_databases()
    
    def init_databases(self):
        """Инициализация баз данных"""
        # База данных периодов
        self.periods_db = self.db_dir / "periods.db"
        conn = sqlite3.connect(str(self.periods_db))
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                name TEXT NOT NULL,
                loggers TEXT,
                required_mode_from REAL,
                required_mode_to REAL
            )
        """)
        conn.commit()
        conn.close()
        
        # База данных статистики логгеров
        self.logger_stats_db = self.db_dir / "logger_stats.db"
        conn = sqlite3.connect(str(self.logger_stats_db))
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS logger_stats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period_id INTEGER NOT NULL,
                logger_number TEXT NOT NULL,
                data_type TEXT NOT NULL,
                min_value REAL,
                max_value REAL,
                avg_value REAL,
                logger_type TEXT,
                FOREIGN KEY (period_id) REFERENCES periods(id)
            )
        """)
        conn.commit()
        conn.close()
        
        # База данных настроек проекта
        self.settings_db = self.db_dir / "settings.db"
        conn = sqlite3.connect(str(self.settings_db))
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.commit()
        conn.close()
        
        # База данных прочей информации
        self.other_info_db = self.db_dir / "other_info.db"
        conn = sqlite3.connect(str(self.other_info_db))
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS other_info (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.commit()
        conn.close()
    
    def cleanup(self):
        """Очистка всех данных сессии"""
        # Удаление файлов из папки inform
        if self.inform_dir.exists():
            for file in self.inform_dir.iterdir():
                if file.is_file():
                    file.unlink()
        
        # Удаление баз данных
        for db_file in [self.periods_db, self.logger_stats_db, self.settings_db]:
            if db_file.exists():
                db_file.unlink()
        
        # Пересоздаем БД
        self.init_databases()
    
    def get_periods_db_path(self):
        """Получить путь к базе данных периодов"""
        return str(self.periods_db)
    
    def get_logger_stats_db_path(self):
        """Получить путь к базе данных статистики логгеров"""
        return str(self.logger_stats_db)
    
    def get_settings_db_path(self):
        """Получить путь к базе данных настроек"""
        return str(self.settings_db)
    
    def get_other_info_db_path(self):
        """Получить путь к базе данных прочей информации"""
        return str(self.other_info_db)
    
    def save_other_info(self, data):
        """Сохранить данные прочей информации"""
        conn = sqlite3.connect(str(self.other_info_db))
        cursor = conn.cursor()
        
        # Очищаем старые данные
        cursor.execute("DELETE FROM other_info")
        
        # Сохраняем новые данные
        for key, value in data.items():
            if value is not None:
                cursor.execute("INSERT OR REPLACE INTO other_info (key, value) VALUES (?, ?)", (key, str(value)))
        
        conn.commit()
        conn.close()
    
    def get_other_info(self):
        """Получить данные прочей информации"""
        conn = sqlite3.connect(str(self.other_info_db))
        cursor = conn.cursor()
        
        cursor.execute("SELECT key, value FROM other_info")
        rows = cursor.fetchall()
        conn.close()
        
        data = {}
        for key, value in rows:
            data[key] = value
        
        return data
