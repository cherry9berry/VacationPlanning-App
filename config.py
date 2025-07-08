#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль конфигурации приложения
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any


class Config:
    """Класс для управления конфигурацией приложения"""
    
    DEFAULT_CONFIG = {
        "employee_template": "templates/employee_template.xlsx",
        "block_report_template": "templates/block_report_template.xlsx", 
        "general_report_template": "templates/global_report_template.xlsx",
        "header_row": 5,
        "processing_time_per_file": 0.3,
        "excel_password": "1111",
        "date_format": "%d.%m.%y",
        "max_employees": 10000,
        "min_employees": 1,
        "window_width": 1000,
        "window_height": 700,
        "log_level": "INFO"
    }
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = Path(config_file)
        self.data = self.DEFAULT_CONFIG.copy()
        self.logger = logging.getLogger(__name__)
    
    def load_or_create_default(self) -> None:
        """Загружает конфигурацию из файла или создает файл по умолчанию"""
        try:
            if self.config_file.exists():
                self.load()
                self.logger.info(f"Конфигурация загружена из {self.config_file}")
            else:
                self.save()
                self.logger.info(f"Создан файл конфигурации по умолчанию: {self.config_file}")
        except Exception as e:
            self.logger.error(f"Ошибка работы с конфигурацией: {e}")
            self.data = self.DEFAULT_CONFIG.copy()
    
    def load(self) -> None:
        """Загружает конфигурацию из файла"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)
                
            # Обновляем только существующие ключи, остальные берем из DEFAULT_CONFIG
            for key, value in loaded_data.items():
                if key in self.DEFAULT_CONFIG:
                    self.data[key] = value
                    
        except FileNotFoundError:
            self.logger.warning(f"Файл конфигурации не найден: {self.config_file}")
            raise
        except json.JSONDecodeError as e:
            self.logger.error(f"Ошибка парсинга JSON: {e}")
            raise
        except Exception as e:
            self.logger.error(f"Ошибка загрузки конфигурации: {e}")
            raise
    
    def save(self) -> None:
        """Сохраняет конфигурацию в файл"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения конфигурации: {e}")
            raise
    
    def get(self, key: str, default=None):
        """Получает значение из конфигурации"""
        return self.data.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Устанавливает значение в конфигурации"""
        self.data[key] = value
    
    def validate_templates(self) -> Dict[str, bool]:
        """Проверяет наличие шаблонов"""
        templates = {
            "employee_template": self.get("employee_template"),
            "block_report_template": self.get("block_report_template"),
            "general_report_template": self.get("general_report_template")
        }
        
        results = {}
        for name, path in templates.items():
            file_path = Path(path)
            results[name] = file_path.exists()
            
        return results
    
    @property
    def employee_template(self) -> str:
        return self.get("employee_template")
    
    @property 
    def block_report_template(self) -> str:
        return self.get("block_report_template")
    
    @property
    def general_report_template(self) -> str:
        return self.get("general_report_template")
    
    @property
    def header_row(self) -> int:
        return self.get("header_row", 5)
    
    @property
    def excel_password(self) -> str:
        return self.get("excel_password", "1111")
    
    @property
    def date_format(self) -> str:
        return self.get("date_format", "%d.%m.%y")
    
    @property
    def max_employees(self) -> int:
        return self.get("max_employees", 10000)
    
    @property
    def min_employees(self) -> int:
        return self.get("min_employees", 1)