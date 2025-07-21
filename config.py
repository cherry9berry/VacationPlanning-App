#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль конфигурации приложения
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any


# ----------------------
# Класс конфигурации приложения
# ----------------------
class Config:
    """Класс для управления конфигурацией приложения"""
    
    # ----------------------
    # Словарь с настройками по умолчанию
    # ----------------------
    DEFAULT_CONFIG = {
        # Путь к шаблону файла сотрудника
        "employee_template": "templates/employee_template v4.2.xlsx",
        # Путь к шаблону отчета по блоку
        "block_report_template": "templates/block_report_template v3.xlsx", 
        # Путь к шаблону общего отчета
        "general_report_template": "templates/global_report_template v3.xlsx",
        # Номер строки с заголовками в Excel-файле сотрудников
        "header_row": 5,
        # Оценочное время обработки одного файла (секунды)
        "processing_time_per_file": 0.6,
        # Пароль для Excel-файлов (если используется)
        "excel_password": "1111",
        # Формат даты для отображения и парсинга
        "date_format": "%d.%m.%y",
        # Максимальное количество сотрудников в одном файле
        "max_employees": 10000,
        # Минимальное количество сотрудников в одном файле
        "min_employees": 1,
        # Ширина окна приложения по умолчанию
        "window_width": 1000,
        # Высота окна приложения по умолчанию
        "window_height": 700,
        # Параметры календарного года
        # Год, для которого строится календарь
        "target_year": 2026,
        # Признак високосного года
        "is_leap_year": False,
        # Количество дней в каждом месяце (январь-декабрь)
        "days_in_months": [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
        # Названия месяцев для отображения
        "month_names": [
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ],
        # Статусы валидации (для проверки заполнения форм)
        "validation_statuses": {
            "not_filled": "Форма не заполнена",
            "filled_incorrect": "Форма заполнена некорректно", 
            "filled_correct": "Форма заполнена корректно"
        },
        # Структура файла сотрудника
        "employee_file_structure": {
            "active_sheet_index": 0,
            "department_cells": {
                "department1": "C2",
                "department2": "C3", 
                "department3": "C4",
                "department4": "C5"
            },
            "employee_columns": {
                "full_name": "C",
                "tab_number": "B",
                "position": "D"
            },
            "employee_data_rows": {
                "start": 15,
                "end": 29
            },
            "vacation_columns": {
                "start_date": "C",
                "end_date": "D"
            },
            "status_cell": "B12"
        }
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
            if path is not None:
                file_path = Path(str(path))
                results[name] = file_path.exists()
            else:
                results[name] = False
            
        return results
    
    @property
    def employee_template(self) -> str:
        value = self.get("employee_template")
        return str(value) if value is not None else ""
    
    @property 
    def block_report_template(self) -> str:
        value = self.get("block_report_template")
        return str(value) if value is not None else ""
    
    @property
    def general_report_template(self) -> str:
        value = self.get("general_report_template")
        return str(value) if value is not None else ""
    
    @property
    def header_row(self) -> int:
        value = self.get("header_row")
        return int(value) if value is not None else 5
    
    @property
    def excel_password(self) -> str:
        value = self.get("excel_password")
        return str(value) if value is not None else "1111"
    
    @property
    def date_format(self) -> str:
        value = self.get("date_format")
        return str(value) if value is not None else "%d.%m.%y"
    
    @property
    def max_employees(self) -> int:
        value = self.get("max_employees")
        return int(value) if value is not None else 10000
    
    @property
    def min_employees(self) -> int:
        value = self.get("min_employees")
        return int(value) if value is not None else 1
    
    # Календарные параметры
    @property
    def target_year(self) -> int:
        value = self.get("target_year")
        return int(value) if value is not None else 2026
    
    @property
    def is_leap_year(self) -> bool:
        value = self.get("is_leap_year")
        return bool(value) if value is not None else False
    
    @property
    def month_names(self) -> list:
        value = self.get("month_names")
        if value is not None:
            return list(value)
        return ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    
    @property
    def days_in_months(self) -> list:
        value = self.get("days_in_months")
        if value is not None:
            return list(value)
        return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    
    @property
    def report_structure(self) -> dict:
        value = self.get("report_structure")
        return dict(value) if value is not None else {}
    
    @property
    def employee_file_structure(self) -> dict:
        value = self.get("employee_file_structure")
        return dict(value) if value is not None else {}
    
    @property
    def validation_statuses(self) -> dict:
        value = self.get("validation_statuses")
        if value is not None:
            return dict(value)
        return {
            "not_filled": "Форма не заполнена",
            "filled_incorrect": "Форма заполнена некорректно", 
            "filled_correct": "Форма заполнена корректно"
        }