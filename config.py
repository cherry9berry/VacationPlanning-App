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
        "employee_template": "templates/employee_template v4.0.xlsx",
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
        
        # Параметры календарного года
        "target_year": 2026,
        "is_leap_year": False,
        "days_in_months": [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
        "month_names": [
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ],
        
        # Структура файла сотрудника для чтения данных
        "employee_file_structure": {
            # Основные параметры
            "active_sheet_index": 0,  # Первый лист (индекс 0)
            "status_cell": "B12",     # Ячейка со статусом формы
            
            # Данные сотрудника
            "employee_data_rows": {"start": 15, "end": 29},  # Строки с данными сотрудников (новый шаблон)
            "employee_columns": {
                "tab_number": "B",    # Столбец табельного номера
                "full_name": "C",     # Столбец ФИО
                "position": "D"       # Столбец должности
            },
            
            # Подразделения в шапке
            "department_cells": {
                "department1": "C2",
                "department2": "C3", 
                "department3": "C4",
                "department4": "C5"
            },
            
            # Периоды отпусков
            "vacation_columns": {
                "start_date": "E",    # Столбец даты начала отпуска
                "end_date": "F",      # Столбец даты окончания отпуска
                "days": "G"           # Столбец количества дней
            },
            
            # Ячейки валидации (для обратной совместимости, но основной статус из B12)
            "validation_cells": {
                "total_days": "J2",   # Общее количество дней
                "error_h2": "H2",     # Ячейка с ошибками H2
                "error_i2": "I2"      # Ячейка с ошибками I2
            }
        },
        
        # Структура отчетов
        "report_structure": {
            # Календарная матрица
            "calendar_start_col": 12,      # Столбец L (начало календаря)
            "calendar_month_row": 7,       # Строка с названиями месяцев
            "calendar_day_row": 8,         # Строка с числами дней
            
            # Данные сотрудников в отчете
            "employee_data_start_row": 9,  # Начальная строка данных сотрудников
            
            # Ячейки шапки отчета (Report лист)
            "report_header_cells": {
                "block_name": "A3",
                "update_date": "A4", 
                "total_employees": "A5",
                "completed": "A6"
            },
            
            # Столбцы таблицы сотрудников (Report лист)
            "report_employee_columns": {
                "row_number": "A",      # № п/п
                "full_name": "B",       # ФИО
                "tab_number": "C",      # Табельный номер
                "position": "D",        # Должность
                "department1": "E",     # Подразделение 1
                "department2": "F",     # Подразделение 2
                "department3": "G",     # Подразделение 3
                "department4": "H",     # Подразделение 4
                "status": "I",          # Статус планирования
                "total_days": "J",      # Итого дней
                "periods_count": "K"    # Количество периодов
            },
            
            # Настройки Print листа
            "print_structure": {
                "block_name_cell": "D4",
                "data_start_row": 9,
                "pagination": {
                    "first_page_records": 14,
                    "other_pages_records": 18
                }
            }
        },
        
        # Статусы валидации (новая логика)
        "validation_statuses": {
            "not_filled": "Форма не заполнена",
            "filled_incorrect": "Форма заполнена некорректно", 
            "filled_correct": "Форма заполнена корректно"
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
    
    # Календарные параметры
    @property
    def target_year(self) -> int:
        return self.get("target_year", 2026)
    
    @property
    def is_leap_year(self) -> bool:
        return self.get("is_leap_year", False)
    
    @property
    def days_in_months(self) -> list:
        return self.get("days_in_months", [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
    
    @property
    def month_names(self) -> list:
        return self.get("month_names", [
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ])
    
    # Структура файла сотрудника
    @property
    def employee_file_structure(self) -> dict:
        return self.get("employee_file_structure", {})
    
    @property
    def report_structure(self) -> dict:
        return self.get("report_structure", {})
    
    @property
    def validation_statuses(self) -> dict:
        return self.get("validation_statuses", {
            "not_filled": "Форма не заполнена",
            "filled_incorrect": "Форма заполнена некорректно", 
            "filled_correct": "Форма заполнена корректно"
        })