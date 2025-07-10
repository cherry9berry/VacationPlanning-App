#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль валидации данных
"""

import logging
from pathlib import Path
from typing import List, Tuple, Optional
import re

import openpyxl
from openpyxl import Workbook

from models import Employee, ValidationResult
from config import Config


class Validator:
    """Класс для валидации файлов и данных"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def validate_staff_file(self, file_path: str) -> Tuple[ValidationResult, List[Employee]]:
        """
        Валидирует файл штатного расписания
        
        Returns:
            Tuple[ValidationResult, List[Employee]]: результат валидации и список сотрудников
        """
        result = ValidationResult()
        employees = []
        
        try:
            # Проверка существования файла
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                result.add_error(f"Файл не существует: {file_path}")
                return result, employees
            
            # Проверка размера файла (максимум 50 МБ)
            file_size = file_path_obj.stat().st_size
            if file_size > 50 * 1024 * 1024:
                result.add_error("Размер файла превышает 50 МБ")
                return result, employees
            
            # Открытие Excel файла
            try:
                workbook = openpyxl.load_workbook(file_path, data_only=True)
            except Exception as e:
                result.add_error(f"Ошибка открытия Excel файла: {e}")
                return result, employees
            
            # Получение первого листа
            if not workbook.worksheets:
                result.add_error("В файле нет листов")
                return result, employees
            
            worksheet = workbook.active
            
            # Проверка заголовков
            header_row = self.config.header_row
            headers = self._get_row_values(worksheet, header_row)
            
            if not headers:
                result.add_error(f"Не удалось прочитать строку заголовков {header_row}")
                return result, employees
            
            # Проверка обязательных заголовков
            required_headers = ["ФИО работника", "Табельный номер", "Подразделение 1"]
            header_map = {header.strip(): idx for idx, header in enumerate(headers) if header}
            
            for required in required_headers:
                if required not in header_map:
                    result.add_error(f"Отсутствует обязательный заголовок: {required}")
            
            if not result.is_valid:
                return result, employees
            
            # Чтение данных сотрудников
            employees = self._read_employees(worksheet, header_row, header_map)
            
            # Валидация данных сотрудников
            self._validate_employees(employees, result)
            
            result.employee_count = len(employees)
            result.unique_tab_numbers = self._count_unique_tab_numbers(employees)
            result.processing_time = len(employees) * self.config.get("processing_time_per_file", 0.3)
            
            self.logger.info(f"Валидация завершена. Найдено сотрудников: {len(employees)}")
            
        except Exception as e:
            result.add_error(f"Неожиданная ошибка при валидации: {e}")
            self.logger.error(f"Ошибка валидации файла: {e}", exc_info=True)
        
        return result, employees
    
    def validate_templates(self) -> ValidationResult:
        """Проверяет наличие всех необходимых шаблонов"""
        result = ValidationResult()
        
        templates = {
            "Шаблон сотрудника": self.config.employee_template,
            "Шаблон отчета по блоку": self.config.block_report_template,
            "Шаблон общего отчета": self.config.general_report_template
        }
        
        for name, path in templates.items():
            if not Path(path).exists():
                result.add_error(f"Отсутствует {name}: {path}")
        
        return result
    
    def validate_output_directory(self, dir_path: str) -> ValidationResult:
        """Проверяет возможность записи в выходную папку"""
        result = ValidationResult()
        
        try:
            # Создаем папку если она не существует
            output_path = Path(dir_path)
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Проверяем права на запись
            test_file = output_path / "test_write_permissions.tmp"
            try:
                test_file.touch()
                test_file.unlink()
            except Exception as e:
                result.add_error(f"Нет прав на запись в папку: {e}")
                
        except Exception as e:
            result.add_error(f"Невозможно создать выходную папку: {e}")
        
        return result
    
    def _get_row_values(self, worksheet, row_num: int) -> List[str]:
        """Получает значения строки"""
        values = []
        for cell in worksheet[row_num]:
            value = cell.value
            if value is None:
                values.append("")
            else:
                values.append(str(value).strip())
        return values
    
    def _read_employees(self, worksheet, header_row: int, header_map: dict) -> List[Employee]:
        """Читает данных сотрудников из листа"""
        employees = []
        
        # Начинаем с строки после заголовков
        for row_num in range(header_row + 1, worksheet.max_row + 1):
            row_values = self._get_row_values(worksheet, row_num)
            
            # Пропускаем пустые строки
            if not any(row_values):
                continue
            
            employee = Employee()
            
            # Заполняем данные сотрудника
            for field_name, column_idx in header_map.items():
                if column_idx < len(row_values):
                    value = row_values[column_idx].strip()
                    
                    if field_name == "ФИО работника":
                        employee.full_name = value
                    elif field_name == "Табельный номер":
                        employee.tab_number = value
                    elif field_name == "Должность":
                        employee.position = value
                    elif field_name == "Подразделение 1":
                        employee.department1 = value
                    elif field_name == "Подразделение 2":
                        employee.department2 = value
                    elif field_name == "Подразделение 3":
                        employee.department3 = value
                    elif field_name == "Подразделение 4":
                        employee.department4 = value
                    # Новые поля
                    elif field_name == "Локация":
                        employee.location = value
                    elif field_name == "Остатки отпуска":
                        employee.vacation_remainder = value
                    elif field_name == "Дата приема":
                        employee.hire_date = value
                    elif field_name == "Дата отсечки периода":
                        employee.period_cutoff_date = value
                    elif field_name == "Дополнительный отпуск НРД":
                        employee.additional_vacation_nrd = value
                    elif field_name == "Дополнительный отпуск Северный":
                        employee.additional_vacation_north = value
            
            # Добавляем только если есть обязательные поля
            if employee.full_name and employee.tab_number and employee.department1:
                employees.append(employee)
        
        return employees
    
    def _validate_employees(self, employees: List[Employee], result: ValidationResult):
        """Валидирует данные сотрудников"""
        tab_numbers = {}
        
        for i, emp in enumerate(employees, 1):
            # Проверка длины строк
            if len(emp.full_name) > 255:
                result.add_warning(f"Строка {i}: ФИО слишком длинное (>255 символов)")
            
            if len(emp.department1) > 255:
                result.add_warning(f"Строка {i}: Название подразделения слишком длинное")
            
            # Проверка табельного номера
            if not emp.tab_number:
                result.add_error(f"Строка {i}: Пустой табельный номер")
                continue
            
            # Проверка формата табельного номера
            if not re.match(r'^\d+$', emp.tab_number):
                result.add_warning(f"Строка {i}: Табельный номер не является числом: {emp.tab_number}")
            
            # Проверка уникальности табельного номера
            if emp.tab_number in tab_numbers:
                tab_numbers[emp.tab_number] += 1
                result.add_warning(f"Дублирующийся табельный номер: {emp.tab_number}")
            else:
                tab_numbers[emp.tab_number] = 1
        
        # Проверка количества сотрудников
        unique_count = len(tab_numbers)
        min_employees = self.config.min_employees
        max_employees = self.config.max_employees
        
        if unique_count < min_employees:
            result.add_error(f"Слишком мало сотрудников: {unique_count} (минимум {min_employees})")
        
        if unique_count > max_employees:
            result.add_error(f"Слишком много сотрудников: {unique_count} (максимум {max_employees})")
    
    def _count_unique_tab_numbers(self, employees: List[Employee]) -> int:
        """Подсчитывает количество уникальных табельных номеров"""
        unique_tabs = set()
        for emp in employees:
            if emp.tab_number:
                unique_tabs.add(emp.tab_number)
        return len(unique_tabs)
    
