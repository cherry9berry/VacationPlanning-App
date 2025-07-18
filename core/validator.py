#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль валидации данных
"""

import logging
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any, Union
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
    
    def validate_staff_file(self, file_path: str) -> Tuple[ValidationResult, List[Dict[str, str]]]:
        """
        Валидирует файл штатного расписания
        
        Returns:
            Tuple[ValidationResult, List[Dict[str, str]]]: результат валидации и список сотрудников
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
            
            # ИСПРАВЛЕНО: Проверка обязательных заголовков БЕЗ чтения rules
            required_fields = ["ФИО работника", "Табельный номер", "Подразделение 1"]
            header_map = {header.strip(): idx for idx, header in enumerate(headers) if header}
            
            for required in required_fields:
                if required not in header_map:
                    result.add_error(f"Отсутствует обязательный заголовок: {required}")
            
            if not result.is_valid:
                return result, employees
            
            # Загружаем правила заполнения из шаблона для динамического маппинга
            from core.excel_handler import ExcelHandler
            excel_handler = ExcelHandler(self.config)
            rules = excel_handler._load_filling_rules(self.config.employee_template)
            needed_fields = list(rules.get('value', {}).values())
            # Чтение данных сотрудников по нужным полям
            all_employees = self._read_employees(worksheet, header_row, header_map, needed_fields)
            
            # ИСПРАВЛЕННАЯ ЛОГИКА: Валидация и фильтрация сотрудников
            employees, validation_stats = self._validate_and_filter_employees(all_employees, result)
            
            result.employee_count = len(employees)
            result.unique_tab_numbers = validation_stats['unique_tab_numbers']
            
            # ИСПРАВЛЕНО: Добавляем явную проверку типов и значение по умолчанию
            processing_time_per_file = self.config.get("processing_time_per_file", 0.3)
            if processing_time_per_file is None:
                processing_time_per_file = 0.3
            result.processing_time = len(employees) * float(processing_time_per_file)
            
            self.logger.info(f"Валидация завершена. Найдено сотрудников: {len(employees)}")
            
        except Exception as e:
            result.add_error(f"Неожиданная ошибка при валидации: {e}")
            self.logger.error(f"Ошибка валидации файла: {e}", exc_info=True)
        
        return result, employees
    
    def _validate_and_filter_employees(self, all_employees: List[Dict[str, str]], result: ValidationResult) -> Tuple[List[Dict[str, str]], Dict[str, int]]:
        """
        ИСПРАВЛЕННАЯ ЛОГИКА: Валидирует сотрудников и исключает дублирующиеся табельные номера
        
        Returns:
            Tuple[List[Dict[str, str]], Dict[str, int]]: отфильтрованный список сотрудников и статистика
        """
        tab_numbers: Dict[str, List[Dict[str, str]]] = {}
        duplicate_tab_numbers: Dict[str, List[str]] = {}
        valid_employees: List[Dict[str, str]] = []
        
        # Первый проход - выявляем дублирующиеся табельные номера
        for i, emp in enumerate(all_employees, 1):
            # Проверка длины строк
            if len(emp.get('ФИО работника', '')) > 255:
                result.add_warning(f"Строка {i}: ФИО слишком длинное (>255 символов)")
            
            if len(emp.get('Подразделение 1', '')) > 255:
                result.add_warning(f"Строка {i}: Название подразделения слишком длинное")
            
            # Проверка табельного номера
            tab_number = emp.get('Табельный номер')
            if not tab_number:
                result.add_error(f"Строка {i}: Пустой табельный номер")
                continue
            
            # ИСПРАВЛЕНО: Проверка формата табельного номера с проверкой на None
            tab_number_str = str(tab_number).strip()
            if tab_number_str and not re.match(r'^\d+$', tab_number_str):
                result.add_warning(f"Строка {i}: Табельный номер не является числом: {tab_number_str}")
            
            # Подсчет табельных номеров
            if tab_number_str in tab_numbers:
                tab_numbers[tab_number_str].append(emp)
            else:
                tab_numbers[tab_number_str] = [emp]
        
        # Второй проход - определяем какие табельные номера дублируются
        for tab_num, emp_list in tab_numbers.items():
            if len(emp_list) > 1:
                # Дублирующийся табельный номер - добавляем в статистику и исключаем всех
                names = [emp.get('ФИО работника', '') for emp in emp_list]
                duplicate_tab_numbers[tab_num] = names
                result.add_warning(f"Дублирующийся табельный номер {tab_num}: {', '.join(names)}")
            else:
                # Уникальный табельный номер - добавляем сотрудника в валидный список
                valid_employees.extend(emp_list)
        
        # Проверка количества сотрудников
        unique_count = len(valid_employees)  # Теперь только уникальные
        min_employees = self.config.min_employees
        max_employees = self.config.max_employees
        
        if unique_count < min_employees:
            result.add_error(f"Слишком мало уникальных сотрудников: {unique_count} (минимум {min_employees})")
        
        if unique_count > max_employees:
            result.add_error(f"Слишком много сотрудников: {unique_count} (максимум {max_employees})")
        
        validation_stats = {
            'unique_tab_numbers': len(tab_numbers) - len(duplicate_tab_numbers),
            'duplicate_count': len(duplicate_tab_numbers),
            'excluded_employees': sum(len(names) for names in duplicate_tab_numbers.values())
        }
        
        return valid_employees, validation_stats
    
    def _read_employees(self, worksheet, header_row: int, header_map: Dict[str, int], rules_value_fields: Optional[List[str]] = None) -> List[Dict[str, str]]:
        """Динамически читает сотрудников по правилам rules['value'] и заголовкам исходника"""
        employees = []
        if rules_value_fields is None:
            # Если не передали список нужных полей, работаем по старой логике
            rules_value_fields = list(header_map.keys())
        
        row_num = header_row + 1
        while True:
            row_values = self._get_row_values(worksheet, row_num)
            if not row_values or all(not val for val in row_values):
                break
            
            employee_data = {}
            for field in rules_value_fields:
                col_idx = header_map.get(field)
                value = row_values[col_idx] if col_idx is not None and col_idx < len(row_values) else ''
                if value is None:
                    value = ''
                employee_data[field] = str(value).strip()
            
            # ИСПРАВЛЕНО: В правильном входном файле нет дат отпусков в строках 15-29
            # Даты отпусков будут заполняться вручную в формах сотрудников
            employee_data['vacation_dates'] = []
            
            # Добавляем только если есть обязательные поля
            if (employee_data.get('ФИО работника') and 
                employee_data.get('Табельный номер') and 
                employee_data.get('Подразделение 1')):
                employees.append(employee_data)
            
            row_num += 1
        
        return employees
    
    def _get_row_values(self, worksheet, row_num: int) -> List[Any]:
        """Получает значения строки из листа Excel"""
        try:
            row = list(worksheet.iter_rows(min_row=row_num, max_row=row_num, values_only=True))[0]
            return [cell if cell is not None else "" for cell in row]
        except (IndexError, AttributeError):
            return []
    
    def _count_unique_tab_numbers(self, employees: List[Dict[str, str]]) -> int:
        """Подсчитывает количество уникальных табельных номеров"""
        unique_tabs = set()
        for emp in employees:
            if emp.get('Табельный номер'):
                unique_tabs.add(emp['Табельный номер'])
        return len(unique_tabs)
    
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