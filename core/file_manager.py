#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль управления файлами и папками
"""

import logging
import time
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Callable, Optional, Tuple

from models import Employee, VacationInfo
from config import Config
from core.excel_handler import ExcelHandler


class FileManager:
    """Класс для управления файлами и структурой папок"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.excel_handler = ExcelHandler(config)
        
    def scan_existing_departments(self, target_directory: str) -> Dict[str, str]:
        """
        Сканирует существующие папки подразделений
        
        Args:
            target_directory: путь к целевой папке
            
        Returns:
            Dict[str, str]: словарь {название_подразделения: путь_к_папке}
        """
        try:
            target_path = Path(target_directory)
            if not target_path.exists():
                self.logger.warning(f"Целевая папка не существует: {target_directory}")
                return {}
            
            departments = {}
            
            # Сканируем все папки в целевой директории
            for item in target_path.iterdir():
                if item.is_dir():
                    # Исключаем системные папки
                    if not item.name.startswith('.') and not item.name.startswith('__'):
                        departments[item.name] = str(item)
            
            self.logger.info(f"Найдено подразделений: {len(departments)}")
            return departments
            
        except Exception as e:
            self.logger.error(f"Ошибка сканирования папки {target_directory}: {e}")
            return {}

    def group_employees_by_department(self, employees: List[Dict[str, str]]) -> Dict[str, List[Dict[str, str]]]:
        """
        Группирует сотрудников по подразделениям
        
        Args:
            employees: список сотрудников
            
        Returns:
            Dict[str, List[Dict[str, str]]]: словарь {подразделение: [сотрудники]}
        """
        try:
            departments = {}
            
            for employee in employees:
                # Используем первое непустое подразделение
                dept_name = None
                for dept in [employee['Подразделение 1'], employee['Подразделение 2'], employee['Подразделение 3'], employee['Подразделение 4']]:
                    if dept and dept.strip():
                        dept_name = dept.strip()
                        break
                
                if not dept_name:
                    dept_name = "Без подразделения"
                
                if dept_name not in departments:
                    departments[dept_name] = []
                
                departments[dept_name].append(employee)
            
            self.logger.info(f"Сотрудники сгруппированы по {len(departments)} подразделениям")
            return departments
            
        except Exception as e:
            self.logger.error(f"Ошибка группировки сотрудников: {e}")
            return {}

    def create_output_directory(self, base_path: str) -> str:
        """
        Создает выходную папку с timestamp
        
        Args:
            base_path: базовый путь для создания папки
            
        Returns:
            str: путь к созданной папке
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = Path(base_path) / f"vacation_files_{timestamp}"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Создана выходная папка: {output_dir}")
            return str(output_dir)
            
        except Exception as e:
            self.logger.error(f"Ошибка создания выходной папки: {e}")
            raise
    
    def create_or_use_department_structure(self, output_dir: str, employees: List[Dict[str, str]]) -> Dict[str, str]:
        """
        Создает структуру папок по подразделениям или использует существующие
        ВАЖНО: НЕ ИЗМЕНЯЕТ СТРУКТУРУ ФАЙЛОВ, только создает недостающие папки
        
        Args:
            output_dir: выходная папка
            employees: список сотрудников
            
        Returns:
            Dict[str, str]: словарь {название_подразделения: путь_к_папке}
        """
        departments = {}
        
        try:
            # Собираем уникальные подразделения из файла
            dept_set = set()
            for emp in employees:
                if emp['Подразделение 1']:
                    dept_set.add(emp['Подразделение 1'])
            
            # Проверяем существующие папки и создаем только недостающие
            output_path = Path(output_dir)
            
            for dept in dept_set:
                clean_dept_name = self._clean_directory_name(dept)
                dept_path = output_path / clean_dept_name
                
                # Создаем папку только если ее нет
                if not dept_path.exists():
                    dept_path.mkdir(parents=True, exist_ok=True)
                    self.logger.info(f"Создана папка отдела: {clean_dept_name}")
                else:
                    self.logger.info(f"Используется существующая папка: {clean_dept_name}")
                
                departments[dept] = str(dept_path)
            
            self.logger.info(f"Подготовлено отделов: {len(departments)}")
            return departments
            
        except Exception as e:
            self.logger.error(f"Ошибка подготовки структуры подразделений: {e}")
            raise
    
    def create_department_structure(self, output_dir: str, employees: List[Dict[str, str]]) -> Dict[str, str]:
        """
        Создает структуру папок по подразделениям
        
        Args:
            output_dir: выходная папка
            employees: список сотрудников
            
        Returns:
            Dict[str, str]: словарь {название_подразделения: путь_к_папке}
        """
        departments = {}
        
        try:
            # Собираем уникальные подразделения
            dept_set = set()
            for emp in employees:
                if emp['Подразделение 1']:
                    dept_set.add(emp['Подразделение 1'])
            
            # Создаем папки для каждого подразделения
            for dept in dept_set:
                clean_dept_name = self._clean_directory_name(dept)
                dept_path = Path(output_dir) / clean_dept_name
                dept_path.mkdir(parents=True, exist_ok=True)
                departments[dept] = str(dept_path)
            
            self.logger.info(f"Создано подразделений: {len(departments)}")
            return departments
            
        except Exception as e:
            self.logger.error(f"Ошибка создания структуры подразделений: {e}")
            raise
    
    def create_employee_files_with_skip(
        self, 
        employees: List[Dict[str, str]], 
        departments: Dict[str, str],
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        department_progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> Tuple[int, int]:
        """
        Создает файлы сотрудников с пропуском существующих
        
        Args:
            employees: список сотрудников
            departments: словарь папок подразделений
            progress_callback: функция обратного вызова для прогресса по файлам
            department_progress_callback: функция обратного вызова для прогресса по отделам
            
        Returns:
            Tuple[int, int]: количество успешно созданных файлов и пропущенных
        """
        total = len(employees)
        success_count = 0
        skipped_count = 0
        
        # Группируем сотрудников по отделам для прогресса
        employees_by_dept = {}
        for emp in employees:
            if emp['Подразделение 1'] not in employees_by_dept:
                employees_by_dept[emp['Подразделение 1']] = []
            employees_by_dept[emp['Подразделение 1']].append(emp)
        
        total_departments = len(employees_by_dept)
        processed_departments = 0
        processed_files_total = 0
        
        for dept_name, dept_employees in employees_by_dept.items():
            # Обновляем прогресс по отделам
            if department_progress_callback:
                department_progress_callback(processed_departments, total_departments, dept_name)
            
            dept_path = departments.get(dept_name)
            if not dept_path:
                self.logger.warning(f"Папка для подразделения {dept_name} не найдена")
                processed_departments += 1
                continue
            
            # Обрабатываем сотрудников в текущем отделе
            for i, employee in enumerate(dept_employees):
                try:
                    # Генерируем имя файла
                    filename = self.excel_handler.generate_output_filename(employee)
                    output_path = Path(dept_path) / filename
                    
                    # Проверяем существование файла 
                    if output_path.exists():
                        skipped_count += 1
                        message = f"Пропущен (уже существует): {employee['ФИО работника']}"
                    else:
                        # Создаем файл сотрудника
                        success = self.excel_handler.create_employee_file(employee, str(output_path))
                        
                        if success:
                            success_count += 1
                            message = f"Создан: {employee['ФИО работника']}"
                        else:
                            message = f"Ошибка создания: {employee['ФИО работника']}"
                    
                    processed_files_total += 1
                    
                    # Обновляем прогресс по файлам
                    if progress_callback:
                        progress_callback(processed_files_total, total, message)
                    
                    # Небольшая задержка для демонстрации прогресса
                    time.sleep(0.05)
                    
                except Exception as e:
                    self.logger.error(f"Ошибка создания файла для {employee['ФИО работника']}: {e}")
                    processed_files_total += 1
                    if progress_callback:
                        progress_callback(processed_files_total, total, f"Ошибка: {employee['ФИО работника']}")
            
            processed_departments += 1
            
            # Финальное обновление прогресса по отделам
            if department_progress_callback:
                department_progress_callback(processed_departments, total_departments, dept_name)
        
        self.logger.info(f"Создано файлов сотрудников: {success_count} сотр., пропущено: {skipped_count} из {total} сотр.")
        return success_count, skipped_count
    
    def create_employee_files(
        self, 
        employees: List[Dict[str, str]], 
        departments: Dict[str, str],
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        department_progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> int:
        """
        Создает файлы сотрудников (старый метод без пропуска)
        
        Args:
            employees: список сотрудников
            departments: словарь папок подразделений
            progress_callback: функция обратного вызова для прогресса по файлам
            department_progress_callback: функция обратного вызова для прогресса по отделам
            
        Returns:
            int: количество успешно созданных файлов
        """
        total = len(employees)
        success_count = 0
        
        # Группируем сотрудников по отделам для прогресса
        employees_by_dept = {}
        for emp in employees:
            if emp['Подразделение 1'] not in employees_by_dept:
                employees_by_dept[emp['Подразделение 1']] = []
            employees_by_dept[emp['Подразделение 1']].append(emp)
        
        total_departments = len(employees_by_dept)
        processed_departments = 0
        processed_files_total = 0
        
        for dept_name, dept_employees in employees_by_dept.items():
            # Обновляем прогресс по отделам
            if department_progress_callback:
                department_progress_callback(processed_departments, total_departments, dept_name)
            
            dept_path = departments.get(dept_name)
            if not dept_path:
                self.logger.warning(f"Папка для подразделения {dept_name} не найдена")
                processed_departments += 1
                continue
            
            # Обрабатываем сотрудников в текущем отделе
            for i, employee in enumerate(dept_employees):
                try:
                    # Генерируем имя файла
                    filename = self.excel_handler.generate_output_filename(employee)
                    output_path = Path(dept_path) / filename
                    
                    # Проверяем существование файла
                    if output_path.exists():
                        message = f"Пропущен (уже существует): {employee['ФИО работника']}"
                    else:
                        # Создаем файл сотрудника
                        success = self.excel_handler.create_employee_file(employee, str(output_path))
                        
                        if success:
                            success_count += 1
                            message = f"Создан: {employee['ФИО работника']}"
                        else:
                            message = f"Ошибка создания: {employee['ФИО работника']}"
                    
                    processed_files_total += 1
                    
                    # Обновляем прогресс по файлам
                    if progress_callback:
                        progress_callback(processed_files_total, total, message)
                    
                    # Небольшая задержка для демонстрации прогресса
                    time.sleep(0.05)
                    
                except Exception as e:
                    self.logger.error(f"Ошибка создания файла для {employee['ФИО работника']}: {e}")
                    processed_files_total += 1
                    if progress_callback:
                        progress_callback(processed_files_total, total, f"Ошибка: {employee['ФИО работника']}")
            
            processed_departments += 1
            
            # Финальное обновление прогресса по отделам
            if department_progress_callback:
                department_progress_callback(processed_departments, total_departments, dept_name)
        
        self.logger.info(f"Создано файлов сотрудников: {success_count} из {total}")
        return success_count
    
    def scan_employee_files(self, base_dir: str) -> Dict[str, List[str]]:
        """
        Сканирует файлы сотрудников в папках подразделений
        
        Args:
            base_dir: базовая папка для сканирования
            
        Returns:
            Dict[str, List[str]]: словарь {название_блока: [список_путей_к_файлам]}
        """
        result = {}
        
        try:
            base_path = Path(base_dir)
            if not base_path.exists():
                self.logger.warning(f"Папка не существует: {base_dir}")
                return result
            
            # Получаем список папок подразделений
            for entry in base_path.iterdir():
                if not entry.is_dir():
                    continue
                
                dept_name = entry.name
                employee_files = self._scan_department_files(entry)
                
                if employee_files:
                    result[dept_name] = employee_files
            
            self.logger.info(f"Найдено блоков: {len(result)}")
            return result
            
        except Exception as e:
            self.logger.error(f"Ошибка сканирования папки: {e}")
            return result
    
    def read_vacation_info_from_files(
        self, 
        files: List[str],
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[VacationInfo]:
        """
        Читает информацию об отпусках из файлов
        
        Args:
            files: список путей к файлам
            progress_callback: функция обратного вызова для прогресса
            
        Returns:
            List[VacationInfo]: список информации об отпусках
        """
        result = []
        total = len(files)
        
        for i, file_path in enumerate(files):
            try:
                vacation_info = self.excel_handler.read_vacation_info_from_file(file_path)
                
                if vacation_info:
                    result.append(vacation_info)
                    message = f"Обработан: {vacation_info.employee['ФИО работника']}"
                else:
                    message = f"Ошибка чтения: {Path(file_path).name}"
                
                if progress_callback:
                    progress_callback(i + 1, total, message)
                
            except Exception as e:
                self.logger.error(f"Ошибка чтения файла {file_path}: {e}")
                if progress_callback:
                    progress_callback(i + 1, total, f"Ошибка: {Path(file_path).name}")
        
        self.logger.info(f"Обработано файлов: {len(result)} из {total}")
        return result
    
    def _scan_department_files(self, dept_path: Path) -> List[str]:
        """
        Сканирует файлы сотрудников в папке подразделения
        
        Args:
            dept_path: путь к папке подразделения
            
        Returns:
            List[str]: список путей к файлам сотрудников
        """
        files = []
        
        try:
            for entry in dept_path.iterdir():
                if not entry.is_file():
                    continue
                
                filename = entry.name
                
                # Проверяем расширение и исключаем файлы отчетов
                if (entry.suffix.lower() == '.xlsx' and 
                    not self._is_report_file(filename)):
                    files.append(str(entry))
            
        except Exception as e:
            self.logger.error(f"Ошибка сканирования папки {dept_path}: {e}")
        
        return files
    
    def _is_report_file(self, filename: str) -> bool:
        """
        Проверяет, является ли файл отчетом
        
        Args:
            filename: имя файла
            
        Returns:
            bool: True если это файл отчета
        """
        report_indicators = [
            "!",
            "summary_",
            "GENERAL_",
            "report_",
            "отчет",
            "ОБЩИЙ",
            "общий",
            "Отчет по блоку"  # Добавляем индикатор для наших отчетов
        ]
        
        filename_lower = filename.lower()
        
        for indicator in report_indicators:
            if (filename.startswith(indicator) or 
                indicator.lower() in filename_lower):
                return True
        
        return False
    
    def _clean_directory_name(self, name: str) -> str:
        """
        Очищает имя папки от недопустимых символов
        
        Args:
            name: исходное имя
            
        Returns:
            str: очищенное имя
        """
        if not name:
            return "unnamed"
        
        # Заменяем недопустимые символы для имен папок
        invalid_chars = r'[<>:"/\\|?*]'
        import re
        clean_name = re.sub(invalid_chars, '_', name)
        
        # Убираем лишние пробелы и точки в конце
        clean_name = clean_name.strip('. ')
        
        # Ограничиваем длину
        if len(clean_name) > 100:
            clean_name = clean_name[:100]
        
        return clean_name or "unnamed"