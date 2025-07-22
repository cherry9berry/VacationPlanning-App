#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Основной процессор для обработки файлов отпусков
"""

import logging
import time
import random
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Callable, Optional, Tuple

from models import (
    Employee, VacationInfo, BlockReport, GeneralReport, 
    ProcessingProgress, OperationLog, ProcessingStatus, ValidationResult, VacationStatus
)
from config import Config
from core.validator import Validator
from core.excel_handler import ExcelHandler
from core.file_manager import FileManager
from core.performance_tracker import PerformanceTracker
from core.employee_file_creator import EmployeeFileCreator
from core.directory_manager import DirectoryManager

import shutil


class VacationProcessor:
    """Основной класс для обработки операций с отпусками"""
    
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.validator = Validator(config)
        self.excel_handler = ExcelHandler(config)
        self.file_manager = FileManager(config)
        self.performance_tracker = PerformanceTracker()
        self.employee_file_creator = EmployeeFileCreator(config)
        self.directory_manager = DirectoryManager(config)

    def create_employee_files_to_existing(
            self, 
            staff_file_path: str, 
            target_directory: str,
            progress_callback: Optional[Callable[[ProcessingProgress], None]] = None,
            department_progress_callback: Optional[Callable[[int, int, str], None]] = None,
            file_progress_callback: Optional[Callable[[int, int, str], None]] = None,
            employees_to_create: Optional[list] = None
        ) -> OperationLog:
            """
            Создает файлы сотрудников в существующей папке
            Делегирует выполнение специализированному классу EmployeeFileCreator
            """
            return self.employee_file_creator.create_employee_files(
                staff_file_path=staff_file_path,
                target_directory=target_directory,
                progress_callback=progress_callback,
                department_progress_callback=department_progress_callback,
                file_progress_callback=file_progress_callback,
                employees_to_create=employees_to_create
            )

    def _clean_filename_for_exe(self, filename: str) -> str:
        """Очищает имя файла для exe от недопустимых символов"""
        if not filename:
            return "unnamed"
        
        # Заменяем недопустимые символы
        clean_name = re.sub(r'[\\/:*?"<>|]', '_', filename)
        clean_name = clean_name.strip('. ')
        
        # Ограничиваем длину
        if len(clean_name) > 80:
            clean_name = clean_name[:80]
        
        return clean_name or "unnamed"

    def scan_target_directory(self, target_directory: str) -> Dict[str, int]:
        """
        Сканирует целевую папку и возвращает информацию о подразделениях
        
        Args:
            target_directory: путь к папке для сканирования
            
        Returns:
            Dict[str, int]: {название_подразделения: количество_файлов}
        """
        try:
            departments = self.directory_manager.scan_existing_departments(target_directory)
            
            # Подсчитываем файлы в каждом подразделении
            departments_info = {}
            for dept_name, dept_path in departments.items():
                files = self.directory_manager._scan_department_files(Path(dept_path))
                departments_info[dept_name] = len(files)
            
            return departments_info
            
        except Exception as e:
            self.logger.error(f"Ошибка сканирования папки {target_directory}: {e}")
            return {}

    def update_block_reports(
        self,
        selected_departments: List[Dict],
        progress_callback: Optional[Callable[[ProcessingProgress], None]] = None
    ) -> OperationLog:
        """
        Обновляет отчеты по выбранным подразделениям
        
        Args:
            selected_departments: список выбранных подразделений в формате [{'name': str, 'path': str, 'files_count': int}]
            progress_callback: функция для обновления прогресса
            
        Returns:
            OperationLog: лог операции
        """
        operation_log = OperationLog("Обновление отчетов по подразделениям")
        operation_log.add_entry("INFO", "Начало обновления отчетов по блокам")
        
        try:
            start_time = datetime.now()
            total_files = sum(dept['files_count'] for dept in selected_departments)
            progress = ProcessingProgress(
                current_operation="Подготовка к созданию отчетов",
                start_time=start_time,
                total_blocks=len(selected_departments),
                total_files=total_files
            )
            
            if progress_callback:
                progress_callback(progress)
            
            success_count = 0
            error_count = 0
            files_processed_total = 0
            
            for i, dept_info in enumerate(selected_departments):
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                files_in_dept = dept_info['files_count']
                
                progress.current_operation = f"Создание отчета: {dept_name}"
                progress.current_block = dept_name
                progress.processed_blocks = i
                if progress_callback:
                    progress_callback(progress)
                
                try:
                    if not dept_path.exists():
                        error_msg = f"Папка подразделения не найдена: {dept_name}"
                        operation_log.add_entry("ERROR", error_msg)
                        error_count += 1
                        continue
                    
                    # Читаем файлы сотрудников
                    employee_files = self.directory_manager._scan_department_files(dept_path)
                    vacation_infos = []
                    
                    files_processed_in_dept = 0
                    for file_path in employee_files:
                        # Дополнительная проверка - исключаем отчеты и временные файлы
                        filename = Path(file_path).name
                        if (filename.startswith('~$') or 
                            filename.startswith('Отчет') or 
                            filename.startswith('отчет') or 
                            'отчет' in filename.lower() or
                            filename.startswith('ОБЩИЙ_ОТЧЕТ') or
                            filename.startswith('общий_отчет')):
                            continue
                        
                        vacation_info = self.excel_handler.read_vacation_info_from_file(file_path)
                        if vacation_info and vacation_info.employee.get('ФИО работника'):
                            vacation_infos.append(vacation_info)
                        
                        files_processed_in_dept += 1
                        files_processed_total += 1
                        progress.processed_files = files_processed_total
                        
                        # ИСПРАВЛЕНИЕ: Обновляем прогресс для каждого файла
                        if progress_callback:
                            progress_callback(progress)
                        
                        time.sleep(0.1)  # Небольшая задержка для демонстрации прогресса
                    
                    # Создаем отчет
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    report_filename = f"Отчет по блоку_{dept_name}_{timestamp}.xlsx"
                    report_path = dept_path / report_filename
                    
                    success = self.excel_handler.create_block_report(
                        dept_name, vacation_infos, str(report_path)
                    )
                    
                    if success:
                        success_count += 1
                        # ИСПРАВЛЕНИЕ: Убираем уровень success для ИТОГ сообщений - просто INFO
                        operation_log.add_entry("INFO", f"Создан отчет: {dept_name}")
                    else:
                        error_count += 1
                        operation_log.add_entry("ERROR", f"Ошибка создания отчета: {dept_name}")
                
                except Exception as e:
                    error_count += 1
                    error_msg = f"Ошибка обработки {dept_name}: {e}"
                    operation_log.add_entry("ERROR", error_msg)
                    self.logger.error(error_msg)
                
                # Обновляем прогресс по блокам
                progress.processed_blocks = i + 1
                if progress_callback:
                    progress_callback(progress)
            
            # Завершение
            end_time = datetime.now()
            duration = end_time - start_time
            
            progress.current_operation = "Отчеты созданы"
            progress.end_time = end_time
            if progress_callback:
                progress_callback(progress)
            
            operation_log.add_entry("INFO", f"Создание отчетов завершено за {duration.total_seconds():.1f} сек")
            operation_log.add_entry("INFO", f"Успешно: {success_count}, Ошибок: {error_count}")
            operation_log.finish(ProcessingStatus.SUCCESS)
            
            return operation_log
            
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
            operation_log.add_entry("ERROR", error_msg)
            self.logger.error(error_msg, exc_info=True)
            operation_log.finish(ProcessingStatus.ERROR)
            
            return operation_log

    def create_general_report(
        self,
        selected_departments: List[Dict],
        base_directory: str,
        progress_callback: Optional[Callable[[ProcessingProgress], None]] = None
    ) -> OperationLog:
        """
        Создает общий отчет по выбранным подразделениям
        Использует ТОЛЬКО блочные отчеты (Отчет по блоку_*), не трогает файлы сотрудников
        """
        operation_log = OperationLog("Создание общего отчета")
        try:
            start_time = datetime.now()
            progress = ProcessingProgress(
                current_operation="Подготовка к созданию общего отчета",
                start_time=start_time,
                total_blocks=len(selected_departments),
                total_files=len(selected_departments)
            )
            if progress_callback:
                progress_callback(progress)
            # 1. Проверка наличия отчетов по блокам
            missing_reports = []
            block_data = []
            all_totals = []
            for dept_info in selected_departments:
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                print(f"DEBUG: Ищу отчет для отдела: {dept_name}, путь: {dept_path}")
                if not dept_path.exists():
                    print(f"DEBUG: Папка не существует: {dept_path}")
                    missing_reports.append(dept_name)
                    continue
                block_report_path = self._find_latest_block_report(str(dept_path), dept_name)
                print(f"DEBUG: Найден файл отчета: {block_report_path}")
                if not block_report_path:
                    print(f"DEBUG: Не найден отчет для отдела: {dept_name}")
                    missing_reports.append(dept_name)
                    continue
                block_info_raw = self.excel_handler.read_block_report_data_by_rules(block_report_path)
                print(f"DEBUG: Данные из отчета {block_report_path}: {block_info_raw}")
                if not block_info_raw:
                    print(f"DEBUG: Не удалось прочитать отчет для отдела: {dept_name}")
                    missing_reports.append(dept_name)
                    continue
                all_totals.append(int(block_info_raw.get('total_employees', 0)))
                block_data.append((dept_name, block_info_raw))
            if missing_reports:
                missing_deps_str = ", ".join(missing_reports)
                error_msg = f"Не найдены валидные блочные отчеты для подразделений: {missing_deps_str}"
                print(f"DEBUG: {error_msg}")
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            total_employees_all = sum(int(b[1].get('total_employees', 0)) for b in block_data)
            print(f"DEBUG: total_employees_all = {total_employees_all}")
            # Формируем финальный список для общего отчета
            final_block_data = []
            for i, (dept_name, block_info_raw) in enumerate(block_data):
                total = int(block_info_raw.get('total_employees', 0))
                correct = int(block_info_raw.get('completed_employees', 0))
                incorrect = int(block_info_raw.get('employees_incorrect', 0))
                not_filled = int(block_info_raw.get('employees_not_filled', 0))
                print(f"DEBUG: Формирую строку для блока {dept_name}: total={total}, correct={correct}, incorrect={incorrect}, not_filled={not_filled}")
                block_info = {
                    'row_number2': i + 1,
                    'report_department1': dept_name,
                    'employees_count_percent': total / total_employees_all if total_employees_all > 0 else 0.0,
                    'employees_count': total,
                    'correct_filled_percent': correct / total if total > 0 else 0.0,
                    'correct_filled': correct,
                    'incorrect_filled_percent': incorrect / total if total > 0 else 0.0,
                    'incorrect_filled': incorrect,
                    'not_filled_percent': not_filled / total if total > 0 else 0.0,
                    'not_filled': not_filled,
                    'update_date': block_info_raw.get('update_date', ''),
                }
                print(f"DEBUG: block_info для {dept_name}: {block_info}")
                final_block_data.append(block_info)
            print(f"DEBUG: Финальные данные для общего отчета: {final_block_data}")
            # 2. Создание общего отчета
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_filename = f"ОБЩИЙ_ОТЧЕТ_{timestamp}.xlsx"
            report_path = Path(base_directory) / report_filename
            print(f"DEBUG: Путь итогового отчета: {report_path}")
            success = self.excel_handler.create_general_report_from_blocks(
                final_block_data, str(report_path)
            )
            if success:
                end_time = datetime.now()
                duration = end_time - start_time
                total_employees_all = sum(b['employees_count'] for b in final_block_data if 'employees_count' in b)
                total_correct_all = sum(b['correct_filled'] for b in final_block_data if 'correct_filled' in b)
                operation_log.add_entry("INFO", f"Общий отчет создан: {report_path}")
                operation_log.add_entry("INFO", f"Блоков: {len(final_block_data)}, Сотрудников: {total_employees_all}, Заполнили корректно: {total_correct_all}")
                operation_log.add_entry("INFO", f"Время выполнения: {duration.total_seconds():.1f} сек")
                operation_log.finish(ProcessingStatus.SUCCESS)
            else:
                error_msg = "Ошибка создания общего отчета"
                print(f"DEBUG: {error_msg}")
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
            return operation_log
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
            print(f"DEBUG: {error_msg}")
            operation_log.add_entry("ERROR", error_msg)
            self.logger.error(error_msg, exc_info=True)
            operation_log.finish(ProcessingStatus.ERROR)
            return operation_log

    def _find_latest_block_report(self, dept_path: str, dept_name: str) -> Optional[str]:
        """
        Находит последний отчет по блоку для подразделения
        
        Args:
            dept_path: путь к папке подразделения
            dept_name: название подразделения
            
        Returns:
            Optional[str]: путь к файлу отчета или None
        """
        try:
            dept_path_obj = Path(dept_path)
            
            if not dept_path_obj.exists():
                self.logger.error(f"Папка подразделения не существует: {dept_path_obj}")
                return None
            
            # Ищем файлы отчетов
            report_files = []
            for file_path in dept_path_obj.iterdir():
                if file_path.is_file() and file_path.suffix.lower() == '.xlsx':
                    filename = file_path.name
                    if (filename.startswith("Отчет по блоку") or 
                        filename.startswith("отчет по блоку") or
                        "отчет" in filename.lower()):
                        report_files.append(file_path)
            
            if not report_files:
                self.logger.warning(f"Отчеты по блоку не найдены в {dept_path_obj}")
                return None
            
            # Возвращаем самый новый файл
            latest_file = max(report_files, key=lambda f: f.stat().st_mtime)
            return str(latest_file)
            
        except Exception as e:
            self.logger.error(f"Ошибка поиска отчета для {dept_name}: {e}")
            return None

    def _is_report_file(self, filename: str) -> bool:
        """Проверяет, является ли файл отчетом"""
        report_indicators = [
            "Отчет", "отчет", "ОБЩИЙ", "общий", "GENERAL", "summary_", "!"
        ]
        filename_lower = filename.lower()
        return any(indicator.lower() in filename_lower for indicator in report_indicators)