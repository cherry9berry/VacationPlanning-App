#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Основной процессор для обработки файлов отпусков
"""

import logging
import time
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Callable, Optional, Tuple

from models import (
    Employee, VacationInfo, BlockReport, GeneralReport, 
    ProcessingProgress, OperationLog, ProcessingStatus, ValidationResult
)
from config import Config
from core.validator import Validator
from core.excel_handler import ExcelHandler
from core.file_manager import FileManager


class VacationProcessor:
    """Основной класс для обработки операций с отпусками"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.validator = Validator(config)
        self.excel_handler = ExcelHandler(config)
        self.file_manager = FileManager(config)
    
    def create_employee_files_to_existing(
        self, 
        staff_file_path: str, 
        target_directory: str,
        progress_callback: Optional[Callable[[ProcessingProgress], None]] = None,
        department_progress_callback: Optional[Callable[[int, int, str], None]] = None,
        file_progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> OperationLog:
        """
        Создает файлы сотрудников в существующей папке (без создания новой структуры vacation_files_*)
        
        Args:
            staff_file_path: путь к файлу штатного расписания
            target_directory: существующая целевая папка для сохранения
            progress_callback: функция для обновления общего прогресса
            department_progress_callback: функция для обновления прогресса по отделам (current, total, name)
            file_progress_callback: функция для обновления прогресса по файлам (current, total, name)
            
        Returns:
            OperationLog: лог операции
        """
        operation_log = OperationLog("Создание файлов сотрудников в существующей структуре")
        operation_log.add_entry("INFO", "Начало создания файлов сотрудников")
        
        try:
            start_time = datetime.now()
            progress = ProcessingProgress(
                current_operation="Начало обработки",
                start_time=start_time
            )
            
            if progress_callback:
                progress_callback(progress)
            
            # 1. Валидация файла штатного расписания
            progress.current_operation = "Валидация файла штатного расписания"
            progress.current_file = Path(staff_file_path).name
            if progress_callback:
                progress_callback(progress)
            
            validation_result, employees = self.validator.validate_staff_file(staff_file_path)
            if not validation_result.is_valid:
                operation_log.add_entry("ERROR", f"Валидация не пройдена: {validation_result.errors}")
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            operation_log.add_entry("INFO", f"Валидация пройдена. Найдено сотрудников: {len(employees)}")
            
            # 2. Группировка сотрудников по отделам
            progress.current_operation = "Группировка сотрудников по отделам"
            if progress_callback:
                progress_callback(progress)
            
            employees_by_dept = self.file_manager.group_employees_by_department(employees)
            
            # 3. Создание структуры папок (используем существующую папку)
            progress.current_operation = "Подготовка структуры папок"
            if progress_callback:
                progress_callback(progress)
            
            departments = self.file_manager.create_or_use_department_structure(target_directory, employees)
            
            # 4. Подготовка прогресса
            total_departments = len(employees_by_dept)
            total_employees = len(employees)
            
            progress.total_blocks = total_departments
            progress.total_files = total_employees
            progress.processed_blocks = 0
            progress.processed_files = 0
            
            if progress_callback:
                progress_callback(progress)
            
            # 5. Создание файлов по отделам
            success_count = 0
            skipped_count = 0
            
            for dept_idx, (dept_name, dept_employees) in enumerate(employees_by_dept.items()):
                progress.current_operation = f"Обработка отдела: {dept_name}"
                progress.current_block = dept_name
                progress.processed_blocks = dept_idx
                
                if department_progress_callback:
                    department_progress_callback(dept_idx, total_departments, dept_name)
                
                if progress_callback:
                    progress_callback(progress)
                
                dept_path = departments.get(dept_name)
                if not dept_path:
                    operation_log.add_entry("ERROR", f"Папка для отдела {dept_name} не найдена")
                    continue
                
                # Обрабатываем сотрудников в текущем отделе
                dept_success = 0
                dept_skipped = 0
                
                for emp_idx, employee in enumerate(dept_employees):
                    try:
                        # Генерируем имя файла
                        filename = self.excel_handler.generate_output_filename(employee)
                        output_path = Path(dept_path) / filename
                        
                        # Проверяем существование файла
                        if output_path.exists():
                            skipped_count += 1
                            dept_skipped += 1
                            message = f"Пропущен: {employee.full_name}"
                        else:
                            # Создаем файл сотрудника
                            create_success = self.excel_handler.create_employee_file(employee, str(output_path))
                            
                            if create_success:
                                success_count += 1
                                dept_success += 1
                                message = f"Создан: {employee.full_name}"
                            else:
                                message = f"Ошибка: {employee.full_name}"
                        
                        progress.processed_files += 1
                        
                        # Обновляем прогресс по файлам в отделе
                        if file_progress_callback:
                            file_progress_callback(emp_idx + 1, len(dept_employees), message)
                        
                        # Обновляем общий прогресс
                        if progress_callback:
                            progress_callback(progress)
                        
                        # Небольшая задержка
                        time.sleep(0.05)
                        
                    except Exception as e:
                        self.logger.error(f"Ошибка создания файла для {employee.full_name}: {e}")
                        progress.processed_files += 1
                        if file_progress_callback:
                            file_progress_callback(emp_idx + 1, len(dept_employees), f"Ошибка: {employee.full_name}")
                
                # Логируем результат по отделу
                operation_log.add_entry("INFO", f"Отдел {dept_name}: создано {dept_success}, пропущено {dept_skipped}")
                
                # Завершаем обработку отдела
                progress.processed_blocks = dept_idx + 1
                
                if department_progress_callback:
                    department_progress_callback(dept_idx + 1, total_departments, dept_name)
            
            # 6. Завершение
            end_time = datetime.now()
            duration = end_time - start_time
            
            progress.current_operation = "Создание файлов завершено"
            progress.end_time = end_time
            if progress_callback:
                progress_callback(progress)
            
            operation_log.add_entry("INFO", f"Создано файлов сотрудников: {success_count}, пропущено: {skipped_count} из {total_employees}")
            operation_log.add_entry("INFO", f"Время выполнения: {duration.total_seconds():.1f} сек")
            operation_log.finish(ProcessingStatus.SUCCESS)
            
            return operation_log
            
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
            operation_log.add_entry("ERROR", error_msg)
            self.logger.error(error_msg, exc_info=True)
            operation_log.finish(ProcessingStatus.ERROR)
            
            return operation_log
    
    def scan_target_directory(self, target_directory: str) -> Dict[str, int]:
        """
        Сканирует целевую папку и возвращает информацию о подразделениях
        
        Args:
            target_directory: путь к папке для сканирования
            
        Returns:
            Dict[str, int]: {название_подразделения: количество_файлов}
        """
        try:
            departments = self.file_manager.scan_existing_departments(target_directory)
            
            # Подсчитываем файлы в каждом подразделении
            departments_info = {}
            for dept_name, dept_path in departments.items():
                files = self.file_manager._scan_department_files(Path(dept_path))
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
            progress = ProcessingProgress(
                current_operation="Подготовка к созданию отчетов",
                start_time=start_time,
                total_blocks=len(selected_departments),
                total_files=sum(dept['files_count'] for dept in selected_departments)
            )
            
            if progress_callback:
                progress_callback(progress)
            
            success_count = 0
            error_count = 0
            
            for i, dept_info in enumerate(selected_departments):
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                
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
                    employee_files = self.file_manager._scan_department_files(dept_path)
                    vacation_infos = []
                    
                    files_processed = 0
                    for file_path in employee_files:
                        vacation_info = self.excel_handler.read_vacation_info_from_file(file_path)
                        if vacation_info:
                            vacation_infos.append(vacation_info)
                        
                        files_processed += 1
                        progress.processed_files = progress.processed_files + 1
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
        
        Args:
            selected_departments: список выбранных подразделений в формате [{'name': str, 'path': str, 'files_count': int}]
            base_directory: базовая папка для сохранения общего отчета
            progress_callback: функция для обновления прогресса
            
        Returns:
            OperationLog: лог операции
        """
        operation_log = OperationLog("Создание общего отчета")
        operation_log.add_entry("INFO", "Начало создания общего отчета")
        
        try:
            start_time = datetime.now()
            progress = ProcessingProgress(
                current_operation="Подготовка к созданию общего отчета",
                start_time=start_time,
                total_blocks=len(selected_departments),
                total_files=len(selected_departments)  # Каждый отдел = один файл для анализа
            )
            
            if progress_callback:
                progress_callback(progress)
            
            # Проверяем наличие отчетов по блокам и собираем данные
            block_data = []
            missing_reports = []
            
            for i, dept_info in enumerate(selected_departments):
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                
                progress.current_operation = f"Проверка отчета: {dept_name}"
                progress.current_block = dept_name
                progress.processed_blocks = i
                progress.processed_files = i
                if progress_callback:
                    progress_callback(progress)
                
                # Ищем последний отчет по блоку
                block_report_path = self._find_latest_block_report(str(dept_path), dept_name)
                
                if block_report_path:
                    # Читаем данные из отчета
                    block_info = self.excel_handler.read_block_report_data(block_report_path)
                    if block_info:
                        block_data.append(block_info)
                        operation_log.add_entry("INFO", f"Найден отчет для {dept_name}")
                    else:
                        missing_reports.append(dept_name)
                        operation_log.add_entry("ERROR", f"Ошибка чтения отчета для {dept_name}")
                else:
                    missing_reports.append(dept_name)
                    operation_log.add_entry("ERROR", f"Отчет по блоку не найден для {dept_name}")
                    self.logger.error(f"Отчет по блоку не найден для подразделения: {dept_name}")
                
                time.sleep(0.2)  # Небольшая задержка для демонстрации прогресса
            
            # Если есть отсутствующие отчеты - прерываем выполнение
            if missing_reports:
                missing_deps_str = ", ".join(missing_reports)
                error_msg = f"Не найдены отчеты по блокам для подразделений: {missing_deps_str}"
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            # Создаем общий отчет
            progress.current_operation = "Создание общего отчета"
            progress.processed_files = len(selected_departments)
            if progress_callback:
                progress_callback(progress)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_filename = f"ОБЩИЙ_ОТЧЕТ_{timestamp}.xlsx"
            report_path = Path(base_directory) / report_filename
            
            success = self.excel_handler.create_general_report_from_blocks(
                block_data, str(report_path)
            )
            
            if success:
                end_time = datetime.now()
                duration = end_time - start_time
                
                progress.current_operation = "Общий отчет создан"
                progress.end_time = end_time
                if progress_callback:
                    progress_callback(progress)
                
                operation_log.add_entry("INFO", f"Общий отчет создан: {report_path}")
                operation_log.add_entry("INFO", f"Время выполнения: {duration.total_seconds():.1f} сек")
                operation_log.finish(ProcessingStatus.SUCCESS)
                
                self.logger.info(f"Общий отчет создан: {report_path}")
                
            else:
                error_msg = "Ошибка создания общего отчета"
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
            
            return operation_log
            
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
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