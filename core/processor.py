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

import shutil
import re


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
        Создает файлы сотрудников в существующей папке
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
            
            # 3. Создание структуры папок
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
            total_success_count = 0
            total_skipped_count = 0
            total_error_count = 0
            
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
                
                # Счетчики для текущего отдела
                dept_success_count = 0
                dept_skipped_count = 0
                dept_error_count = 0
                
                # Обрабатываем сотрудников в текущем отделе
                for emp_idx, employee in enumerate(dept_employees):
                    try:
                        # Генерируем имя файла
                        filename = self.excel_handler.generate_output_filename(employee)
                        output_path = Path(dept_path) / filename
                        
                        # Проверяем существование файла
                        if output_path.exists():
                            dept_skipped_count += 1
                            total_skipped_count += 1
                            message = f"Пропущен: {employee.full_name}"
                        else:
                            # Создаем файл сотрудника
                            try:
                                create_success = self.excel_handler.create_employee_file(employee, str(output_path))
                                
                                if create_success:
                                    dept_success_count += 1
                                    total_success_count += 1
                                    message = f"Создан: {employee.full_name}"
                                else:
                                    dept_error_count += 1
                                    total_error_count += 1
                                    message = f"Ошибка: {employee.full_name}"
                                    operation_log.add_entry("ERROR", f"Не удалось создать файл для {employee.full_name}")
                            
                            except Exception as e:
                                dept_error_count += 1
                                total_error_count += 1
                                message = f"КРИТИЧЕСКАЯ ОШИБКА: {employee.full_name}"
                                operation_log.add_entry("ERROR", f"Критическая ошибка создания файла {employee.full_name}: {str(e)}")
                                
                                # Если есть ошибки в шаблоне - прерываем весь процесс
                                if "rules" in str(e).lower() or "шаблон" in str(e).lower() or "правил" in str(e).lower():
                                    operation_log.add_entry("ERROR", f"КРИТИЧЕСКАЯ ОШИБКА ШАБЛОНА: {str(e)}")
                                    operation_log.finish(ProcessingStatus.ERROR)
                                    return operation_log
                        
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
                        dept_error_count += 1
                        total_error_count += 1
                        self.logger.error(f"Ошибка создания файла для {employee.full_name}: {e}")
                        progress.processed_files += 1
                        if file_progress_callback:
                            file_progress_callback(emp_idx + 1, len(dept_employees), f"Ошибка: {employee.full_name}")
                
                # ИСПРАВЛЕНИЕ: Правильный подсчет по отделу
                if dept_success_count > 0 or dept_skipped_count > 0 or dept_error_count > 0:
                    operation_log.add_entry("INFO", f"Отдел {dept_name}: создано {dept_success_count}, пропущено {dept_skipped_count}")
                
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
            
            # Проверяем результат
            if total_error_count > 0:
                operation_log.add_entry("ERROR", f"Обработка завершена с ошибками: успешно {total_success_count}, ошибок {total_error_count}, пропущено {total_skipped_count}")
                operation_log.finish(ProcessingStatus.ERROR)
            else:
                operation_log.add_entry("INFO", f"Создано файлов сотрудников: {total_success_count}, пропущено: {total_skipped_count} из {total_employees}")
                operation_log.add_entry("INFO", f"Время выполнения: {duration.total_seconds():.1f} сек")
                operation_log.finish(ProcessingStatus.SUCCESS)
            
            return operation_log
            
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
            operation_log.add_entry("ERROR", error_msg)
            self.logger.error(error_msg, exc_info=True)
            operation_log.finish(ProcessingStatus.ERROR)
            
            return operation_log
        
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
                    employee_files = self.file_manager._scan_department_files(dept_path)
                    vacation_infos = []
                    
                    files_processed_in_dept = 0
                    for file_path in employee_files:
                        vacation_info = self.excel_handler.read_vacation_info_from_file(file_path)
                        if vacation_info:
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
            
            # 1. ПРЕДВАРИТЕЛЬНАЯ ПРОВЕРКА - все папки должны содержать отчеты
            progress.current_operation = "Проверка наличия отчетов по блокам"
            if progress_callback:
                progress_callback(progress)
            
            missing_reports = []
            multiple_reports_info = []
            
            for dept_info in selected_departments:
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                
                if not dept_path.exists():
                    missing_reports.append(dept_name)
                    continue
                    
                # Ищем отчеты в папке
                report_files = []
                for file_path in dept_path.iterdir():
                    if file_path.is_file() and file_path.suffix.lower() == '.xlsx':
                        filename = file_path.name
                        if (filename.startswith("Отчет по блоку") or 
                            filename.startswith("отчет по блоку") or
                            "отчет" in filename.lower()):
                            report_files.append(file_path)
                
                if not report_files:
                    missing_reports.append(dept_name)
                elif len(report_files) > 1:
                    # Находим самый новый отчет
                    latest_file = max(report_files, key=lambda f: f.stat().st_mtime)
                    
                    # Извлекаем дату из названия файла
                    date_match = re.search(r'(\d{8}_\d{6})', latest_file.name)
                    if date_match:
                        date_str = date_match.group(1)
                        try:
                            parsed_date = datetime.strptime(date_str, "%Y%m%d_%H%M%S")
                            date_display = parsed_date.strftime("%d.%m.%Y %H:%M")
                        except ValueError:
                            date_display = date_str
                    else:
                        date_display = "неизвестная дата"
                    
                    multiple_reports_info.append({
                        'dept_name': dept_name,
                        'count': len(report_files),
                        'selected_file': latest_file.name,
                        'date_display': date_display
                    })
            
            # Если есть отсутствующие отчеты - прерываем
            if missing_reports:
                missing_deps_str = ", ".join(missing_reports)
                error_msg = f"Не найдены отчеты по блокам для подразделений: {missing_deps_str}"
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            # Логируем информацию о множественных отчетах
            for info in multiple_reports_info:
                log_msg = f"В отделе '{info['dept_name']}' найдено {info['count']} отчетов по блоку, будет использован отчет '{info['selected_file']}', так как в его названии самая актуальная дата ({info['date_display']})"
                operation_log.add_entry("INFO", log_msg)
            
            # 2. РАСКИДЫВАНИЕ create_report.exe ПО ПАПКАМ
            progress.current_operation = "Раскидывание скриптов по отделам"
            if progress_callback:
                progress_callback(progress)
            
            # Ищем create_report.exe в папке приложения
            exe_source_path = None
            possible_paths = [
                Path("create_report.exe"),
                Path("dist/create_report.exe"),
                Path("build/create_report.exe"),
                Path("release/create_report.exe")
            ]
            
            for path in possible_paths:
                if path.exists():
                    exe_source_path = path
                    break
            
            if not exe_source_path:
                error_msg = "Файл create_report.exe не найден для раскидывания по отделам"
                operation_log.add_entry("ERROR", error_msg)
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            # Раскидываем exe по папкам
            for dept_info in selected_departments:
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                
                if dept_path.exists():
                    clean_dept_name = self._clean_filename_for_exe(dept_name)
                    target_filename = f"Скрипт для сборки отчета по блоку '{clean_dept_name}'.exe"
                    target_path = dept_path / target_filename
                    
                    try:
                        shutil.copy2(exe_source_path, target_path)
                        operation_log.add_entry("INFO", f"Скрипт скопирован в {dept_name}")
                    except Exception as e:
                        operation_log.add_entry("ERROR", f"Ошибка копирования скрипта в {dept_name}: {e}")
            
            # 3. СБОР ДАННЫХ ИЗ ОТЧЕТОВ
            progress.current_operation = "Сбор данных из отчетов по блокам"
            if progress_callback:
                progress_callback(progress)
            
            block_data = []
            import random
            
            for i, dept_info in enumerate(selected_departments):
                dept_name = dept_info['name']
                dept_path = Path(dept_info['path'])
                
                # Генерируем случайное время для этого блока (1.5-3 сек)
                block_processing_time = random.uniform(1.5, 3.0)
                block_start_time = time.time()
                
                progress.current_operation = f"Обработка отчета: {dept_name}"
                progress.current_block = dept_name
                progress.processed_blocks = i
                progress.processed_files = i
                if progress_callback:
                    progress_callback(progress)
                
                # Ищем отчет (уже проверили что он есть)
                block_report_path = self._find_latest_block_report(str(dept_path), dept_name)
                
                if block_report_path:
                    # Читаем данные из отчета
                    block_info = self.excel_handler.read_block_report_data(block_report_path)
                    if block_info:
                        block_data.append(block_info)
                        operation_log.add_entry("INFO", f"Данные собраны из отчета для {dept_name}")
                    else:
                        operation_log.add_entry("ERROR", f"Ошибка чтения данных из отчета для {dept_name}")
                
                # Эмуляция обработки с прогресс-баром
                while time.time() - block_start_time < block_processing_time:
                    time.sleep(0.1)
                    # Прогресс внутри блока обновляется в reports_window.py
                    if progress_callback:
                        progress_callback(progress)
            
            # Завершаем обработку блоков
            progress.processed_blocks = len(selected_departments)
            progress.processed_files = len(selected_departments)
            if progress_callback:
                progress_callback(progress)
            
            # 4. СОЗДАНИЕ ОБЩЕГО ОТЧЕТА
            progress.current_operation = "Создание файла общего отчета"
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