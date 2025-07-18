#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Специализированный класс для создания файлов сотрудников
"""

import logging
import time
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Callable, Optional, Tuple

from models import ProcessingProgress, OperationLog, ProcessingStatus
from config import Config
from core.events import event_bus, EventType
from core.validator import Validator
from core.excel_handler import ExcelHandler
from core.directory_manager import DirectoryManager
from core.transaction_manager import TransactionManager


class EmployeeFileCreator:
    """Класс для создания файлов сотрудников"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.validator = Validator(config)
        self.excel_handler = ExcelHandler(config)
        self.directory_manager = DirectoryManager(config)
        self.transaction_manager = TransactionManager()
    
    def create_employee_files(
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
        """
        operation_log = OperationLog("Создание файлов сотрудников")
        operation_log.add_entry("INFO", "Начало создания файлов сотрудников")
        
        try:
            start_time = datetime.now()
            progress = ProcessingProgress(
                current_operation="Начало обработки",
                start_time=start_time
            )
            
            self._emit_progress_update(progress, progress_callback)
            
            # Начинаем транзакцию
            backup_dir = str(Path(target_directory).parent / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            if not self.transaction_manager.begin_transaction(backup_dir):
                operation_log.add_entry("ERROR", "Не удалось начать транзакцию")
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            operation_log.add_entry("INFO", f"Транзакция начата с резервным копированием в {backup_dir}")
            
            # 1. Валидация файла штатного расписания
            progress.current_operation = "Валидация файла штатного расписания"
            progress.current_file = Path(staff_file_path).name
            self._emit_progress_update(progress, progress_callback)
            
            validation_result, employees = self.validator.validate_staff_file(staff_file_path)
            self.logger.info(f"Валидация завершена, найдено {len(employees)} сотрудников")
            
            if not validation_result.is_valid:
                operation_log.add_entry("ERROR", f"Валидация не пройдена: {validation_result.errors}")
                operation_log.finish(ProcessingStatus.ERROR)
                return operation_log
            
            operation_log.add_entry("INFO", f"Валидация пройдена. Найдено сотрудников: {len(employees)}")
            
            # 2. Используем только сотрудников для создания, если передан список
            if employees_to_create is not None:
                employees = employees_to_create
            
            # 3. Создание структуры папок через DirectoryManager
            progress.current_operation = "Подготовка структуры папок"
            self._emit_progress_update(progress, progress_callback)
            
            departments = self.directory_manager.create_department_structure(
                target_directory, employees
            )
            
            # 4. Подготовка прогресса
            employees_by_dept = self._group_employees_by_department(employees)
            total_departments = len(employees_by_dept)
            total_employees = len(employees)
            
            progress.total_blocks = total_departments
            progress.total_files = total_employees
            progress.processed_blocks = 0
            progress.processed_files = 0
            
            self.logger.debug(f"Подготовлено {total_departments} отделов, {total_employees} сотрудников")
            self._emit_progress_update(progress, progress_callback)
            
            # 5. Начинаем отслеживание производительности
            self.excel_handler.performance_tracker.start_batch()
            
            # 6. Создание файлов по отделам
            self.logger.debug("Начинаем создание файлов")
            total_success_count = 0
            total_error_count = 0
            
            for dept_idx, (dept_name, dept_employees) in enumerate(employees_by_dept.items()):
                # Проверка на остановку через события
                if self._should_stop():
                    operation_log.add_entry("INFO", "Операция остановлена пользователем")
                    operation_log.finish(ProcessingStatus.CANCELLED)
                    return operation_log
                
                progress.current_operation = f"Обработка отдела: {dept_name}"
                progress.current_block = dept_name
                progress.processed_blocks = dept_idx
                
                if department_progress_callback:
                    department_progress_callback(dept_idx, total_departments, dept_name)
                
                self._emit_progress_update(progress, progress_callback)
                
                dept_path = departments.get(dept_name)
                if not dept_path:
                    operation_log.add_entry("ERROR", f"Папка для отдела {dept_name} не найдена")
                    continue
                
                # Счетчики для текущего отдела
                dept_success_count = 0
                dept_error_count = 0
                
                # Обрабатываем сотрудников в текущем отделе
                for emp_idx, employee in enumerate(dept_employees):
                    if self._should_stop():
                        operation_log.add_entry("INFO", "Операция остановлена пользователем")
                        operation_log.finish(ProcessingStatus.CANCELLED)
                        return operation_log
                    
                    try:
                        # Генерируем имя файла
                        filename = self.excel_handler.generate_output_filename(employee)
                        output_path = Path(dept_path) / filename
                        
                        # Добавляем операцию в транзакцию
                        self.transaction_manager.add_file_creation(str(output_path), employee)
                        
                        # Проверяем, существует ли файл уже
                        if output_path.exists():
                            self.logger.info(f"Файл уже существует, пропускаем: {filename}")
                            dept_success_count += 1
                            total_success_count += 1
                            message = f"Пропущен (существует): {employee['ФИО работника']}"
                            
                            # Отправляем событие о пропуске файла
                            event_bus.emit_simple(
                                EventType.FILE_CREATED,
                                {"file_path": str(output_path), "employee": employee, "skipped": True},
                                "EmployeeFileCreator"
                            )
                            success = True
                        else:
                            # Создаем файл сотрудника
                            success = self.excel_handler.create_employee_file(employee, str(output_path))
                        
                        if success:
                            dept_success_count += 1
                            total_success_count += 1
                            message = f"Создан: {employee['ФИО работника']}"
                            
                            # Отправляем событие о создании файла
                            event_bus.emit_simple(
                                EventType.FILE_CREATED,
                                {"file_path": str(output_path), "employee": employee},
                                "EmployeeFileCreator"
                            )
                        else:
                            dept_error_count += 1
                            total_error_count += 1
                            message = f"Ошибка создания: {employee['ФИО работника']}"
                            
                            # Отправляем событие об ошибке
                            event_bus.emit_simple(
                                EventType.ERROR_OCCURRED,
                                {"error": "Ошибка создания файла", "employee": employee},
                                "EmployeeFileCreator"
                            )
                        
                        progress.processed_files += 1
                        
                        # Обновляем прогресс по файлам в отделе
                        if file_progress_callback:
                            file_progress_callback(emp_idx + 1, len(dept_employees), message)
                        
                        self._emit_progress_update(progress, progress_callback)
                        
                        # Минимальная задержка для обновления UI
                        time.sleep(0.001)
                        
                    except Exception as e:
                        dept_error_count += 1
                        total_error_count += 1
                        self.logger.error(f"Ошибка создания файла для {employee['ФИО работника']}: {e}")
                        
                        # Отправляем событие об ошибке
                        event_bus.emit_simple(
                            EventType.ERROR_OCCURRED,
                            {"error": str(e), "employee": employee},
                            "EmployeeFileCreator"
                        )
                        
                        progress.processed_files += 1
                        if file_progress_callback:
                            file_progress_callback(emp_idx + 1, len(dept_employees), f"Ошибка: {employee['ФИО работника']}")
                        self._emit_progress_update(progress, progress_callback)
                
                # Логируем результаты по отделу
                if dept_success_count > 0 or dept_error_count > 0:
                    operation_log.add_entry("INFO", f"Отдел {dept_name}: создано {dept_success_count}")
                
                # Завершаем обработку отдела
                progress.processed_blocks = dept_idx + 1
                
                if department_progress_callback:
                    department_progress_callback(dept_idx + 1, total_departments, dept_name)
            
            # 7. Завершение
            end_time = datetime.now()
            duration = end_time - start_time
            
            progress.current_operation = "Файлы созданы"
            progress.end_time = end_time
            self._emit_progress_update(progress, progress_callback)
            
            # Получаем отчет о производительности
            performance_report = self.excel_handler.performance_tracker.finish_batch()
            
            # Итоговая статистика
            operation_log.add_entry("INFO", f"Создание файлов завершено")
            operation_log.add_entry("INFO", f"Успешно создано: {total_success_count} файлов")
            
            if total_error_count > 0:
                operation_log.add_entry("WARNING", f"Ошибок при создании: {total_error_count} файлов")
            
            # Добавляем статистику производительности
            operation_log.add_entry("INFO", f"Время выполнения: {duration}")
            operation_log.add_entry("INFO", f"Среднее время на файл: {performance_report.average_duration_per_file:.2f}с")
            
            # Выводим подробный отчет о производительности в консоль
            print("\n" + performance_report.format_report())
            
            # Очищаем кэш для освобождения памяти
            self.excel_handler.clear_cache()
            
            operation_log.add_entry("INFO", f"Время выполнения: {duration.total_seconds():.1f} сек")
            
            # Завершаем транзакцию
            if self.transaction_manager.is_active:
                if total_error_count == 0:
                    # Если нет ошибок, подтверждаем транзакцию
                    if self.transaction_manager.commit_transaction():
                        operation_log.add_entry("INFO", "Транзакция подтверждена успешно")
                    else:
                        operation_log.add_entry("WARNING", "Ошибка подтверждения транзакции")
                else:
                    # Если есть ошибки, откатываем транзакцию
                    if self.transaction_manager.rollback_transaction():
                        operation_log.add_entry("INFO", "Транзакция откачена из-за ошибок")
                    else:
                        operation_log.add_entry("WARNING", "Ошибка отката транзакции")
            
            operation_log.finish(ProcessingStatus.SUCCESS)
            
            self.logger.info(f"Создание файлов завершено: {total_success_count} создано, {total_error_count} ошибок")
            
            # Отправляем событие о завершении операции
            event_bus.emit_simple(
                EventType.OPERATION_COMPLETE,
                {
                    "success_count": total_success_count,
                    "error_count": total_error_count,
                    "duration": duration.total_seconds()
                },
                "EmployeeFileCreator"
            )
            
            return operation_log
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            
            error_msg = f"Критическая ошибка: {e}"
            operation_log.add_entry("ERROR", error_msg)
            self.logger.error(error_msg, exc_info=True)
            
            # Откатываем транзакцию при критической ошибке
            if self.transaction_manager.is_active:
                if self.transaction_manager.rollback_transaction():
                    operation_log.add_entry("INFO", "Транзакция откачена из-за критической ошибки")
                else:
                    operation_log.add_entry("WARNING", "Ошибка отката транзакции при критической ошибке")
            
            operation_log.finish(ProcessingStatus.ERROR)
            
            # Отправляем событие об ошибке
            event_bus.emit_simple(
                EventType.ERROR_OCCURRED,
                {"error": error_msg, "critical": True},
                "EmployeeFileCreator"
            )
            
            return operation_log
    
    def _group_employees_by_department(self, employees: List[Dict]) -> Dict[str, List[Dict]]:
        """Группирует сотрудников по отделам"""
        departments = {}
        for employee in employees:
            dept = employee.get('Подразделение 1', '')
            if dept:
                if dept not in departments:
                    departments[dept] = []
                departments[dept].append(employee)
        return departments
    
    def _emit_progress_update(self, progress: ProcessingProgress, callback: Optional[Callable]):
        """Отправляет событие обновления прогресса"""
        if callback:
            callback(progress)
        
        # Также отправляем через шину событий
        event_bus.emit_simple(
            EventType.PROGRESS_UPDATE,
            {
                "current_operation": progress.current_operation,
                "processed_files": progress.processed_files,
                "total_files": progress.total_files,
                "processed_blocks": progress.processed_blocks,
                "total_blocks": progress.total_blocks,
                "file_progress_percent": progress.file_progress_percent,
                "block_progress_percent": progress.block_progress_percent
            },
            "EmployeeFileCreator"
        )
    
    def _should_stop(self) -> bool:
        """Проверяет, нужно ли остановить операцию"""
        # Пока простая заглушка, можно расширить через события
        return False 