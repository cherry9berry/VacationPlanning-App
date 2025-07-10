#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модели данных для приложения Vacation Tool
"""

from dataclasses import dataclass, field
from datetime import datetime, date
from typing import List, Optional, Dict, Any
from enum import Enum


class VacationStatus(Enum):
    """Статусы планирования отпуска"""
    OK = "Ок"
    PARTIAL = "Частично"
    EMPTY = "Не заполнено"
    ERROR = "Ошибка"


class ProcessingStatus(Enum):
    """Статусы обработки"""
    SUCCESS = "SUCCESS"
    ERROR = "ERROR"
    CANCELLED = "CANCELLED"
    IN_PROGRESS = "IN_PROGRESS"
    COMPLETED = "COMPLETED"  # Добавляем статус завершения


@dataclass
class Employee:
    """Модель сотрудника"""
    full_name: str = ""
    tab_number: str = ""
    position: str = ""  # Должность
    department1: str = ""
    department2: str = ""
    department3: str = ""
    department4: str = ""
    file_path: str = ""
    
    # Новые поля
    location: str = ""  # Локация сотрудника
    vacation_remainder: str = ""  # Остатки отпуска сотрудника за прошлый период
    hire_date: str = ""  # Дата приема сотрудника
    period_cutoff_date: str = ""  # Дата отсечки периода
    additional_vacation_nrd: str = ""  # Дополнительный отпуск за НРД
    additional_vacation_north: str = ""  # Дополнительный отпуск Северный
    
    def __post_init__(self):
        # Очистка строк от лишних пробелов
        self.full_name = self.full_name.strip()
        self.tab_number = self.tab_number.strip()
        self.position = self.position.strip()
        self.department1 = self.department1.strip()
        self.department2 = self.department2.strip()
        self.department3 = self.department3.strip()
        self.department4 = self.department4.strip()
        self.location = self.location.strip()
        self.vacation_remainder = self.vacation_remainder.strip()
        self.hire_date = self.hire_date.strip()
        self.period_cutoff_date = self.period_cutoff_date.strip()
        self.additional_vacation_nrd = self.additional_vacation_nrd.strip()
        self.additional_vacation_north = self.additional_vacation_north.strip()


@dataclass
class VacationPeriod:
    """Период отпуска"""
    start_date: date
    end_date: date
    days: int = 0
    
    def __post_init__(self):
        if self.days == 0:
            self.days = (self.end_date - self.start_date).days + 1


@dataclass
class VacationInfo:
    """Информация об отпусках сотрудника"""
    employee: Employee
    periods: List[VacationPeriod] = field(default_factory=list)
    total_days: int = 0
    status: VacationStatus = VacationStatus.EMPTY
    periods_count: int = 0
    has_long_period: bool = False  # Есть ли период >= 14 дней
    validation_errors: List[str] = field(default_factory=list)  # Ошибки валидации
    
    def __post_init__(self):
        self.periods_count = len(self.periods)
        self.total_days = sum(period.days for period in self.periods)
        self.has_long_period = any(period.days >= 14 for period in self.periods)
        self._update_status()
    
    def _update_status(self):
        """Обновляет статус на основе данных"""
        if not self.periods:
            self.status = VacationStatus.EMPTY
        elif self.validation_errors:
            self.status = VacationStatus.ERROR
        elif self.total_days >= 28 and self.has_long_period:
            self.status = VacationStatus.OK
        else:
            self.status = VacationStatus.PARTIAL


@dataclass
class BlockReport:
    """Отчет по блоку/подразделению"""
    block_name: str
    employees: List[VacationInfo] = field(default_factory=list)
    total_employees: int = 0
    status_ok: int = 0
    status_partial: int = 0
    status_empty: int = 0
    status_error: int = 0  # Добавляем подсчет ошибок
    average_days: float = 0.0
    completion_percentage: float = 0.0
    
    def __post_init__(self):
        self.total_employees = len(self.employees)
        self._calculate_statistics()
    
    def _calculate_statistics(self):
        """Вычисляет статистику блока"""
        if not self.employees:
            return
            
        status_counts = {status: 0 for status in VacationStatus}
        total_days = 0
        
        for emp in self.employees:
            status_counts[emp.status] += 1
            total_days += emp.total_days
        
        self.status_ok = status_counts[VacationStatus.OK]
        self.status_partial = status_counts[VacationStatus.PARTIAL]
        self.status_empty = status_counts[VacationStatus.EMPTY]
        self.status_error = status_counts[VacationStatus.ERROR]
        
        self.average_days = total_days / self.total_employees if self.total_employees > 0 else 0
        self.completion_percentage = (self.status_ok / self.total_employees * 100) if self.total_employees > 0 else 0


@dataclass
class GeneralReport:
    """Общий отчет по всем блокам"""
    block_reports: List[BlockReport] = field(default_factory=list)
    total_employees: int = 0
    status_ok: int = 0
    status_partial: int = 0
    status_empty: int = 0
    status_error: int = 0
    average_days: float = 0.0
    average_periods: float = 0.0
    completion_percentage: float = 0.0
    generated_at: datetime = field(default_factory=datetime.now)
    
    def __post_init__(self):
        self._calculate_totals()
    
    def _calculate_totals(self):
        """Вычисляет общую статистику"""
        if not self.block_reports:
            return
        
        self.total_employees = sum(block.total_employees for block in self.block_reports)
        self.status_ok = sum(block.status_ok for block in self.block_reports)
        self.status_partial = sum(block.status_partial for block in self.block_reports)
        self.status_empty = sum(block.status_empty for block in self.block_reports)
        self.status_error = sum(block.status_error for block in self.block_reports)
        
        if self.total_employees > 0:
            total_days = sum(block.average_days * block.total_employees for block in self.block_reports)
            self.average_days = total_days / self.total_employees
            
            total_periods = sum(
                sum(emp.periods_count for emp in block.employees) 
                for block in self.block_reports
            )
            self.average_periods = total_periods / self.total_employees
            
            self.completion_percentage = self.status_ok / self.total_employees * 100


@dataclass
class ValidationResult:
    """Результат валидации"""
    is_valid: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    employee_count: int = 0
    unique_tab_numbers: int = 0
    processing_time: float = 0.0
    
    def add_error(self, message: str):
        """Добавляет ошибку"""
        self.errors.append(message)
        self.is_valid = False
    
    def add_warning(self, message: str):
        """Добавляет предупреждение"""
        self.warnings.append(message)


@dataclass
class ProcessingProgress:
    """Прогресс обработки"""
    current_operation: str = ""
    current_file: str = ""
    processed_files: int = 0
    total_files: int = 0
    current_block: str = ""
    processed_blocks: int = 0
    total_blocks: int = 0
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    elapsed_time: float = 0.0
    estimated_time: float = 0.0
    speed: float = 0.0  # файлов в секунду
    status: ProcessingStatus = ProcessingStatus.IN_PROGRESS
    error_message: str = ""
    
    @property
    def file_progress_percent(self) -> float:
        """Процент выполнения по файлам"""
        if self.total_files == 0:
            return 0.0
        return (self.processed_files / self.total_files) * 100
    
    @property
    def block_progress_percent(self) -> float:
        """Процент выполнения по блокам"""
        if self.total_blocks == 0:
            return 0.0
        return (self.processed_blocks / self.total_blocks) * 100


@dataclass
class LogEntry:
    """Запись в логе операции"""
    timestamp: datetime
    level: str  # INFO, WARN, ERROR
    operation: str
    message: str
    details: str = ""


@dataclass
class OperationLog:
    """Лог операции"""
    operation_name: str
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    entries: List[LogEntry] = field(default_factory=list)
    status: ProcessingStatus = ProcessingStatus.IN_PROGRESS
    
    def add_entry(self, level: str, message: str, details: str = ""):
        """Добавляет запись в лог"""
        entry = LogEntry(
            timestamp=datetime.now(),
            level=level,
            operation=self.operation_name,
            message=message,
            details=details
        )
        self.entries.append(entry)
    
    def finish(self, status: ProcessingStatus):
        """Завершает операцию"""
        self.end_time = datetime.now()
        self.status = status
    
    @property
    def duration(self) -> float:
        """Длительность операции в секундах"""
        if self.end_time is None:
            return (datetime.now() - self.start_time).total_seconds()
        return (self.end_time - self.start_time).total_seconds()