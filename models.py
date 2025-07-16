#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модели данных для приложения
"""

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from typing import List, Optional, Dict


class VacationStatus(Enum):
    """Статусы валидации отпусков (обновленные под новый шаблон)"""
    NOT_FILLED = "Форма не заполнена"
    FILLED_INCORRECT = "Форма заполнена некорректно"
    FILLED_CORRECT = "Форма заполнена корректно"


class ProcessingStatus(Enum):
    """Статусы выполнения операций"""
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    ERROR = "error"
    CANCELLED = "cancelled"


@dataclass
class Employee:
    """Модель сотрудника"""
    full_name: str = ""
    tab_number: str = ""
    position: str = ""
    department1: str = ""
    department2: str = ""
    department3: str = ""
    department4: str = ""
    
    # Дополнительные поля
    location: str = ""
    vacation_remainder: str = ""
    hire_date: str = ""
    period_cutoff_date: str = ""
    additional_vacation_nrd: str = ""
    additional_vacation_north: str = ""
    
    # Служебные поля
    file_path: Optional[str] = None  # Путь к файлу сотрудника для чтения статуса


@dataclass
class VacationPeriod:
    """Период отпуска"""
    start_date: date
    end_date: date
    days: int = 0
    
    def __post_init__(self):
        """Автоматически вычисляет количество дней если не указано"""
        if self.days == 0:
            self.days = (self.end_date - self.start_date).days + 1


@dataclass
class VacationInfo:
    """Информация об отпусках сотрудника"""
    employee: Dict[str, str]
    periods: List[VacationPeriod] = field(default_factory=list)
    status: VacationStatus = VacationStatus.NOT_FILLED
    validation_errors: List[str] = field(default_factory=list)
    
    @property
    def total_days(self) -> int:
        """Общее количество дней отпуска"""
        return sum(period.days for period in self.periods)
    
    @property
    def periods_count(self) -> int:
        """Количество периодов отпуска"""
        return len(self.periods)
    
    @property
    def has_long_period(self) -> bool:
        """Есть ли период 14+ дней"""
        return any(period.days >= 14 for period in self.periods)
    
    def get_status_text(self) -> str:
        """Возвращает текстовое представление статуса"""
        return self.status.value


@dataclass
class ValidationResult:
    """Результат валидации"""
    is_valid: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    employee_count: int = 0
    unique_tab_numbers: int = 0
    processing_time: float = 0.0
    
    def add_error(self, message: str) -> None:
        """Добавляет ошибку"""
        self.errors.append(message)
        self.is_valid = False
    
    def add_warning(self, message: str) -> None:
        """Добавляет предупреждение"""
        self.warnings.append(message)
    
    def get_summary(self) -> str:
        """Возвращает краткое описание результата"""
        if self.is_valid:
            return f"✓ Валидация пройдена. Найдено сотрудников: {self.employee_count}"
        else:
            return f"✗ Найдено ошибок: {len(self.errors)}"


@dataclass
class ProcessingResult:
    """Результат обработки файлов"""
    success: bool = True
    created_files: int = 0
    skipped_files: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    processing_time: float = 0.0
    
    def add_error(self, message: str) -> None:
        """Добавляет ошибку"""
        self.errors.append(message)
        self.success = False
    
    def add_warning(self, message: str) -> None:
        """Добавляет предупреждение"""
        self.warnings.append(message)


@dataclass
class ProcessingProgress:
    """Прогресс выполнения операции"""
    current_operation: str = ""
    current_file: str = ""
    current_block: str = ""
    processed_files: int = 0
    total_files: int = 0
    processed_blocks: int = 0
    total_blocks: int = 0
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
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
class OperationLog:
    """Лог выполнения операции"""
    operation_name: str
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    status: ProcessingStatus = ProcessingStatus.RUNNING
    entries: List[dict] = field(default_factory=list)
    
    def add_entry(self, level: str, message: str) -> None:
        """Добавляет запись в лог"""
        entry = {
            "timestamp": datetime.now(),
            "level": level,
            "message": message
        }
        self.entries.append(entry)
    
    def finish(self, status: ProcessingStatus) -> None:
        """Завершает операцию"""
        self.status = status
        self.end_time = datetime.now()
    
    @property
    def duration(self) -> Optional[float]:
        """Длительность операции в секундах"""
        if self.end_time and self.start_time:
            return (self.end_time - self.start_time).total_seconds()
        return None


@dataclass
class BlockReport:
    """Отчет по блоку/подразделению"""
    block_name: str
    total_employees: int = 0
    completed_employees: int = 0
    remaining_employees: int = 0
    percentage: float = 0.0
    update_date: str = ""
    file_path: str = ""
    
    @property
    def completion_rate(self) -> float:
        """Процент завершения"""
        if self.total_employees == 0:
            return 0.0
        return (self.completed_employees / self.total_employees) * 100


@dataclass
class GeneralReport:
    """Общий отчет по всем блокам"""
    blocks: List[BlockReport] = field(default_factory=list)
    total_employees: int = 0
    total_completed: int = 0
    total_remaining: int = 0
    overall_percentage: float = 0.0
    creation_date: str = ""
    file_path: str = ""
    
    def calculate_totals(self) -> None:
        """Пересчитывает общие показатели"""
        self.total_employees = sum(block.total_employees for block in self.blocks)
        self.total_completed = sum(block.completed_employees for block in self.blocks)
        self.total_remaining = self.total_employees - self.total_completed
        self.overall_percentage = (self.total_completed / self.total_employees * 100) if self.total_employees > 0 else 0.0