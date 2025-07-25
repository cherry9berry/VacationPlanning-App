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


