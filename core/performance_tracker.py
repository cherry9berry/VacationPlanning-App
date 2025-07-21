#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для отслеживания производительности создания файлов
"""

import time
import logging
from typing import Dict, List, Optional
from dataclasses import dataclass, field
from datetime import datetime, timedelta


@dataclass
class FilePerformanceStats:
    """Статистика производительности для одного файла"""
    filename: str
    start_time: float
    end_time: Optional[float] = None
    duration: Optional[float] = None
    success: bool = False
    error_message: Optional[str] = None
    
    def finish(self, success: bool = True, error_message: Optional[str] = None):
        """Завершает отслеживание файла"""
        self.end_time = time.time()
        self.duration = self.end_time - self.start_time
        self.success = success
        self.error_message = error_message


@dataclass
class PerformanceReport:
    """Отчет о производительности"""
    total_files: int
    successful_files: int
    failed_files: int
    skipped_files: int
    total_duration: float
    average_duration_per_file: float
    fastest_file: Optional[FilePerformanceStats] = None
    slowest_file: Optional[FilePerformanceStats] = None
    files_stats: List[FilePerformanceStats] = field(default_factory=list)
    
    def format_report(self) -> str:
        """Форматирует отчет в читаемый вид"""
        report = []
        report.append("=" * 60)
        report.append("СТАТИСТИКА ПРОИЗВОДИТЕЛЬНОСТИ")
        report.append("=" * 60)
        
        # Общая статистика
        report.append(f"Всего файлов: {self.total_files}")
        report.append(f"Успешно создано: {self.successful_files}")
        report.append(f"Ошибок: {self.failed_files}")
        report.append(f"Пропущено: {self.skipped_files}")
        report.append("")
        
        # Временная статистика
        total_time_str = str(timedelta(seconds=int(self.total_duration)))
        avg_time_str = f"{self.average_duration_per_file:.2f}с"
        
        report.append(f"Общее время: {total_time_str}")
        report.append(f"Среднее время на файл: {avg_time_str}")
        
        if self.fastest_file:
            report.append(f"Самый быстрый файл: {self.fastest_file.filename} ({self.fastest_file.duration:.2f}с)")
        
        if self.slowest_file:
            report.append(f"Самый медленный файл: {self.slowest_file.filename} ({self.slowest_file.duration:.2f}с)")
        
        report.append("")
        
        # Производительность
        if self.total_duration > 0:
            files_per_second = self.successful_files / self.total_duration
            report.append(f"Скорость: {files_per_second:.2f} файлов/сек")
        

        
        return "\n".join(report)


class PerformanceTracker:
    """Класс для отслеживания производительности создания файлов"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.files_stats: List[FilePerformanceStats] = []
        self.start_time: Optional[float] = None
        self.end_time: Optional[float] = None
        self.skipped_count: int = 0
        
    def start_batch(self):
        """Начинает отслеживание пакета файлов"""
        self.start_time = time.time()
        self.files_stats.clear()
        self.skipped_count = 0
        self.logger.info("Начато отслеживание производительности")
        
    def start_file(self, filename: str) -> FilePerformanceStats:
        """Начинает отслеживание одного файла"""
        stats = FilePerformanceStats(
            filename=filename,
            start_time=time.time()
        )
        self.files_stats.append(stats)
        return stats
    
    def skip_file(self, filename: str):
        """Отмечает файл как пропущенный"""
        self.skipped_count += 1
        self.logger.debug(f"Файл пропущен: {filename}")
    
    def finish_batch(self) -> PerformanceReport:
        """Завершает отслеживание и создает отчет"""
        self.end_time = time.time()
        
        if self.start_time is None:
            raise ValueError("Отслеживание не было начато")
        
        total_duration = self.end_time - self.start_time
        
        # Фильтруем только успешные файлы для расчета средней скорости
        successful_files = [f for f in self.files_stats if f.success and f.duration is not None]
        failed_files = [f for f in self.files_stats if not f.success]
        
        # Рассчитываем статистику
        successful_count = len(successful_files)
        failed_count = len(failed_files)
        total_files = successful_count + failed_count + self.skipped_count
        
        if successful_files:
            average_duration = sum(f.duration for f in successful_files if f.duration is not None) / len(successful_files)
            fastest_file = min(successful_files, key=lambda f: f.duration if f.duration is not None else float('inf'))
            slowest_file = max(successful_files, key=lambda f: f.duration if f.duration is not None else 0.0)
        else:
            average_duration = 0
            fastest_file = None
            slowest_file = None
        
        report = PerformanceReport(
            total_files=total_files,
            successful_files=successful_count,
            failed_files=failed_count,
            skipped_files=self.skipped_count,
            total_duration=total_duration,
            average_duration_per_file=average_duration,
            fastest_file=fastest_file,
            slowest_file=slowest_file,
            files_stats=self.files_stats.copy()
        )
        
        self.logger.info(f"Отслеживание завершено. Создано {successful_count} файлов за {total_duration:.2f}с")
        return report 