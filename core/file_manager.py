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
from core.directory_manager import DirectoryManager


class FileManager:
    """Класс для управления файлами и структурой папок"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.excel_handler = ExcelHandler(config)
        self.directory_manager = DirectoryManager(config)
        


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
            self.directory_manager.ensure_directory_exists(output_dir)
            
            self.logger.info(f"Создана выходная папка: {output_dir}")
            return str(output_dir)
            
        except Exception as e:
            self.logger.error(f"Ошибка создания выходной папки: {e}")
            raise
    

    

    

    
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
                employee_files = self.directory_manager._scan_department_files(entry)
                
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
    
