#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль управления структурой папок
"""

import logging
import re
from pathlib import Path
from typing import List, Dict

from config import Config
from core.events import event_bus, EventType


class DirectoryManager:
    """Класс для управления структурой папок"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def create_department_structure(self, output_dir: str, employees: List[Dict[str, str]]) -> Dict[str, str]:
        """
        Создает структуру папок по подразделениям или использует существующие
        
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
                if emp.get('Подразделение 1'):
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
                    
                    # Отправляем событие о создании папки
                    event_bus.emit_simple(
                        EventType.DIRECTORY_CREATED,
                        {"directory_path": str(dept_path), "department_name": dept},
                        "DirectoryManager"
                    )
                else:
                    self.logger.info(f"Используется существующая папка: {clean_dept_name}")
                
                departments[dept] = str(dept_path)
            
            self.logger.info(f"Подготовлено отделов: {len(departments)}")
            return departments
            
        except Exception as e:
            self.logger.error(f"Ошибка создания структуры папок: {e}")
            raise
    
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
    
    def _clean_directory_name(self, name: str) -> str:
        """
        Очищает имя папки от недопустимых символов
        
        Args:
            name: исходное имя папки
            
        Returns:
            str: очищенное имя папки
        """
        if not name:
            return "unnamed"
        
        # Заменяем недопустимые символы
        invalid_chars = r'[<>:"/\\|?*]'
        clean_name = re.sub(invalid_chars, '_', name)
        clean_name = clean_name.strip('. ')
        
        # Ограничиваем длину
        if len(clean_name) > 100:
            clean_name = clean_name[:100]
        
        return clean_name or "unnamed"
    
    def _scan_department_files(self, dept_path: Path) -> List[str]:
        """
        Сканирует файлы в папке подразделения
        
        Args:
            dept_path: путь к папке подразделения
            
        Returns:
            List[str]: список путей к файлам
        """
        files = []
        try:
            for file_path in dept_path.iterdir():
                if file_path.is_file() and file_path.suffix.lower() == '.xlsx':
                    # Исключаем отчеты (файлы начинающиеся с '!')
                    if not file_path.name.startswith('!'):
                        files.append(str(file_path))
        except Exception as e:
            self.logger.error(f"Ошибка сканирования папки {dept_path}: {e}")
        
        return files
    
    def ensure_directory_exists(self, directory_path: Path) -> None:
        """
        Обеспечивает существование папки, создает если не существует
        
        Args:
            directory_path: путь к папке
        """
        try:
            if not directory_path.exists():
                directory_path.mkdir(parents=True, exist_ok=True)
                self.logger.debug(f"Создана папка: {directory_path}")
                
                # Отправляем событие о создании папки
                event_bus.emit_simple(
                    EventType.DIRECTORY_CREATED,
                    {"directory_path": str(directory_path)},
                    "DirectoryManager"
                )
        except Exception as e:
            self.logger.error(f"Ошибка создания папки {directory_path}: {e}")
            raise 