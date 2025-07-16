#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Утилитарный модуль для общих функций
"""

import re
import logging
from pathlib import Path
from typing import Optional, Union, Any
from datetime import datetime, date


class FileUtils:
    """Утилиты для работы с файлами"""
    
    @staticmethod
    def clean_filename(name: str, max_length: int = 100) -> str:
        """
        Очищает имя файла/папки от недопустимых символов
        
        Args:
            name: исходное имя
            max_length: максимальная длина имени
            
        Returns:
            str: очищенное имя
        """
        if not name:
            return "unnamed"
        
        # Заменяем недопустимые символы для имен файлов/папок
        invalid_chars = r'[<>:"/\\|?*]'
        clean_name = re.sub(invalid_chars, '_', name)
        
        # Убираем лишние пробелы и точки в конце
        clean_name = clean_name.strip('. ')
        
        # Ограничиваем длину
        if len(clean_name) > max_length:
            clean_name = clean_name[:max_length]
        
        return clean_name or "unnamed"
    
    @staticmethod
    def ensure_directory(path: Union[str, Path]) -> Path:
        """
        Создает папку если она не существует
        
        Args:
            path: путь к папке
            
        Returns:
            Path: объект Path для папки
        """
        path_obj = Path(path)
        path_obj.mkdir(parents=True, exist_ok=True)
        return path_obj
    
    @staticmethod
    def safe_copy_file(source: Union[str, Path], destination: Union[str, Path]) -> bool:
        """
        Безопасно копирует файл
        
        Args:
            source: исходный файл
            destination: целевой файл
            
        Returns:
            bool: True если копирование успешно
        """
        try:
            import shutil
            
            source_path = Path(source)
            dest_path = Path(destination)
            
            # Создаем папку назначения если нужно
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Копируем файл
            shutil.copy2(source_path, dest_path)
            return True
            
        except Exception:
            return False
    
    @staticmethod
    def get_file_size_mb(file_path: Union[str, Path]) -> float:
        """
        Получает размер файла в мегабайтах
        
        Args:
            file_path: путь к файлу
            
        Returns:
            float: размер в МБ
        """
        try:
            file_path_obj = Path(file_path)
            if file_path_obj.exists():
                return file_path_obj.stat().st_size / (1024 * 1024)
            return 0.0
        except Exception:
            return 0.0


class DataUtils:
    """Утилиты для работы с данными"""
    
    @staticmethod
    def convert_to_number(value: Any) -> Union[int, float, str]:
        """
        Преобразует значение к числовому типу, если возможно
        
        Args:
            value: значение для преобразования
            
        Returns:
            Union[int, float, str]: преобразованное значение
        """
        if value is None or value == '':
            return ''
        
        # Если это уже число, возвращаем как есть
        if isinstance(value, (int, float)):
            return value
        
        # Преобразуем в строку для дальнейшей обработки
        str_value = str(value).strip()
        
        if not str_value:
            return ''
        
        # Попытка преобразовать в число
        try:
            # Сначала пробуем целое число
            if '.' not in str_value and ',' not in str_value:
                # Проверяем, что это только цифры (возможно с минусом)
                if str_value.lstrip('-').isdigit():
                    return int(str_value)
            
            # Попытка преобразовать в float
            # Заменяем запятую на точку для российского формата
            float_value = str_value.replace(',', '.')
            if DataUtils._is_float(float_value):
                return float(float_value)
        except (ValueError, TypeError):
            pass
        
        # Если не удалось преобразовать в число, возвращаем как строку
        return str_value
    
    @staticmethod
    def _is_float(value: str) -> bool:
        """Проверяет, является ли строка числом с плавающей точкой"""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def safe_get_dict_value(data: dict, key: str, default: Any = '') -> Any:
        """
        Безопасно получает значение из словаря
        
        Args:
            data: словарь
            key: ключ
            default: значение по умолчанию
            
        Returns:
            Any: значение или default
        """
        if not isinstance(data, dict):
            return default
        
        value = data.get(key, default)
        return value if value is not None else default
    
    @staticmethod
    def parse_date(date_value: Any) -> Optional[date]:
        """
        Парсит дату из различных форматов
        
        Args:
            date_value: значение даты
            
        Returns:
            Optional[date]: дата или None
        """
        if not date_value:
            return None
        
        if isinstance(date_value, date):
            return date_value
        if isinstance(date_value, datetime):
            return date_value.date()
        
        date_str = str(date_value).strip()
        if not date_str:
            return None
        
        formats = ["%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"]
        
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        
        return None


class ValidationUtils:
    """Утилиты для валидации данных"""
    
    @staticmethod
    def validate_file_exists(file_path: Union[str, Path]) -> bool:
        """
        Проверяет существование файла
        
        Args:
            file_path: путь к файлу
            
        Returns:
            bool: True если файл существует
        """
        try:
            return Path(file_path).exists()
        except Exception:
            return False
    
    @staticmethod
    def validate_file_size(file_path: Union[str, Path], max_size_mb: float = 50) -> bool:
        """
        Проверяет размер файла
        
        Args:
            file_path: путь к файлу
            max_size_mb: максимальный размер в МБ
            
        Returns:
            bool: True если размер допустимый
        """
        try:
            file_size_mb = FileUtils.get_file_size_mb(file_path)
            return file_size_mb <= max_size_mb
        except Exception:
            return False
    
    @staticmethod
    def validate_tab_number(tab_number: str) -> bool:
        """
        Проверяет формат табельного номера
        
        Args:
            tab_number: табельный номер
            
        Returns:
            bool: True если формат правильный
        """
        if not tab_number:
            return False
        
        tab_number_str = str(tab_number).strip()
        return tab_number_str.isdigit() and len(tab_number_str) > 0
    
    @staticmethod
    def validate_string_length(value: str, max_length: int = 255) -> bool:
        """
        Проверяет длину строки
        
        Args:
            value: строка
            max_length: максимальная длина
            
        Returns:
            bool: True если длина допустимая
        """
        if not value:
            return True
        
        return len(str(value)) <= max_length


class LoggingUtils:
    """Утилиты для логирования"""
    
    @staticmethod
    def log_operation_start(logger: logging.Logger, operation_name: str, **kwargs):
        """
        Логирует начало операции
        
        Args:
            logger: логгер
            operation_name: название операции
            **kwargs: дополнительные параметры
        """
        params = ", ".join(f"{k}={v}" for k, v in kwargs.items())
        logger.info(f"Начинаем операцию: {operation_name}" + (f" ({params})" if params else ""))
    
    @staticmethod
    def log_operation_end(logger: logging.Logger, operation_name: str, success: bool = True, **kwargs):
        """
        Логирует завершение операции
        
        Args:
            logger: логгер
            operation_name: название операции
            success: успех операции
            **kwargs: дополнительные параметры
        """
        status = "успешно" if success else "с ошибкой"
        params = ", ".join(f"{k}={v}" for k, v in kwargs.items())
        logger.info(f"Операция {operation_name} завершена {status}" + (f" ({params})" if params else ""))
    
    @staticmethod
    def log_processing_progress(logger: logging.Logger, current: int, total: int, item_name: str = "элемент"):
        """
        Логирует прогресс обработки
        
        Args:
            logger: логгер
            current: текущее количество
            total: общее количество
            item_name: название элемента
        """
        percentage = (current / total * 100) if total > 0 else 0
        logger.debug(f"Обработано {current}/{total} {item_name} ({percentage:.1f}%)")


class ExcelUtils:
    """Утилиты для работы с Excel"""
    
    @staticmethod
    def safe_get_cell_value(worksheet, cell_address: str) -> Any:
        """
        Безопасно получает значение ячейки
        
        Args:
            worksheet: лист Excel
            cell_address: адрес ячейки
            
        Returns:
            Any: значение ячейки или None
        """
        try:
            return worksheet[cell_address].value
        except Exception:
            return None
    
    @staticmethod
    def safe_set_cell_value(worksheet, cell_address: str, value: Any) -> bool:
        """
        Безопасно устанавливает значение ячейки
        
        Args:
            worksheet: лист Excel
            cell_address: адрес ячейки
            value: значение
            
        Returns:
            bool: True если установка успешна
        """
        try:
            # Преобразуем значение к правильному типу
            converted_value = DataUtils.convert_to_number(value)
            worksheet[cell_address] = converted_value
            return True
        except Exception:
            return False
    
    @staticmethod
    def is_valid_cell_address(address: str) -> bool:
        """
        Проверяет валидность адреса ячейки
        
        Args:
            address: адрес ячейки
            
        Returns:
            bool: True если адрес валидный
        """
        return bool(re.match(r'^[A-Z]+[0-9]+$', address)) 