#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Менеджер транзакций для обеспечения атомарности операций
"""

import logging
import shutil
from pathlib import Path
from typing import List, Dict, Callable, Optional, Any
from dataclasses import dataclass
from datetime import datetime

from core.events import event_bus, EventType


@dataclass
class TransactionOperation:
    """Операция в транзакции"""
    operation_type: str  # 'create_file', 'create_directory', 'delete_file'
    path: str
    metadata: Dict[str, Any]
    backup_path: Optional[str] = None
    
    def __init__(self, operation_type: str, path: str, metadata: Optional[Dict[str, Any]] = None, backup_path: Optional[str] = None):
        self.operation_type = operation_type
        self.path = path
        self.metadata = metadata if metadata is not None else {}
        self.backup_path = backup_path


class TransactionManager:
    """Менеджер транзакций для обеспечения атомарности операций"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._current_transaction: List[TransactionOperation] = []
        self._backup_dir: Optional[Path] = None
        self._transaction_active = False
    
    def begin_transaction(self, backup_dir: Optional[str] = None) -> bool:
        """
        Начинает новую транзакцию
        
        Args:
            backup_dir: директория для резервных копий (опционально)
            
        Returns:
            bool: True если транзакция начата успешно
        """
        if self._transaction_active:
            self.logger.warning("Транзакция уже активна")
            return False
        
        try:
            self._current_transaction = []
            self._transaction_active = True
            
            if backup_dir:
                self._backup_dir = Path(backup_dir)
                self._backup_dir.mkdir(parents=True, exist_ok=True)
                self.logger.info(f"Транзакция начата с резервным копированием в {backup_dir}")
            else:
                self._backup_dir = None
                self.logger.info("Транзакция начата без резервного копирования")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка начала транзакции: {e}")
            self._transaction_active = False
            return False
    
    def commit_transaction(self) -> bool:
        """
        Подтверждает транзакцию
        
        Returns:
            bool: True если транзакция подтверждена успешно
        """
        if not self._transaction_active:
            self.logger.warning("Нет активной транзакции для подтверждения")
            return False
        
        try:
            # Очищаем резервные копии если они есть
            if self._backup_dir and self._backup_dir.exists():
                shutil.rmtree(self._backup_dir)
                self.logger.info("Резервные копии удалены после подтверждения транзакции")
            
            self._current_transaction.clear()
            self._transaction_active = False
            self._backup_dir = None
            
            self.logger.info("Транзакция подтверждена успешно")
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка подтверждения транзакции: {e}")
            return False
    
    def rollback_transaction(self) -> bool:
        """
        Откатывает транзакцию
        
        Returns:
            bool: True если транзакция откачена успешно
        """
        if not self._transaction_active:
            self.logger.warning("Нет активной транзакции для отката")
            return False
        
        try:
            rollback_success = True
            
            # Откатываем операции в обратном порядке
            for operation in reversed(self._current_transaction):
                try:
                    if operation.operation_type == 'create_file':
                        rollback_success &= self._rollback_create_file(operation)
                    elif operation.operation_type == 'create_directory':
                        rollback_success &= self._rollback_create_directory(operation)
                    elif operation.operation_type == 'delete_file':
                        rollback_success &= self._rollback_delete_file(operation)
                except Exception as e:
                    self.logger.error(f"Ошибка отката операции {operation.operation_type}: {e}")
                    rollback_success = False
            
            # Очищаем состояние транзакции
            self._current_transaction.clear()
            self._transaction_active = False
            
            if self._backup_dir and self._backup_dir.exists():
                try:
                    shutil.rmtree(self._backup_dir)
                except Exception as e:
                    self.logger.warning(f"Не удалось удалить резервную директорию: {e}")
            
            self._backup_dir = None
            
            if rollback_success:
                self.logger.info("Транзакция откачена успешно")
            else:
                self.logger.warning("Транзакция откачена с ошибками")
            
            return rollback_success
            
        except Exception as e:
            self.logger.error(f"Критическая ошибка отката транзакции: {e}")
            self._transaction_active = False
            return False
    
    def add_file_creation(self, file_path: str, employee: Optional[Dict[str, Any]] = None) -> bool:
        """
        Добавляет операцию создания файла в транзакцию
        
        Args:
            file_path: путь к создаваемому файлу
            employee: данные сотрудника (для метаданных)
            
        Returns:
            bool: True если операция добавлена успешно
        """
        if not self._transaction_active:
            return True  # Если транзакция не активна, просто продолжаем
        
        try:
            path = Path(file_path)
            
            # Создаем резервную копию если файл существует
            backup_path: Optional[str] = None
            if path.exists() and self._backup_dir:
                backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{path.name}"
                backup_path = str(self._backup_dir / backup_name)
                shutil.copy2(path, backup_path)
                self.logger.debug(f"Создана резервная копия: {backup_path}")
            
            operation = TransactionOperation(
                operation_type='create_file',
                path=str(path),
                metadata={'employee': employee} if employee is not None else {},
                backup_path=backup_path
            )
            
            self._current_transaction.append(operation)
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка добавления операции создания файла: {e}")
            return False
    
    def add_directory_creation(self, dir_path: str, department_name: Optional[str] = None) -> bool:
        """
        Добавляет операцию создания директории в транзакцию
        
        Args:
            dir_path: путь к создаваемой директории
            department_name: название отдела (для метаданных)
            
        Returns:
            bool: True если операция добавлена успешно
        """
        if not self._transaction_active:
            return True  # Если транзакция не активна, просто продолжаем
        
        try:
            operation = TransactionOperation(
                operation_type='create_directory',
                path=str(Path(dir_path)),
                metadata={'department_name': department_name} if department_name is not None else {},
                backup_path=None
            )
            
            self._current_transaction.append(operation)
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка добавления операции создания директории: {e}")
            return False
    
    def _rollback_create_file(self, operation: TransactionOperation) -> bool:
        """Откатывает создание файла"""
        try:
            path = Path(operation.path)
            
            if path.exists():
                if operation.backup_path and Path(operation.backup_path).exists():
                    # Восстанавливаем из резервной копии
                    shutil.copy2(operation.backup_path, path)
                    self.logger.debug(f"Восстановлен файл из резервной копии: {operation.path}")
                else:
                    # Удаляем созданный файл
                    path.unlink()
                    self.logger.debug(f"Удален созданный файл: {operation.path}")
                
                # Отправляем событие об откате
                event_bus.emit_simple(
                    EventType.ERROR_OCCURRED,
                    {"error": "Файл откачен", "file_path": operation.path},
                    "TransactionManager"
                )
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка отката создания файла {operation.path}: {e}")
            return False
    
    def _rollback_create_directory(self, operation: TransactionOperation) -> bool:
        """Откатывает создание директории"""
        try:
            path = Path(operation.path)
            
            if path.exists() and path.is_dir():
                # Удаляем директорию только если она пустая
                try:
                    path.rmdir()
                    self.logger.debug(f"Удалена созданная директория: {operation.path}")
                except OSError:
                    # Директория не пустая - оставляем как есть
                    self.logger.warning(f"Директория не пустая, оставляем: {operation.path}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка отката создания директории {operation.path}: {e}")
            return False
    
    def _rollback_delete_file(self, operation: TransactionOperation) -> bool:
        """Откатывает удаление файла"""
        try:
            if operation.backup_path and Path(operation.backup_path).exists():
                # Восстанавливаем из резервной копии
                shutil.copy2(operation.backup_path, operation.path)
                self.logger.debug(f"Восстановлен удаленный файл: {operation.path}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка отката удаления файла {operation.path}: {e}")
            return False
    
    @property
    def is_active(self) -> bool:
        """Возвращает True если транзакция активна"""
        return self._transaction_active
    
    @property
    def operation_count(self) -> int:
        """Возвращает количество операций в текущей транзакции"""
        return len(self._current_transaction) 