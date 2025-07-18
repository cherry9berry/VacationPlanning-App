#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Система событий для замены прямых ссылок на GUI
"""

from typing import Dict, List, Callable, Any
from dataclasses import dataclass
from enum import Enum
import logging


class EventType(Enum):
    """Типы событий в системе"""
    FILE_CREATED = "file_created"
    DIRECTORY_CREATED = "directory_created"
    PROGRESS_UPDATE = "progress_update"
    DEPARTMENT_PROGRESS = "department_progress"
    FILE_PROGRESS = "file_progress"
    OPERATION_COMPLETE = "operation_complete"
    ERROR_OCCURRED = "error_occurred"


@dataclass
class Event:
    """Базовый класс события"""
    event_type: EventType
    data: Dict[str, Any]
    source: str = "unknown"


class EventBus:
    """Шина событий для развязки компонентов"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._listeners: Dict[EventType, List[Callable]] = {}
    
    def subscribe(self, event_type: EventType, callback: Callable[[Event], None]):
        """Подписка на событие"""
        if event_type not in self._listeners:
            self._listeners[event_type] = []
        self._listeners[event_type].append(callback)
        self.logger.debug(f"Подписка на событие {event_type.value}")
    
    def unsubscribe(self, event_type: EventType, callback: Callable[[Event], None]):
        """Отписка от события"""
        if event_type in self._listeners:
            try:
                self._listeners[event_type].remove(callback)
                self.logger.debug(f"Отписка от события {event_type.value}")
            except ValueError:
                pass
    
    def emit(self, event: Event):
        """Отправка события"""
        if event.event_type in self._listeners:
            for callback in self._listeners[event.event_type]:
                try:
                    callback(event)
                except Exception as e:
                    self.logger.error(f"Ошибка в обработчике события {event.event_type.value}: {e}")
    
    def emit_simple(self, event_type: EventType, data: Dict[str, Any], source: str = "unknown"):
        """Упрощенная отправка события"""
        event = Event(event_type, data, source)
        self.emit(event)


# Глобальная шина событий
event_bus = EventBus() 