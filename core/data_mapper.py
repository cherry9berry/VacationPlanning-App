#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для динамического маппинга данных
"""

import logging
from datetime import datetime, date
from typing import Dict, Any, List, Optional
from models import VacationInfo, VacationPeriod, VacationStatus


class DataMapper:
    """Класс для динамического маппинга данных между различными форматами"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def map_vacation_info_to_rules(self, vacation_info: VacationInfo, index: int, prefix: str = '') -> Dict[str, Any]:
        """
        Маппит данные сотрудника для Report листа
        
        Args:
            vacation_info: Информация об отпуске сотрудника
            index: Индекс строки
            prefix: Префикс для полей
            
        Returns:
            Словарь с данными для заполнения
        """
        employee = vacation_info.employee
        
        return {
            f'{prefix}employee_name': employee.get('ФИО работника', ''),
            f'{prefix}tab_number': employee.get('Табельный номер', ''),
            f'{prefix}position': employee.get('Должность', ''),
            f'{prefix}department1': employee.get('Подразделение 1', ''),
            f'{prefix}department2': employee.get('Подразделение 2', ''),
            f'{prefix}department3': employee.get('Подразделение 3', ''),
            f'{prefix}department4': employee.get('Подразделение 4', ''),
            f'{prefix}status': vacation_info.get_status_text(),
            f'{prefix}total_days': sum(p.days for p in vacation_info.periods) if vacation_info.periods else '',
            f'{prefix}periods_count': len(vacation_info.periods) if vacation_info.periods else '',
            f'{prefix}row_number': index + 1
        }
    
    def map_period_data_to_rules(self, period_data: Dict[str, Any], index: int, prefix: str = '') -> Dict[str, Any]:
        """
        Маппит данные периода для Print листа
        
        Args:
            period_data: Данные периода отпуска
            index: Индекс строки
            prefix: Префикс для полей
            
        Returns:
            Словарь с данными для заполнения
        """
        employee = period_data.get('employee', {})
        start_date = period_data.get('start_date')
        end_date = period_data.get('end_date')
        days = period_data.get('days', 0)
        
        return {
            f'{prefix}employee_name': employee.get('ФИО работника', ''),
            f'{prefix}tab_number': employee.get('Табельный номер', ''),
            f'{prefix}position': employee.get('Должность', ''),
            f'{prefix}start_date': self._format_date(start_date) if start_date else '',
            f'{prefix}end_date': self._format_date(end_date) if end_date else '',
            f'{prefix}duration': str(days) if days else '',
            f'{prefix}signature': '',
            f'{prefix}acknowledgment_date': '',
            f'{prefix}notes': '',
            f'{prefix}row_number': index + 1
        }
    
    def map_block_data_to_rules(self, block_data: Dict[str, Any], index: int, prefix: str = '') -> Dict[str, Any]:
        """
        Маппит данные блока для общего отчета
        
        Args:
            block_data: Данные блока
            index: Индекс строки
            prefix: Префикс для полей
            
        Returns:
            Словарь с данными для заполнения
        """
        return {
            f'{prefix}row_number2': index + 1,  # Исправлено: row_number2 вместо row_number
            f'{prefix}report_department1': block_data.get('block_name', ''),  # Исправлено: report_department1 вместо department1
            f'{prefix}percentage': f"{block_data.get('percentage', 0)}%",
            f'{prefix}employees_count': block_data.get('total_employees', 0),
            f'{prefix}correct_filled': block_data.get('completed_employees', 0),
            f'{prefix}incorrect_filled': block_data.get('employees_incorrect', 0),
            f'{prefix}not_filled': block_data.get('employees_not_filled', 0),
            f'{prefix}update_date': block_data.get('update_date', '')
        }
    
    def map_report_header_data(self, block_name: str, vacation_infos: List[VacationInfo]) -> Dict[str, Any]:
        """
        Маппит данные заголовка отчета по блоку
        
        Args:
            block_name: Название блока
            vacation_infos: Список информации об отпусках
            
        Returns:
            Словарь с данными заголовка
        """
        total_employees = len(vacation_infos)
        employees_filled = sum(1 for info in vacation_infos if info.status != VacationStatus.NOT_FILLED)
        employees_correct = sum(1 for info in vacation_infos if info.status == VacationStatus.FILLED_CORRECT)
        
        return {
            'block_name': block_name,
            'update_date': self._format_datetime(datetime.now()),
            'total_employees': str(total_employees),
            'employees_filled': str(employees_filled),
            'employees_correct': str(employees_correct)
        }
    
    def map_general_header_data(self, block_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Маппит данные заголовка общего отчета
        
        Args:
            block_data: Список данных блоков
            
        Returns:
            Словарь с данными заголовка
        """
        return {
            'update_date2': self._format_datetime(datetime.now()),
            'blocks_count': len(block_data)
        }
    
    def _format_date(self, date_obj: date) -> str:
        """Форматирует дату в строку DD.MM.YYYY"""
        if not date_obj:
            return ''
        return date_obj.strftime('%d.%m.%Y')
    
    def _format_datetime(self, datetime_obj: datetime) -> str:
        """Форматирует дату и время в строку DD.MM.YYYY HH:MM"""
        if not datetime_obj:
            return ''
        return datetime_obj.strftime('%d.%m.%Y %H:%M') 