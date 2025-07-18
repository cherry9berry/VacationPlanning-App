#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отладка правил заполнения
"""

import sys
import os
sys.path.append('.')

from core.excel_handler import ExcelHandler
from config import Config

def debug_rules():
    """Отладка правил заполнения"""
    
    # Загружаем конфигурацию
    config = Config()
    config.load()
    
    # Создаем ExcelHandler
    excel_handler = ExcelHandler(config)
    
    # Получаем правила
    template_path = config.employee_template
    rules = excel_handler._get_cached_rules(template_path)
    
    print("Правила заполнения:")
    print("==================")
    
    for rule_type, rule_items in rules.items():
        print(f"\nТип правила: {rule_type}")
        for cell_address, field_name in rule_items.items():
            print(f"  {cell_address} -> {field_name}")
    
    # Тестовые данные
    test_employee = {
        'Табельный номер': '12345',
        'ФИО работника': 'Тестовый Сотрудник',
        'Должность': 'Тестовая Должность',
        'Подразделение 1': 'Тестовый Отдел',
        'Подразделение 2': '',
        'Подразделение 3': '',
        'Подразделение 4': ''
    }
    
    print("\nДанные сотрудника:")
    print("==================")
    for key, value in test_employee.items():
        print(f"  {key}: '{value}'")
    
    print("\nСоответствие правил и данных:")
    print("=============================")
    
    for rule_type, rule_items in rules.items():
        if rule_type == 'value':
            for cell_address, field_name in rule_items.items():
                value = test_employee.get(field_name, 'НЕ НАЙДЕНО')
                print(f"  {cell_address} -> {field_name} = '{value}'")

if __name__ == "__main__":
    debug_rules() 