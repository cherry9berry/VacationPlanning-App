#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отладка rules в шаблонах отчетов
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import Config
from core.excel_handler import ExcelHandler

def debug_report_rules():
    """Анализирует rules в шаблонах отчетов"""
    config = Config()
    excel_handler = ExcelHandler(config)
    
    print("Анализ rules в шаблонах отчетов")
    print("=" * 50)
    
    # Анализ шаблона отчета по блоку
    print("\n1. ШАБЛОН ОТЧЕТА ПО БЛОКУ:")
    print("-" * 30)
    block_template = config.block_report_template
    print(f"Путь: {block_template}")
    
    if os.path.exists(block_template):
        try:
            block_rules = excel_handler._load_filling_rules(block_template)
            print("Правила заполнения:")
            
            for rule_type, rules in block_rules.items():
                print(f"\nТип правила: {rule_type}")
                for cell_address, field_name in rules.items():
                    print(f"  {cell_address} -> {field_name}")
        except Exception as e:
            print(f"Ошибка загрузки rules: {e}")
    else:
        print("Файл не найден!")
    
    # Анализ шаблона общего отчета
    print("\n\n2. ШАБЛОН ОБЩЕГО ОТЧЕТА:")
    print("-" * 30)
    general_template = config.general_report_template
    print(f"Путь: {general_template}")
    
    if os.path.exists(general_template):
        try:
            general_rules = excel_handler._load_filling_rules(general_template)
            print("Правила заполнения:")
            
            for rule_type, rules in general_rules.items():
                print(f"\nТип правила: {rule_type}")
                for cell_address, field_name in rules.items():
                    print(f"  {cell_address} -> {field_name}")
        except Exception as e:
            print(f"Ошибка загрузки rules: {e}")
    else:
        print("Файл не найден!")

if __name__ == "__main__":
    debug_report_rules() 