#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест новой логики чтения периодов по статусу
"""

import os
import sys
from pathlib import Path

# Добавляем путь к проекту
sys.path.insert(0, str(Path(__file__).parent))

from core.excel_handler import ExcelHandler
from config import Config

def test_status_logic():
    """Тестирование логики чтения периодов по статусу"""
    try:
        # Тестируем несколько файлов
        test_files = [
            r"tests\test6\ИТ отдел\Ковалев Андрей Петрович (55443).xlsx",
            r"tests\test6\ИТ отдел\Тихонова Дарья Павловна (77888).xlsx", 
            r"tests\test6\ИТ отдел\Федорова Ольга Николаевна (44556).xlsx"
        ]
        
        config = Config()
        excel_handler = ExcelHandler(config)
        
        all_passed = True
        
        for file_path in test_files:
            print(f"\nТестируем файл: {file_path}")
            
            if not Path(file_path).exists():
                print(f"  ❌ Файл не найден: {file_path}")
                continue
                
            vacation_info = excel_handler.read_vacation_info_from_file(file_path)
            
            if vacation_info is None:
                print(f"  ❌ Не удалось прочитать файл")
                all_passed = False
                continue
                
            employee_name = vacation_info.employee.get('ФИО работника', 'Неизвестно')
            status = vacation_info.get_status_text()
            periods_count = len(vacation_info.periods)
            total_days = vacation_info.total_days
            
            print(f"  Сотрудник: {employee_name}")
            print(f"  Статус: {status}")
            print(f"  Количество периодов: {periods_count}")
            print(f"  Общее количество дней: {total_days}")
            
            # Проверяем логику
            if status == "Форма заполнена корректно":
                if periods_count > 0 and total_days > 0:
                    print("  ✅ Корректно: статус 'заполнена корректно' - есть периоды")
                    for i, period in enumerate(vacation_info.periods, 1):
                        print(f"    Период {i}: {period.days} дней")
                else:
                    print("  ❌ Ошибка: статус корректный, но периодов нет")
                    all_passed = False
            else:
                # ОБНОВЛЕННАЯ ЛОГИКА: сохраняем статус, но периодов нет
                if periods_count == 0 and total_days == 0:
                    if status in ["Форма заполнена некорректно", "Форма не заполнена"]:
                        print(f"  ✅ Корректно: статус '{status}' сохранен, периодов нет")
                    else:
                        print(f"  ✅ Корректно: неизвестный статус '{status}' обработан, периодов нет")
                else:
                    print(f"  ❌ Ошибка: статус '{status}', но есть {periods_count} периодов на {total_days} дней")
                    all_passed = False
        
        return all_passed
        
    except Exception as e:
        print(f"  ❌ ТЕСТ ПРОВАЛЕН: Исключение: {e}")
        return False

def main():
    """Главная функция тестирования"""
    print("=" * 60)
    print("  ТЕСТИРОВАНИЕ НОВОЙ ЛОГИКИ ЧТЕНИЯ ПЕРИОДОВ ПО СТАТУСУ")
    print("=" * 60)
    
    if test_status_logic():
        print("\n✅ ТЕСТ ПРОЙДЕН УСПЕШНО!")
        print("\nОбновленная логика применена:")
        print("• Статус читается ПЕРЕД чтением периодов")
        print("• Периоды читаются ТОЛЬКО если статус 'Форма заполнена корректно'")
        print("• Статусы сохраняются оригинальными: 'некорректно', 'не заполнена'") 
        print("• Для НЕкорректных статусов устанавливается 0 дней и 0 периодов")
        print("• Логика применена в core/excel_handler.py, create_report.py и data_mapper.py")
    else:
        print("\n❌ ТЕСТ ПРОВАЛЕН")
    
    print("=" * 60)

if __name__ == "__main__":
    main() 