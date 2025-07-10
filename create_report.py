#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Создание отчета по блоку - Автономный скрипт
Использует файлы сотрудников в текущей папке для создания отчета по подразделению
"""

import os
import sys
import time
import shutil
from pathlib import Path
from datetime import datetime, date
from typing import List, Optional, Dict
import re

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ОШИБКА: Не установлена библиотека openpyxl")
    print("Установите: pip install openpyxl")
    input("Нажмите Enter для выхода...")
    sys.exit(1)


class Employee:
    """Простая модель сотрудника"""
    def __init__(self):
        self.full_name = ""
        self.tab_number = ""
        self.position = ""
        self.department1 = ""
        self.department2 = ""
        self.department3 = ""
        self.department4 = ""


class VacationPeriod:
    """Период отпуска"""
    def __init__(self, start_date: date, end_date: date, days: int = 0):
        self.start_date = start_date
        self.end_date = end_date
        self.days = days if days > 0 else (end_date - start_date).days + 1


class VacationInfo:
    """Информация об отпусках сотрудника"""
    def __init__(self, employee: Employee, periods: List[VacationPeriod] = None):
        self.employee = employee
        self.periods = periods or []
        self.total_days = sum(period.days for period in self.periods)
        self.periods_count = len(self.periods)
        self.has_long_period = any(period.days >= 14 for period in self.periods)
        self.validation_errors = []
        
        # Определяем статус
        if not self.periods:
            self.status = "Не заполнено"
        elif self.validation_errors:
            self.status = "Ошибка"
        elif self.total_days >= 28 and self.has_long_period:
            self.status = "Ок"
        else:
            self.status = "Частично"


class BlockReportCreator:
    """Создатель отчетов по блокам"""
    
    # Константы для 2026 года
    DAYS_IN_MONTH_2026 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    MONTH_NAMES = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    
    def __init__(self):
        self.template_path = r"M:\Подразделения\АУП\Стажерская программа\Отпуск Р7\templates\block_report_template.xlsx"
        
    def validate_template(self) -> bool:
        """Проверяет наличие шаблона"""
        if not Path(self.template_path).exists():
            print(f"ОШИБКА: Шаблон не найден по пути: {self.template_path}")
            return False
        print(f"✓ Шаблон найден")
        return True
    
    def scan_employee_files(self, directory: str) -> List[str]:
        """Сканирует файлы сотрудников в папке"""
        employee_files = []
        directory_path = Path(directory)
        
        print(f"Сканирование папки: {directory_path.absolute()}")
        
        for file_path in directory_path.iterdir():
            if not file_path.is_file() or file_path.suffix.lower() != '.xlsx':
                continue
            
            filename = file_path.name
            
            # Исключаем отчеты и системные файлы
            if (filename.startswith('~') or 
                filename.startswith('!') or
                filename.startswith('Отчет') or
                filename.startswith('отчет') or
                filename.startswith('ОБЩИЙ') or
                'report' in filename.lower()):
                continue
            
            employee_files.append(str(file_path))
        
        print(f"✓ Найдено файлов сотрудников: {len(employee_files)}")
        return employee_files
    
    def parse_date(self, date_value) -> Optional[date]:
        """Парсит дату из различных форматов"""
        if not date_value:
            return None
        
        # Если уже date или datetime
        if isinstance(date_value, date):
            return date_value
        if isinstance(date_value, datetime):
            return date_value.date()
        
        # Если строка
        date_str = str(date_value).strip()
        if not date_str:
            return None
        
        # Попробуем различные форматы
        formats = ["%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"]
        
        for fmt in formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt).date()
                return parsed_date
            except ValueError:
                continue
        
        return None
    
    def get_cell_value(self, worksheet, cell_address: str):
        """Безопасно получает значение ячейки"""
        try:
            cell = worksheet[cell_address]
            return cell.value
        except Exception:
            return None
    
    def read_vacation_info_from_file(self, file_path: str) -> Optional[VacationInfo]:
        """Читает информацию об отпусках из файла сотрудника"""
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            # Читаем базовую информацию о сотруднике
            employee = Employee()
            
            # Ищем первую заполненную строку для получения базовой информации (строки 9-23)
            for row in range(9, 24):
                tab_number = self.get_cell_value(worksheet, f"B{row}")
                full_name = self.get_cell_value(worksheet, f"C{row}")
                position = self.get_cell_value(worksheet, f"D{row}")
                
                if tab_number and full_name:
                    employee.tab_number = str(tab_number).strip()
                    employee.full_name = str(full_name).strip()
                    if position:
                        employee.position = str(position).strip()
                    break
            
            # Читаем подразделения из шапки файла (C2:C5)
            employee.department1 = str(self.get_cell_value(worksheet, "C2") or "").strip()
            employee.department2 = str(self.get_cell_value(worksheet, "C3") or "").strip()
            employee.department3 = str(self.get_cell_value(worksheet, "C4") or "").strip()
            employee.department4 = str(self.get_cell_value(worksheet, "C5") or "").strip()
            
            # Читаем периоды отпусков из строк 9-23
            periods = []
            
            for row in range(9, 24):
                start_date_value = self.get_cell_value(worksheet, f"E{row}")
                end_date_value = self.get_cell_value(worksheet, f"F{row}")
                days_value = self.get_cell_value(worksheet, f"G{row}")
                
                if not start_date_value or not end_date_value:
                    continue
                
                try:
                    # Парсим даты
                    start_date = self.parse_date(start_date_value)
                    end_date = self.parse_date(end_date_value)
                    
                    if not start_date or not end_date:
                        continue
                    
                    # Парсим количество дней
                    days = 0
                    if days_value:
                        try:
                            days = int(days_value)
                        except (ValueError, TypeError):
                            days = (end_date - start_date).days + 1
                    else:
                        days = (end_date - start_date).days + 1
                    
                    period = VacationPeriod(start_date=start_date, end_date=end_date, days=days)
                    periods.append(period)
                    
                except Exception as e:
                    print(f"ПРЕДУПРЕЖДЕНИЕ: Ошибка обработки периода в строке {row}: {e}")
                    continue
            
            # Читаем результаты валидации
            validation_h2 = str(self.get_cell_value(worksheet, "H2") or "").strip()
            validation_i2 = str(self.get_cell_value(worksheet, "I2") or "").strip()
            validation_j2 = self.get_cell_value(worksheet, "J2") or 0
            
            # Создаем VacationInfo
            vacation_info = VacationInfo(employee=employee, periods=periods)
            
            # Определяем статус на основе валидаций
            if "ОШИБКА" in validation_h2:
                vacation_info.validation_errors.append(validation_h2)
            
            if "ОШИБКА" in validation_i2:
                vacation_info.validation_errors.append(validation_i2)
            
            try:
                total_days = int(validation_j2) if validation_j2 else 0
                if total_days < 28:
                    vacation_info.validation_errors.append(f"ОШИБКА: Недостаточно дней отпуска. Запланировано {total_days} дней, требуется минимум 28.")
            except (ValueError, TypeError):
                vacation_info.validation_errors.append("ОШИБКА: Не удалось определить общее количество дней отпуска.")
            
            # Обновляем статус
            if not vacation_info.validation_errors:
                vacation_info.status = "Ок"
            else:
                vacation_info.status = "Ошибка"
            
            workbook.close()
            return vacation_info
            
        except Exception as e:
            print(f"ОШИБКА: Не удалось прочитать файл {file_path}: {e}")
            return None
    
    def get_calendar_column(self, target_date: date, start_col: int) -> Optional[int]:
        """Вычисляет номер столбца для конкретной даты в календарной матрице"""
        if target_date.year != 2026:
            return None
        
        col_offset = 0
        # Считаем смещение по месяцам
        for month in range(1, target_date.month):
            col_offset += self.DAYS_IN_MONTH_2026[month - 1]
        
        # Добавляем день месяца
        col_offset += target_date.day - 1
        
        return start_col + col_offset
    
    def fill_calendar_matrix(self, worksheet, vacation_infos: List[VacationInfo]):
        """Заполняет календарную матрицу на листе Report"""
        try:
            start_col = 12  # Столбец L
            
            # Заполняем месяца в строке 7
            col_offset = 0
            for month_idx, month_name in enumerate(self.MONTH_NAMES):
                month_col = start_col + col_offset
                worksheet.cell(row=7, column=month_col, value=month_name)
                
                # Заполняем дни месяца в строке 8
                days_in_month = self.DAYS_IN_MONTH_2026[month_idx]
                for day in range(1, days_in_month + 1):
                    day_col = start_col + col_offset + day - 1
                    worksheet.cell(row=8, column=day_col, value=day)
                
                col_offset += days_in_month
            
            # Заполняем отпуска для каждого сотрудника
            for emp_idx, vacation_info in enumerate(vacation_infos):
                emp_row = emp_idx + 9  # Строка сотрудника
                
                for period in vacation_info.periods:
                    # Заполняем дни отпуска единицами
                    current_date = period.start_date
                    while current_date <= period.end_date:
                        if current_date.year == 2026:  # Только для 2026 года
                            day_col = self.get_calendar_column(current_date, start_col)
                            if day_col:
                                worksheet.cell(row=emp_row, column=day_col, value=1)
                        
                        # Переход к следующему дню
                        from datetime import timedelta
                        current_date = current_date + timedelta(days=1)
                        
                        if current_date > period.end_date:
                            break
                            
        except Exception as e:
            print(f"ОШИБКА: Ошибка заполнения календарной матрицы: {e}")
    
    def create_block_report(self, block_name: str, vacation_infos: List[VacationInfo], output_path: str) -> bool:
        """Создает отчет по блоку с календарной матрицей"""
        try:
            # Копируем шаблон
            shutil.copy2(self.template_path, output_path)
            
            # Открываем файл для заполнения
            workbook = openpyxl.load_workbook(output_path)
            
            # Заполняем лист Report
            self.fill_report_sheet(workbook, block_name, vacation_infos)
            
            # Заполняем лист Print
            self.fill_print_sheet(workbook, block_name, vacation_infos)
            
            workbook.save(output_path)
            workbook.close()
            
            print(f"✓ Отчет создан: {output_path}")
            return True
            
        except Exception as e:
            print(f"ОШИБКА: Ошибка создания отчета по блоку {block_name}: {e}")
            return False
    
    def fill_report_sheet(self, workbook, block_name: str, vacation_infos: List[VacationInfo]):
        """Заполняет лист Report"""
        if 'Report' not in workbook.sheetnames:
            print("ОШИБКА: Лист 'Report' не найден в шаблоне")
            return
        
        worksheet = workbook['Report']
        current_time = datetime.now()
        
        # Шапка A3:A6
        worksheet["A3"] = block_name
        worksheet["A4"] = f"Дата обновления: {current_time.strftime('%d.%m.%Y %H:%M')}"
        worksheet["A5"] = f"Количество сотрудников: {len(vacation_infos)}"
        
        # Подсчет завершивших планирование
        completed = sum(1 for vi in vacation_infos if vi.status == "Ок")
        percentage = (completed / len(vacation_infos) * 100) if vacation_infos else 0
        worksheet["A6"] = f"Закончили планирование: {completed} ({percentage:.0f}%)"
        
        # Заполняем таблицу сотрудников (начиная с строки 9)
        for i, vacation_info in enumerate(vacation_infos):
            row = i + 9
            emp = vacation_info.employee
            
            worksheet[f"A{row}"] = i + 1  # №
            worksheet[f"B{row}"] = emp.full_name  # ФИО
            worksheet[f"C{row}"] = emp.tab_number  # Таб. Номер
            worksheet[f"D{row}"] = emp.position  # Должность
            worksheet[f"E{row}"] = emp.department1  # Подразделение 1
            worksheet[f"F{row}"] = emp.department2  # Подразделение 2
            worksheet[f"G{row}"] = emp.department3  # Подразделение 3
            worksheet[f"H{row}"] = emp.department4  # Подразделение 4
            
            # Статус планирования
            if vacation_info.status == "Ок":
                worksheet[f"I{row}"] = "Ок"
            else:
                errors = vacation_info.validation_errors
                worksheet[f"I{row}"] = "\n".join(errors) if errors else "Ошибка"
            
            worksheet[f"J{row}"] = vacation_info.total_days  # Итого дней
            worksheet[f"K{row}"] = vacation_info.periods_count  # Кол-во периодов
            
            # Добавляем границы
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            # Применить границы к диапазону строки (столбцы A-K + календарь)
            for col in range(1, 378):  # A до столбца календаря
                worksheet.cell(row=row, column=col).border = thin_border
        
        # Заполняем календарную матрицу
        self.fill_calendar_matrix(worksheet, vacation_infos)
    
    def fill_print_sheet(self, workbook, block_name: str, vacation_infos: List[VacationInfo]):
        """Заполняет лист Print в нормализованном виде"""
        if 'Print' not in workbook.sheetnames:
            print("ОШИБКА: Лист 'Print' не найден в шаблоне")
            return
        
        worksheet = workbook['Print']
        
        # D4 - название блока
        worksheet["D4"] = block_name
        
        # Нормализуем данные - каждый период отпуска = отдельная строка
        normalized_data = []
        for vacation_info in vacation_infos:
            emp = vacation_info.employee
            
            if not vacation_info.periods:
                # Если нет периодов, добавляем пустую строку
                normalized_data.append({
                    'employee': emp,
                    'period_num': 0,
                    'start_date': None,
                    'end_date': None,
                    'days': 0
                })
            else:
                # Добавляем строку для каждого периода
                for period_idx, period in enumerate(vacation_info.periods, 1):
                    normalized_data.append({
                        'employee': emp,
                        'period_num': period_idx,
                        'start_date': period.start_date,
                        'end_date': period.end_date,
                        'days': period.days
                    })
        
        # Заполняем данные с учетом разбивки по страницам
        current_row = 9  # Начинаем с 9 строки
        records_on_page = 0
        max_records_first_page = 14
        max_records_other_pages = 18
        is_first_page = True
        
        for record_idx, record in enumerate(normalized_data):
            # Проверяем нужность новой страницы
            max_records = max_records_first_page if is_first_page else max_records_other_pages
            
            if records_on_page >= max_records:
                # Добавляем заголовки на новой странице
                current_row += 1  # Пропускаем строку
                self.add_print_headers(worksheet, current_row)
                current_row += 1
                records_on_page = 0
                is_first_page = False
            
            # Заполняем строку данных
            emp = record['employee']
            
            worksheet[f"A{current_row}"] = record_idx + 1  # № п/п
            worksheet[f"B{current_row}"] = emp.tab_number  # Табельный номер
            worksheet[f"C{current_row}"] = emp.full_name  # ФИО
            worksheet[f"D{current_row}"] = emp.position  # Должность
            
            if record['start_date']:
                worksheet[f"E{current_row}"] = record['start_date'].strftime('%d.%m.%Y')  # Дата начала
                worksheet[f"F{current_row}"] = record['end_date'].strftime('%d.%m.%Y')  # Дата окончания
                worksheet[f"G{current_row}"] = record['days']  # Продолжительность
            
            current_row += 1
            records_on_page += 1
    
    def add_print_headers(self, worksheet, row: int):
        """Добавляет заголовки таблицы для печати"""
        headers = [
            "№ п/п", "Табельный номер", "ФИО", "Должность",
            "Дата начала отпуска", "Дата окончания отпуска",
            "Продолжительность (календарных дней)",
            "Подпись работника", "Дата ознакомления работника", "Примечание"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            worksheet.cell(row=row, column=col_idx, value=header)


def main():
    """Главная функция скрипта"""
    print("=" * 60)
    print("  СОЗДАНИЕ ОТЧЕТА ПО БЛОКУ")
    print("  Автономный скрипт для создания отчетов по подразделениям")
    print("=" * 60)
    print()
    
    creator = BlockReportCreator()
    
    # 1. Проверяем шаблон
    print("1. Проверка шаблона...")
    if not creator.validate_template():
        input("Нажмите Enter для выхода...")
        return
    
    # 2. Определяем текущую папку
    current_dir = os.getcwd()
    print(f"2. Текущая папка: {current_dir}")
    
    # 3. Сканируем файлы сотрудников
    print("3. Поиск файлов сотрудников...")
    employee_files = creator.scan_employee_files(current_dir)
    
    if not employee_files:
        print("ОШИБКА: В текущей папке не найдено файлов сотрудников (.xlsx)")
        input("Нажмите Enter для выхода...")
        return
    
    # 4. Читаем данные из файлов
    print("4. Чтение данных из файлов...")
    vacation_infos = []
    invalid_files = []
    
    for i, file_path in enumerate(employee_files, 1):
        print(f"   Обработка {i}/{len(employee_files)}: {Path(file_path).name}")
        vacation_info = creator.read_vacation_info_from_file(file_path)
        
        if vacation_info:
            vacation_infos.append(vacation_info)
        else:
            invalid_files.append(Path(file_path).name)
    
    if not vacation_infos:
        print("ОШИБКА: Не удалось прочитать ни одного файла сотрудника")
        input("Нажмите Enter для выхода...")
        return
    
    print(f"✓ Успешно обработано файлов: {len(vacation_infos)} из {len(employee_files)}")
    
    if invalid_files:
        print(f"Сотрудников с неверно заполненным файлом: {len(invalid_files)}")
        for invalid_file in invalid_files:
            print(f"   • {invalid_file}")
    
    # 5. Определяем название блока из первого сотрудника
    block_name = vacation_infos[0].employee.department1 or "Неизвестное подразделение"
    print(f"5. Название блока: {block_name}")
    
    # 6. Создаем отчет
    print("6. Создание отчета...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"Отчет по блоку_{block_name}_{timestamp}.xlsx"
    
    # Очищаем имя файла от недопустимых символов
    invalid_chars = r'[\\/:*?"<>|]'
    output_filename = re.sub(invalid_chars, '_', output_filename)
    output_path = Path(current_dir) / output_filename
    
    success = creator.create_block_report(block_name, vacation_infos, str(output_path))
    
    if success:
        print()
        print("=" * 60)
        print("  ОТЧЕТ УСПЕШНО СОЗДАН!")
        print("=" * 60)
        print(f"Файл: {output_filename}")
        print(f"Подразделение: {block_name}")
        print(f"Сотрудников: {len(vacation_infos)}")
        
        # Статистика по статусам с детализацией ошибок
        status_counts = {}
        error_types = {}
        
        for vi in vacation_infos:
            status = vi.status
            status_counts[status] = status_counts.get(status, 0) + 1
            
            # Детализация ошибок
            if vi.validation_errors:
                for error in vi.validation_errors:
                    if "Недостаточно дней отпуска" in error:
                        error_types["Недостаточно дней отпуска"] = error_types.get("Недостаточно дней отпуска", 0) + 1
                    elif "пересечение периодов" in error.lower():
                        error_types["Пересечение периодов"] = error_types.get("Пересечение периодов", 0) + 1
                    elif "Не удалось определить" in error:
                        error_types["Проблемы с подсчетом дней"] = error_types.get("Проблемы с подсчетом дней", 0) + 1
                    else:
                        error_types["Прочие ошибки"] = error_types.get("Прочие ошибки", 0) + 1
        
        print("Статистика планирования:")
        for status, count in status_counts.items():
            print(f"  {status}: {count}")
        
        if error_types:
            print("Типы ошибок:")
            for error_type, count in error_types.items():
                print(f"  {error_type}: {count}")
        
        print()
        print("Отчет создан в текущей папке.")
    else:
        print("ОШИБКА: Не удалось создать отчет")
    
    print()
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    main()