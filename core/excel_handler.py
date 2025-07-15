#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль работы с Excel файлами
"""

import logging
import shutil
from pathlib import Path
from datetime import datetime, date
from typing import List, Optional, Dict, Tuple
import re

import openpyxl
from openpyxl.styles import Border, Side

from models import Employee, VacationInfo, VacationPeriod, VacationStatus
from config import Config


class ExcelHandler:
    """Класс для работы с Excel файлами"""
    
    # Константы для 2026 года (не високосный)
    DAYS_IN_MONTH_2026 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    MONTH_NAMES = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)

    def _parse_cell_address(self, address: str) -> Tuple[bool, str, Optional[str]]:
        """Парсит адрес ячейки"""
        address = address.strip()
        
        if address.startswith('#'):
            raise ValueError(f"Недопустимый адрес ячейки: {address}")
        
        if address.startswith('='):
            clean_address = address[1:]
            if '!' in clean_address:
                parts = clean_address.split('!')
                if len(parts) != 2:
                    raise ValueError(f"Неверный формат ссылки: {address}")
                sheet_name = parts[0].strip("'\"")
                cell_ref = parts[1]
                return True, cell_ref, sheet_name
            else:
                return True, clean_address, None
        else:
            return False, address, None

    def _fill_cell_or_range(self, workbook, sheet_name: Optional[str], cell_ref: str, value: any):
        """Заполняет ячейку или диапазон значением"""
        worksheet = workbook[sheet_name] if sheet_name and sheet_name in workbook.sheetnames else workbook.active
        
        if ';' in cell_ref:
            for single_ref in cell_ref.split(';'):
                single_ref = single_ref.strip()
                if single_ref:
                    worksheet[single_ref] = value
        elif ':' in cell_ref:
            cell_range = worksheet[cell_ref]
            if hasattr(cell_range, '__iter__'):
                for row in cell_range:
                    if hasattr(row, '__iter__'):
                        for cell in row:
                            cell.value = value
                    else:
                        row.value = value
            else:
                cell_range.value = value
        else:
            worksheet[cell_ref] = value

    def _load_filling_rules(self, template_path: str) -> Dict[str, Dict[str, str]]:
        """Загружает правила заполнения из листа 'rules'"""
        rules = {'value': {}, 'header': {}, 'read': {}}
        
        workbook = openpyxl.load_workbook(template_path, data_only=False)
        
        if 'rules' not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Лист 'rules' не найден в шаблоне {template_path}")
        
        rules_sheet = workbook['rules']
        
        for row in range(2, rules_sheet.max_row + 1):
            source_cell = rules_sheet.cell(row=row, column=1)
            target_cell = rules_sheet.cell(row=row, column=2)
            type_cell = rules_sheet.cell(row=row, column=3)
            
            source_address = source_cell.value if source_cell.data_type == 'f' else source_cell.value
            target_field = target_cell.value
            rule_type = type_cell.value
            
            if source_address and target_field and rule_type:
                source_address = str(source_address).strip()
                target_field = str(target_field).strip()
                rule_type = str(rule_type).strip().lower()
                
                if rule_type in ['value', 'header', 'read']:
                    rules[rule_type][source_address] = target_field
        
        workbook.close()
        
        if not any(rules.values()):
            raise ValueError(f"Лист 'rules' пуст или не содержит корректных правил в {template_path}")
        
        return rules

    def _apply_rules_to_template(self, workbook, rules: Dict[str, Dict[str, str]], data_dict: Dict[str, any]):
        """Применяет правила заполнения к шаблону"""
        for rule_type, rule_items in rules.items():
            for cell_address, field_name in rule_items.items():
                if rule_type == 'value':
                    # Для value - только значение без заголовка
                    value = data_dict.get(field_name, '')
                elif rule_type == 'header':
                    # Для header - НЕ заменяем на field_name, оставляем как есть в шаблоне
                    continue  # Пропускаем header rules при заполнении
                elif rule_type == 'read':
                    # Для read - пропускаем, это для чтения данных из файлов сотрудников
                    continue
                else:
                    continue
                
                is_formula, clean_address, sheet_name = self._parse_cell_address(cell_address)
                self._fill_cell_or_range(workbook, sheet_name, clean_address, value)

    def create_employee_file(self, employee: Employee, output_path: str) -> bool:
        """Создает файл сотрудника на основе шаблона с rules"""
        template_path = Path(self.config.employee_template)
        if not template_path.exists():
            raise FileNotFoundError(f"Шаблон сотрудника не найден: {template_path}")
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)
        
        rules = self._load_filling_rules(str(template_path))
        
        employee_data = {
            'ФИО работника': employee.full_name,
            'Табельный номер': employee.tab_number,
            'Должность': getattr(employee, 'position', ''),
            'Подразделение 1': employee.department1,
            'Подразделение 2': employee.department2,
            'Подразделение 3': employee.department3,
            'Подразделение 4': employee.department4,
            'Локация': getattr(employee, 'location', ''),
            'Остатки отпуска': getattr(employee, 'vacation_remainder', ''),
            'Дата приема': getattr(employee, 'hire_date', ''),
            'Дата отсечки периода': getattr(employee, 'period_cutoff_date', ''),
            'Дополнительный отпуск НРД': getattr(employee, 'additional_vacation_nrd', ''),
            'Дополнительный отпуск Северный': getattr(employee, 'additional_vacation_north', '')
        }
        
        workbook = openpyxl.load_workbook(output_path)
        self._apply_rules_to_template(workbook, rules, employee_data)
        workbook.save(output_path)
        workbook.close()
        
        return True

    def read_vacation_info_from_file(self, file_path: str) -> Optional[VacationInfo]:
        """Читает информацию об отпусках из файла сотрудника"""
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            employee = Employee()
            employee.file_path = file_path  # Сохраняем путь к файлу для чтения статуса
            
            # Ищем первую заполненную строку для базовой информации
            for row in range(9, 24):
                tab_number = self._get_cell_value(worksheet, f"B{row}")
                full_name = self._get_cell_value(worksheet, f"C{row}")
                position = self._get_cell_value(worksheet, f"D{row}")
                
                if tab_number and full_name:
                    employee.tab_number = str(tab_number).strip()
                    employee.full_name = str(full_name).strip()
                    if position:
                        employee.position = str(position).strip()
                    break
            
            # Читаем подразделения из шапки файла
            employee.department1 = str(self._get_cell_value(worksheet, "C2") or "").strip()
            employee.department2 = str(self._get_cell_value(worksheet, "C3") or "").strip()
            employee.department3 = str(self._get_cell_value(worksheet, "C4") or "").strip()
            employee.department4 = str(self._get_cell_value(worksheet, "C5") or "").strip()
            
            # Читаем периоды отпусков
            periods = []
            for row in range(9, 24):
                start_date_value = self._get_cell_value(worksheet, f"E{row}")
                end_date_value = self._get_cell_value(worksheet, f"F{row}")
                days_value = self._get_cell_value(worksheet, f"G{row}")
                
                if not start_date_value or not end_date_value:
                    continue
                
                try:
                    start_date = self._parse_date(start_date_value)
                    end_date = self._parse_date(end_date_value)
                    
                    if not start_date or not end_date:
                        continue
                    
                    days = int(days_value) if days_value else (end_date - start_date).days + 1
                    periods.append(VacationPeriod(start_date=start_date, end_date=end_date, days=days))
                    
                except Exception as e:
                    self.logger.warning(f"Ошибка обработки периода в строке {row}: {e}")
                    continue
            
            # Читаем результаты валидации
            validation_errors = []
            validation_h2 = str(self._get_cell_value(worksheet, "H2") or "").strip()
            validation_i2 = str(self._get_cell_value(worksheet, "I2") or "").strip()
            validation_j2 = self._get_cell_value(worksheet, "J2") or 0
            
            if "ОШИБКА" in validation_h2:
                validation_errors.append(validation_h2)
            if "ОШИБКА" in validation_i2:
                validation_errors.append(validation_i2)
            
            try:
                total_days = int(validation_j2) if validation_j2 else 0
                if total_days < 28:
                    validation_errors.append(f"ОШИБКА: Недостаточно дней отпуска. Запланировано {total_days} дней, требуется минимум 28.")
            except (ValueError, TypeError):
                validation_errors.append("ОШИБКА: Не удалось определить общее количество дней отпуска.")
            
            vacation_info = VacationInfo(employee=employee, periods=periods)
            vacation_info.validation_errors = validation_errors
            vacation_info.status = VacationStatus.OK if not validation_errors else VacationStatus.ERROR
            
            workbook.close()
            return vacation_info
            
        except Exception as e:
            self.logger.error(f"Ошибка чтения файла {file_path}: {e}")
            return None

    def create_block_report(self, block_name: str, vacation_infos: List[VacationInfo], output_path: str) -> bool:
        """Создает отчет по блоку с использованием rules"""
        # Пробуем новый шаблон, fallback на старый
        template_path = Path(self.config.block_report_template.replace("block_report_template.xlsx", "block_report_template v3.xlsx"))
        if not template_path.exists():
            template_path = Path(self.config.block_report_template)
            
        if not template_path.exists():
            raise FileNotFoundError(f"Шаблон отчета не найден: {template_path}")
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)
        
        rules = self._load_filling_rules(str(template_path))
        workbook = openpyxl.load_workbook(output_path)
        
        self._fill_report_with_rules(workbook, block_name, vacation_infos, rules)
        
        workbook.save(output_path)
        workbook.close()
        
        self.logger.info(f"Создан отчет по блоку: {output_path}")
        return True

    def _fill_report_with_rules(self, workbook, block_name: str, vacation_infos: List[VacationInfo], rules: Dict[str, Dict[str, str]]):
        """Заполняет отчет используя rules"""
        current_time = datetime.now()
        total_employees = len(vacation_infos)
        
        # Правильная логика подсчета на основе статусов из файлов сотрудников
        employees_filled = 0  # Статус НЕ "Форма не заполнена"
        employees_correct = 0  # Статус "Форма заполнена корректно"
        
        for vi in vacation_infos:
            # Читаем статус из файла сотрудника используя read rules
            employee_status = self._read_employee_status(vi, rules)
            
            if employee_status and employee_status != "Форма не заполнена":
                employees_filled += 1
                
            if employee_status == "Форма заполнена корректно":
                employees_correct += 1
        
        # Данные для заполнения - только значения без заголовков
        report_data = {
            'block_name': block_name,
            'update_date': current_time.strftime('%d.%m.%Y %H:%M'),  # Только время
            'total_employees': str(total_employees),  # Только цифра
            'employees_filled': str(employees_filled),  # Только цифра
            'employees_correct': str(employees_correct),  # Только цифра
        }
        
        # Применяем rules
        self._apply_rules_to_template(workbook, rules, report_data)
        
        # Заполняем таблицы данных
        self._fill_employee_tables(workbook, vacation_infos, rules)
        
        # Заполняем календарь на Report листе
        if 'Report' in workbook.sheetnames:
            self._fill_calendar_matrix(workbook['Report'], vacation_infos)

    def _read_employee_status(self, vacation_info: VacationInfo, rules: Dict[str, Dict[str, str]]) -> Optional[str]:
        """Читает статус сотрудника из его файла используя read rules"""
        if not vacation_info.employee.file_path:
            return None
            
        try:
            # Ищем правило для чтения статуса
            read_rules = rules.get('read', {})
            status_cell = None
            
            for cell_address, field_name in read_rules.items():
                if field_name == 'report_status':
                    # Парсим адрес ячейки для чтения статуса
                    is_formula, clean_address, sheet_name = self._parse_cell_address(cell_address)
                    status_cell = clean_address
                    break
            
            if not status_cell:
                return None
            
            # Читаем статус из файла сотрудника
            workbook = openpyxl.load_workbook(vacation_info.employee.file_path, data_only=True)
            worksheet = workbook.active
            
            status_value = self._get_cell_value(worksheet, status_cell)
            workbook.close()
            
            return str(status_value).strip() if status_value else None
            
        except Exception as e:
            self.logger.warning(f"Ошибка чтения статуса для {vacation_info.employee.full_name}: {e}")
            return None

    def _fill_employee_tables(self, workbook, vacation_infos: List[VacationInfo], rules: Dict[str, Dict[str, str]]):
        """Заполняет таблицы сотрудников на Report и Print листах"""
        # Report лист - таблица сотрудников
        if 'Report' in workbook.sheetnames:
            self._fill_table_by_prefix(workbook['Report'], vacation_infos, rules, 'report_', self._get_report_row_data)
        
        # Print лист - нормализованная таблица периодов
        if 'Print' in workbook.sheetnames:
            normalized_data = self._normalize_vacation_data(vacation_infos)
            self._fill_table_by_prefix(workbook['Print'], normalized_data, rules, 'print_', self._get_print_row_data)
            self._apply_borders_to_table(workbook['Print'], len(normalized_data))

    def _fill_table_by_prefix(self, worksheet, data_list: List, rules: Dict[str, Dict[str, str]], prefix: str, row_data_func):
        """Универсальная функция заполнения таблицы по префиксу"""
        header_rules = rules.get('header', {})
        column_mapping = {}
        base_row = 9
        
        # Определяем mapping столбцов и базовую строку
        for cell_address, field_name in header_rules.items():
            if field_name.startswith(prefix):
                try:
                    is_formula, clean_address, sheet_name = self._parse_cell_address(cell_address)
                    if ':' not in clean_address and ';' not in clean_address:
                        col_match = re.search(r'([A-Z]+)', clean_address)
                        row_match = re.search(r'(\d+)', clean_address)
                        
                        if col_match:
                            column_mapping[field_name] = col_match.group(1)
                        
                        if field_name == f'{prefix}row_number' and row_match:
                            base_row = int(row_match.group(1)) + 1
                except:
                    continue
        
        # Заполняем данные
        for i, data_item in enumerate(data_list):
            row = base_row + i
            row_data = row_data_func(data_item, i)
            
            for field_name, value in row_data.items():
                full_field_name = f'{prefix}{field_name}'
                if full_field_name in column_mapping:
                    worksheet[f"{column_mapping[full_field_name]}{row}"] = value

    def _get_report_row_data(self, vacation_info: VacationInfo, index: int) -> Dict[str, any]:
        """Возвращает данные строки для Report листа"""
        emp = vacation_info.employee
        return {
            'row_number': index + 1,
            'employee_name': emp.full_name,
            'tab_number': emp.tab_number,
            'position': getattr(emp, 'position', ''),
            'department1': emp.department1,
            'department2': emp.department2,
            'department3': emp.department3,
            'department4': emp.department4,
            'status': "Ок" if vacation_info.status == VacationStatus.OK else "\n".join(getattr(vacation_info, 'validation_errors', []) or ["Ошибка"]),
            'total_days': vacation_info.total_days,
            'periods_count': vacation_info.periods_count
        }

    def _get_print_row_data(self, normalized_record: Dict, index: int) -> Dict[str, any]:
        """Возвращает данные строки для Print листа"""
        emp = normalized_record['employee']
        return {
            'row_number': index + 1,
            'tab_number': emp.tab_number,
            'employee_name': emp.full_name,
            'position': getattr(emp, 'position', ''),
            'start_date': normalized_record['start_date'].strftime('%d.%m.%Y') if normalized_record['start_date'] else '',
            'end_date': normalized_record['end_date'].strftime('%d.%m.%Y') if normalized_record['end_date'] else '',
            'duration': normalized_record['days'] if normalized_record['days'] > 0 else ''
        }

    def _normalize_vacation_data(self, vacation_infos: List[VacationInfo]) -> List[Dict]:
        """Нормализует данные отпусков - каждый период = отдельная строка"""
        normalized_data = []
        for vacation_info in vacation_infos:
            emp = vacation_info.employee
            if not vacation_info.periods:
                normalized_data.append({
                    'employee': emp, 'period_num': 0, 'start_date': None, 'end_date': None, 'days': 0
                })
            else:
                for period_idx, period in enumerate(vacation_info.periods, 1):
                    normalized_data.append({
                        'employee': emp, 'period_num': period_idx,
                        'start_date': period.start_date, 'end_date': period.end_date, 'days': period.days
                    })
        return normalized_data

    def _apply_borders_to_table(self, worksheet, data_count: int):
        """Применяет границы к таблице Print листа"""
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        
        for row in range(9, 9 + data_count):
            for col in range(1, 11):  # A-J
                worksheet.cell(row=row, column=col).border = thin_border

    def _fill_calendar_matrix(self, worksheet, vacation_infos: List[VacationInfo]):
        """Заполняет календарную матрицу"""
        start_col = 12
        col_offset = 0
        
        for month_idx, month_name in enumerate(self.MONTH_NAMES):
            month_col = start_col + col_offset
            worksheet.cell(row=7, column=month_col, value=month_name)
            
            days_in_month = self.DAYS_IN_MONTH_2026[month_idx]
            for day in range(1, days_in_month + 1):
                day_col = start_col + col_offset + day - 1
                worksheet.cell(row=8, column=day_col, value=day)
            
            col_offset += days_in_month
        
        for emp_idx, vacation_info in enumerate(vacation_infos):
            emp_row = emp_idx + 9
            for period in vacation_info.periods:
                current_date = period.start_date
                while current_date <= period.end_date:
                    if current_date.year == 2026:
                        day_col = self._get_calendar_column(current_date, start_col)
                        if day_col:
                            worksheet.cell(row=emp_row, column=day_col, value=1)
                    
                    from datetime import timedelta
                    current_date = current_date + timedelta(days=1)
                    if current_date > period.end_date:
                        break

    def _get_calendar_column(self, target_date: date, start_col: int) -> Optional[int]:
        """Вычисляет столбец для даты в календаре"""
        if target_date.year != 2026:
            return None
        
        col_offset = sum(self.DAYS_IN_MONTH_2026[:target_date.month - 1]) + target_date.day - 1
        return start_col + col_offset

    def read_block_report_data(self, report_path: str) -> Optional[Dict]:
        """Читает данные из отчета по блоку"""
        try:
            workbook = openpyxl.load_workbook(report_path, data_only=True)
            if 'Report' not in workbook.sheetnames:
                return None
            
            worksheet = workbook['Report']
            
            block_name = str(self._get_cell_value(worksheet, "A3") or "").strip()
            update_date_raw = str(self._get_cell_value(worksheet, "A4") or "").strip()
            total_employees_raw = str(self._get_cell_value(worksheet, "A5") or "").strip()
            completed_raw = str(self._get_cell_value(worksheet, "A6") or "").strip()
            
            update_date = update_date_raw.replace("Дата обновления:", "").strip() if "Дата обновления:" in update_date_raw else ""
            
            total_employees = 0
            if "Количество сотрудников:" in total_employees_raw:
                try:
                    total_employees = int(total_employees_raw.split(":")[1].strip())
                except (ValueError, IndexError):
                    pass
            
            completed_employees = 0
            percentage = 0
            if "Закончили планирование:" in completed_raw:
                try:
                    parts = completed_raw.split(":")[1].strip().split("(")
                    completed_employees = int(parts[0].strip())
                    if len(parts) > 1:
                        percentage = int(parts[1].replace("%)", "").strip())
                except (ValueError, IndexError):
                    pass
            
            workbook.close()
            
            return {
                'block_name': block_name,
                'total_employees': total_employees,
                'completed_employees': completed_employees,
                'remaining_employees': total_employees - completed_employees,
                'percentage': percentage,
                'update_date': update_date
            }
            
        except Exception as e:
            self.logger.error(f"Ошибка чтения отчета {report_path}: {e}")
            return None

    def create_general_report_from_blocks(self, block_data: List[Dict], output_path: str) -> bool:
        """Создает общий отчет"""
        template_path = Path(self.config.general_report_template)
        if not template_path.exists():
            raise FileNotFoundError(f"Шаблон общего отчета не найден: {template_path}")
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)
        
        workbook = openpyxl.load_workbook(output_path)
        worksheet = workbook.active
        
        # Заполняем данные
        for i, data in enumerate(block_data):
            row = 6 + i
            if i > 0:  # Вставляем строки для дополнительных блоков
                worksheet.insert_rows(row, 1)
            
            worksheet[f"A{row}"] = i + 1
            worksheet[f"B{row}"] = data['block_name']
            worksheet[f"C{row}"] = f"{data['percentage']}%"
            worksheet[f"D{row}"] = data['total_employees']
            worksheet[f"E{row}"] = data['completed_employees']
            worksheet[f"F{row}"] = data['remaining_employees']
            worksheet[f"G{row}"] = data['update_date']
        
        # Границы
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        for i in range(len(block_data)):
            row = 6 + i
            for col in range(1, 8):
                worksheet.cell(row=row, column=col).border = thin_border
        
        # Формулы итого
        if block_data:
            total_row = 6 + len(block_data) + 1
            data_end_row = 6 + len(block_data) - 1
            worksheet[f"C{total_row}"] = f'=ROUND(E{total_row}/D{total_row}*100,0)&"%"'
            worksheet[f"D{total_row}"] = f'=SUM(D6:D{data_end_row})'
            worksheet[f"E{total_row}"] = f'=SUM(E6:E{data_end_row})'
            worksheet[f"F{total_row}"] = f'=SUM(F6:F{data_end_row})'
        
        workbook.save(output_path)
        workbook.close()
        return True

    def _get_cell_value(self, worksheet, cell_address: str):
        """Безопасно получает значение ячейки"""
        try:
            return worksheet[cell_address].value
        except Exception:
            return None

    def _parse_date(self, date_value) -> Optional[date]:
        """Парсит дату из различных форматов"""
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

    def generate_output_filename(self, employee: Employee) -> str:
        """Генерирует имя файла для сотрудника"""
        clean_fio = self._clean_filename(employee.full_name)
        clean_tab_num = self._clean_filename(employee.tab_number)
        return f"{clean_fio} ({clean_tab_num}).xlsx"

    def generate_block_report_filename(self, block_name: str) -> str:
        """Генерирует имя файла отчета по блоку"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        clean_block_name = self._clean_filename(block_name)
        return f"Отчет по блоку_{clean_block_name}_{timestamp}.xlsx"

    def _clean_filename(self, filename: str) -> str:
        """Очищает имя файла от недопустимых символов"""
        if not filename:
            return "unnamed"
        
        clean_name = re.sub(r'[\\/:*?"<>|]', '_', filename).strip()
        return clean_name[:100] if len(clean_name) > 100 else clean_name or "unnamed"
    

    