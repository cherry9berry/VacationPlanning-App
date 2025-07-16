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


    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        # ДОБАВЛЯЕМ КЭШИРОВАНИЕ RULES
        self._cached_rules = {}
        self._cached_templates = {}
    
    def _get_cached_rules(self, template_path: str) -> Dict[str, Dict[str, str]]:
        """Получает rules из кэша или загружает их"""
        if template_path not in self._cached_rules:
            print(f"=== DEBUG EXCEL: Загружаем rules в кэш для {template_path} ===")
            self._cached_rules[template_path] = self._load_filling_rules(template_path)
        else:
            print(f"=== DEBUG EXCEL: Используем rules из кэша для {template_path} ===")
        
        return self._cached_rules[template_path]
        
    def create_employee_file(self, employee: Employee, output_path: str) -> bool:
        """Создает файл сотрудника на основе шаблона с rules"""
        print(f"=== DEBUG EXCEL: create_employee_file вызван для {employee['ФИО работника']} ===")
        print(f"=== DEBUG EXCEL: output_path = {output_path} ===")
     
        template_path = Path(self.config.employee_template)
        print(f"=== DEBUG EXCEL: template_path = {template_path} ===")
        
        if not template_path.exists():
            print(f"=== DEBUG EXCEL: ОШИБКА - Шаблон не найден: {template_path} ===")
            raise FileNotFoundError(f"Шаблон сотрудника не найден: {template_path}")
        
        print(f"=== DEBUG EXCEL: Создаем папку для файла ===")
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        print(f"=== DEBUG EXCEL: Копируем шаблон ===")
        shutil.copy2(template_path, output_path)
        
        print(f"=== DEBUG EXCEL: Загружаем rules из кэша ===")
        try:
            # ИСПОЛЬЗУЕМ КЭШИРОВАННЫЕ RULES
            rules = self._get_cached_rules(str(template_path))
            print(f"=== DEBUG EXCEL: Rules загружены: {len(rules)} типов ===")
            
            # Проверяем что в rules есть данные
            total_rules = sum(len(rule_items) for rule_items in rules.values())
            print(f"=== DEBUG EXCEL: Всего правил: {total_rules} ===")
            
            if total_rules == 0:
                print(f"=== DEBUG EXCEL: ВНИМАНИЕ - Rules пусты! ===")
                return False
                
        except Exception as e:
            print(f"=== DEBUG EXCEL: ОШИБКА при загрузке rules: {e} ===")
            import traceback
            traceback.print_exc()
            return False
        
        print(f"=== DEBUG EXCEL: Подготавливаем данные сотрудника ===")
        # Динамическое формирование data_dict на основе rules['value']
        def field_name_to_attr(field_name):
            # Маппинг для несовпадающих имён (дополняй по необходимости)
            mapping = {
                'ФИО работника': 'full_name',
                'Табельный номер': 'tab_number',
                'Должность': 'position',
                'Подразделение 1': 'department1',
                'Подразделение 2': 'department2',
                'Подразделение 3': 'department3',
                'Подразделение 4': 'department4',
                'Локация графика работы': 'location',
                'Дата приема на работу': 'hire_date',
                'Основной отпуск к дате отсечки': 'vacation_remainder',
                'Дополнительный отпуск НРД к дате отсечки': 'additional_vacation_nrd',
                'Дополнительный северный отпуск к дате отсечки': 'additional_vacation_north',
                'Дата выгрузки': 'period_cutoff_date',
                # Добавь сюда все нестандартные поля, если появятся
            }
            return mapping.get(field_name, field_name)

        data_dict = {}
        for field_name in rules.get('value', {}).values():
            value = employee.get(field_name, '')
            if value is None:
                value = ''
            data_dict[field_name] = value
        print(f"=== DEBUG EXCEL: Данные подготовлены: {len(data_dict)} полей ===")
        print(f"=== DEBUG EXCEL: Данные сотрудника: {data_dict} ===")
        
        print(f"=== DEBUG EXCEL: Открываем файл для заполнения ===")
        try:
            # ОПТИМИЗАЦИЯ: Используем более быстрые параметры
            workbook = openpyxl.load_workbook(
                output_path,
                data_only=False,  # Сохраняем формулы
                read_only=False,
                keep_links=False  # Не читать внешние ссылки
            )
            print(f"=== DEBUG EXCEL: Файл открыт для заполнения ===")
            
            print(f"=== DEBUG EXCEL: Применяем rules к файлу ===")
            self._apply_rules_to_template(workbook, rules, data_dict)
            
            print(f"=== DEBUG EXCEL: Сохраняем файл ===")
            workbook.save(output_path)
            workbook.close()
            
            print(f"=== DEBUG EXCEL: Файл сохранен успешно ===")
            # Автотест: проверить, что реально записалось в файл
            self.test_created_employee_file(output_path, rules)
            return True
            
        except Exception as e:
            print(f"=== DEBUG EXCEL: ОШИБКА при заполнении файла: {e} ===")
            import traceback
            traceback.print_exc()
            return False

    def _apply_rules_to_template(self, workbook, rules: Dict[str, Dict[str, str]], data_dict: Dict[str, any]):
        """Применяет правила заполнения к шаблону"""
        print(f"=== DEBUG EXCEL: Применяем rules, типов: {len(rules)} ===")
        
        for rule_type, rule_items in rules.items():
            print(f"=== DEBUG EXCEL: Обрабатываем тип {rule_type}, правил: {len(rule_items)} ===")
            
            for cell_address, field_name in rule_items.items():
                if rule_type == 'value':
                    # Для value - только значение без заголовка
                    value = data_dict.get(field_name, '')
                    print(f"=== DEBUG: Заполняем {cell_address} = '{value}' (поле: {field_name}) ===")
                    
                    try:
                        is_formula, clean_address, sheet_name = self._parse_cell_address(cell_address)
                        print(f"=== DEBUG: Парсинг: формула={is_formula}, адрес='{clean_address}', лист='{sheet_name}' ===")
                        
                        self._fill_cell_or_range(workbook, sheet_name, clean_address, value)
                        print(f"=== DEBUG: Успешно заполнено {cell_address} значением '{value}' ===")
                        
                    except Exception as e:
                        print(f"=== DEBUG: ОШИБКА при заполнении {cell_address}: {e} ===")
                        import traceback
                        traceback.print_exc()
                        
                elif rule_type == 'header':
                    # Для header - НЕ заменяем на field_name, оставляем как есть в шаблоне
                    print(f"=== DEBUG: Пропускаем header rule: {cell_address} -> {field_name} ===")
                    continue
                elif rule_type == 'read':
                    # Для read - пропускаем, это для чтения данных из файлов сотрудников
                    print(f"=== DEBUG: Пропускаем read rule: {cell_address} -> {field_name} ===")
                    continue
                else:
                    print(f"=== DEBUG: Неизвестный тип rule: {rule_type} ===")
                    continue

    def _parse_cell_address(self, address: str) -> tuple:
        """Парсит адрес ячейки, возвращает (is_formula, clean_address, sheet_name)"""
        print(f"=== DEBUG: Парсинг адреса '{address}' ===")
        
        is_formula = address.startswith('=')
        
        if is_formula:
            # Убираем знак равенства и извлекаем адрес
            formula = address[1:]
            print(f"=== DEBUG: Это формула: '{formula}' ===")
            
            # Простой парсинг для формул вида 'Лист'!A1
            if '!' in formula:
                sheet_part, cell_part = formula.split('!', 1)
                sheet_name = sheet_part.strip("'\"")
                clean_address = cell_part.strip()
                print(f"=== DEBUG: Лист '{sheet_name}', адрес '{clean_address}' ===")
            else:
                sheet_name = None
                clean_address = formula.strip()
                print(f"=== DEBUG: Без листа, адрес '{clean_address}' ===")
        else:
            # Обычный адрес - может быть именованный диапазон
            sheet_name = None
            clean_address = address.strip()
            print(f"=== DEBUG: Обычный адрес '{clean_address}' ===")
        
        print(f"=== DEBUG: Результат парсинга: формула={is_formula}, адрес='{clean_address}', лист='{sheet_name}' ===")
        return is_formula, clean_address, sheet_name

    def _fill_cell_or_range(self, workbook, sheet_name: str, address: str, value):
        """Заполняет ячейку или диапазон значением"""
        print(f"=== DEBUG: Заполняем лист='{sheet_name}', адрес='{address}', значение='{value}' ===")
        
        try:
            # Определяем активный лист
            if sheet_name:
                if sheet_name in workbook.sheetnames:
                    worksheet = workbook[sheet_name]
                    print(f"=== DEBUG: Используем лист '{sheet_name}' ===")
                else:
                    print(f"=== DEBUG: Лист '{sheet_name}' не найден, используем первый лист ===")
                    worksheet = workbook.worksheets[0]
            else:
                worksheet = workbook.worksheets[0]
                print(f"=== DEBUG: Используем первый лист ===")
            
            # Проверяем, является ли адрес стандартным (A1, B2, etc.)
            import re
            if re.match(r'^[A-Z]+[0-9]+$', address):
                # Стандартный адрес ячейки
                worksheet[address] = value
                print(f"=== DEBUG: Заполнили стандартную ячейку {address} значением '{value}' ===")
            elif ':' in address:
                # Это диапазон - заполняем первую ячейку
                start_cell = address.split(':')[0]
                if re.match(r'^[A-Z]+[0-9]+$', start_cell):
                    worksheet[start_cell] = value
                    print(f"=== DEBUG: Заполнили первую ячейку диапазона {start_cell} значением '{value}' ===")
                else:
                    print(f"=== DEBUG: Неверный формат диапазона {address} ===")
            else:
                # Возможно именованный диапазон или нестандартный адрес
                print(f"=== DEBUG: Попытка заполнить именованный диапазон или нестандартный адрес '{address}' ===")
                try:
                    # Попробуем найти именованный диапазон
                    if address in workbook.defined_names:
                        print(f"=== DEBUG: Найден именованный диапазон '{address}' ===")
                        # Для именованного диапазона получаем его адрес
                        defn = workbook.defined_names[address]
                        if defn.attr_text:
                            # Парсим адрес именованного диапазона
                            range_text = defn.attr_text
                            print(f"=== DEBUG: Адрес именованного диапазона: '{range_text}' ===")
                            
                            # Простой парсинг для Sheet1!$A$1 формата
                            if '!' in range_text:
                                sheet_part, cell_part = range_text.split('!', 1)
                                cell_part = cell_part.replace('$', '')  # Убираем символы $
                                if re.match(r'^[A-Z]+[0-9]+$', cell_part):
                                    worksheet[cell_part] = value
                                    print(f"=== DEBUG: Заполнили именованный диапазон {cell_part} значением '{value}' ===")
                                else:
                                    print(f"=== DEBUG: Неверный формат ячейки в именованном диапазоне: '{cell_part}' ===")
                            else:
                                print(f"=== DEBUG: Неверный формат именованного диапазона: '{range_text}' ===")
                    else:
                        print(f"=== DEBUG: Именованный диапазон '{address}' не найден ===")
                        # Попытаемся обработать как обычную ячейку
                        worksheet[address] = value
                        print(f"=== DEBUG: Заполнили как обычную ячейку '{address}' значением '{value}' ===")
                except Exception as e2:
                    print(f"=== DEBUG: Ошибка при работе с именованным диапазоном: {e2} ===")
                    raise e2
                    
        except Exception as e:
            print(f"=== DEBUG: Ошибка при заполнении {address}: {e} ===")
            # Попробуем через координаты
            try:
                from openpyxl.utils import coordinate_to_tuple
                if re.match(r'^[A-Z]+[0-9]+$', address):
                    row, col = coordinate_to_tuple(address)
                    worksheet = workbook.worksheets[0]
                    worksheet.cell(row=row, column=col, value=value)
                    print(f"=== DEBUG: Заполнили через координаты row={row}, col={col} значением '{value}' ===")
                else:
                    print(f"=== DEBUG: Не могу преобразовать адрес '{address}' в координаты ===")
                    raise e
            except Exception as e2:
                print(f"=== DEBUG: Ошибка при заполнении через координаты: {e2} ===")
                raise e2
        
    def clear_cache(self):
        """Очищает кэш rules"""
        self._cached_rules.clear()
        self._cached_templates.clear()
        print("=== DEBUG EXCEL: Кэш очищен ===")











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

    def read_vacation_info_from_file(self, file_path: str) -> Optional[VacationInfo]:
        """Читает информацию об отпусках из файла сотрудника"""
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            
            # Получаем структуру из конфига
            file_structure = self.config.employee_file_structure
            sheet_index = file_structure.get("active_sheet_index", 0)
            worksheet = workbook.worksheets[sheet_index]
            
            employee = Employee()
            # Удаляем/комментируем присваивания для Employee-объектов, так как теперь используется dict
            # employee.file_path = file_path
            # employee.tab_number = str(tab_number).strip()
            # employee.full_name = str(full_name).strip()
            # employee.position = str(position).strip()
            # employee.department1 = str(self._get_cell_value(worksheet, dept_cells.get("department1", "C2")) or "").strip()
            # employee.department2 = str(self._get_cell_value(worksheet, dept_cells.get("department2", "C3")) or "").strip()
            # employee.department3 = str(self._get_cell_value(worksheet, dept_cells.get("department3", "C4")) or "").strip()
            # employee.department4 = str(self._get_cell_value(worksheet, dept_cells.get("department4", "C5")) or "").strip()
            
            # Читаем периоды отпусков
            vacation_columns = file_structure.get("vacation_columns", {})
            periods = []
            
            for row in range(data_rows["start"], data_rows["end"] + 1):
                start_date_value = self._get_cell_value(worksheet, f"{vacation_columns.get('start_date', 'E')}{row}")
                end_date_value = self._get_cell_value(worksheet, f"{vacation_columns.get('end_date', 'F')}{row}")
                days_value = self._get_cell_value(worksheet, f"{vacation_columns.get('days', 'G')}{row}")
                
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
            
            # Читаем статус из B12 (новая логика)
            status_cell = file_structure.get("status_cell", "B12")
            status_value = self._get_cell_value(worksheet, status_cell)
            status_text = str(status_value).strip() if status_value else ""
            
            # Определяем статус на основе значения в B12
            statuses = self.config.validation_statuses
            if status_text == statuses["not_filled"]:
                vacation_status = VacationStatus.NOT_FILLED
            elif status_text == statuses["filled_correct"]:
                vacation_status = VacationStatus.FILLED_CORRECT
            elif status_text == statuses["filled_incorrect"]:
                vacation_status = VacationStatus.FILLED_INCORRECT
            else:
                # Если статус не распознан, пытаемся определить по содержимому
                if not periods:
                    vacation_status = VacationStatus.NOT_FILLED
                elif "некорректно" in status_text.lower() or "ошибка" in status_text.lower():
                    vacation_status = VacationStatus.FILLED_INCORRECT
                else:
                    vacation_status = VacationStatus.FILLED_INCORRECT  # По умолчанию
            
            vacation_info = VacationInfo(employee=employee, periods=periods, status=vacation_status)
            
            # Для отладки сохраняем текст статуса
            if status_text and vacation_status != VacationStatus.FILLED_CORRECT:
                vacation_info.validation_errors = [status_text]
            
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
        
        # Новая логика подсчета на основе 3 статусов
        not_filled_count = 0      # "Форма не заполнена"
        filled_incorrect_count = 0  # "Форма заполнена некорректно"
        filled_correct_count = 0    # "Форма заполнена корректно"
        
        statuses = self.config.validation_statuses
        
        for vi in vacation_infos:
            if vi.status == VacationStatus.NOT_FILLED:
                not_filled_count += 1
            elif vi.status == VacationStatus.FILLED_INCORRECT:
                filled_incorrect_count += 1
            elif vi.status == VacationStatus.FILLED_CORRECT:
                filled_correct_count += 1
        
        # Для совместимости с шаблонами считаем:
        employees_filled = filled_incorrect_count + filled_correct_count  # Все кроме "не заполнена"
        employees_correct = filled_correct_count  # Только корректно заполненные
        
        # Данные для заполнения
        report_data = {
            'block_name': block_name,
            'update_date': current_time.strftime('%d.%m.%Y %H:%M'),
            'total_employees': str(total_employees),
            'employees_filled': str(employees_filled),
            'employees_correct': str(employees_correct),
        }
        
        # Применяем rules
        self._apply_rules_to_template(workbook, rules, report_data)
        
        # Заполняем таблицы данных
        self._fill_employee_tables(workbook, vacation_infos, rules)
        
        # Заполняем календарь на Report листе
        if 'Report' in workbook.sheetnames:
            self._fill_calendar_matrix(workbook['Report'], vacation_infos)

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
        
        # Получаем структуру из конфига
        report_structure = self.config.report_structure
        base_row = report_structure.get("employee_data_start_row", 9)
        
        # Определяем mapping столбцов
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
        
        # Определяем текст статуса для отображения
        if vacation_info.status == VacationStatus.FILLED_CORRECT:
            status_text = "Ок"
        elif vacation_info.status == VacationStatus.NOT_FILLED:
            status_text = "Не заполнено"
        else:  # FILLED_INCORRECT
            # Показываем ошибки если есть, иначе общий текст
            status_text = "\n".join(vacation_info.validation_errors) if vacation_info.validation_errors else "Заполнено с ошибками"
        
        return {
            'row_number': index + 1,
            'employee_name': emp['ФИО работника'],
            'tab_number': emp['Табельный номер'],
            'position': emp['Должность'],
            'department1': emp['Подразделение 1'],
            'department2': emp['Подразделение 2'],
            'department3': emp['Подразделение 3'],
            'department4': emp['Подразделение 4'],
            'status': status_text,
            'total_days': vacation_info.total_days,
            'periods_count': vacation_info.periods_count
        }

    def _get_print_row_data(self, normalized_record: Dict, index: int) -> Dict[str, any]:
        """Возвращает данные строки для Print листа"""
        emp = normalized_record['employee']
        return {
            'row_number': index + 1,
            'tab_number': emp['Табельный номер'],
            'employee_name': emp['ФИО работника'],
            'position': emp['Должность'],
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
        
        report_structure = self.config.report_structure
        start_row = report_structure.get("employee_data_start_row", 9)
        
        for row in range(start_row, start_row + data_count):
            for col in range(1, 11):  # A-J
                worksheet.cell(row=row, column=col).border = thin_border

    def _fill_calendar_matrix(self, worksheet, vacation_infos: List[VacationInfo]):
        """Заполняет календарную матрицу"""
        # Получаем параметры из конфига
        report_structure = self.config.report_structure
        start_col = report_structure.get("calendar_start_col", 12)
        month_row = report_structure.get("calendar_month_row", 7)
        day_row = report_structure.get("calendar_day_row", 8)
        employee_start_row = report_structure.get("employee_data_start_row", 9)
        
        # Получаем данные календаря из конфига
        month_names = self.config.month_names
        days_in_months = self.config.days_in_months
        target_year = self.config.target_year
        
        col_offset = 0
        for month_idx, month_name in enumerate(month_names):
            month_col = start_col + col_offset
            worksheet.cell(row=month_row, column=month_col, value=month_name)
            
            days_in_month = days_in_months[month_idx]
            for day in range(1, days_in_month + 1):
                day_col = start_col + col_offset + day - 1
                worksheet.cell(row=day_row, column=day_col, value=day)
            
            col_offset += days_in_month
        
        for emp_idx, vacation_info in enumerate(vacation_infos):
            emp_row = employee_start_row + emp_idx
            for period in vacation_info.periods:
                current_date = period.start_date
                while current_date <= period.end_date:
                    if current_date.year == target_year:
                        day_col = self._get_calendar_column(current_date, start_col)
                        if day_col:
                            worksheet.cell(row=emp_row, column=day_col, value=1)
                    
                    from datetime import timedelta
                    current_date = current_date + timedelta(days=1)
                    if current_date > period.end_date:
                        break

    def _get_calendar_column(self, target_date: date, start_col: int) -> Optional[int]:
        """Вычисляет столбец для даты в календаре"""
        target_year = self.config.target_year
        days_in_months = self.config.days_in_months
        
        if target_date.year != target_year:
            return None
        
        col_offset = sum(days_in_months[:target_date.month - 1]) + target_date.day - 1
        return start_col + col_offset

    def read_block_report_data(self, report_path: str) -> Optional[Dict]:
        """Читает данные из отчета по блоку"""
        try:
            workbook = openpyxl.load_workbook(report_path, data_only=True)
            if 'Report' not in workbook.sheetnames:
                return None
            
            worksheet = workbook['Report']
            
            # Используем структуру из конфига
            report_structure = self.config.report_structure
            header_cells = report_structure.get("report_header_cells", {})
            
            block_name = str(self._get_cell_value(worksheet, header_cells.get("block_name", "A3")) or "").strip()
            update_date_raw = str(self._get_cell_value(worksheet, header_cells.get("update_date", "A4")) or "").strip()
            total_employees_raw = str(self._get_cell_value(worksheet, header_cells.get("total_employees", "A5")) or "").strip()
            completed_raw = str(self._get_cell_value(worksheet, header_cells.get("completed", "A6")) or "").strip()
            
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
        clean_fio = self._clean_filename(employee['ФИО работника'])
        clean_tab_num = self._clean_filename(employee['Табельный номер'])
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

    def test_created_employee_file(self, file_path: str, rules: Dict[str, Dict[str, str]]):
        """Проверяет значения в созданном файле по адресам из rules['value'] и выводит их в консоль"""
        import openpyxl
        print(f"=== TEST: Проверка заполнения файла: {file_path} ===")
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        for cell_address, field_name in rules.get('value', {}).items():
            try:
                is_formula, clean_address, sheet_name = self._parse_cell_address(cell_address)
                if sheet_name and sheet_name in workbook.sheetnames:
                    ws = workbook[sheet_name]
                else:
                    ws = workbook.worksheets[0]
                value = ws[clean_address].value if clean_address in ws else None
                print(f"  {field_name} ({cell_address}): {value}")
            except Exception as e:
                print(f"  {field_name} ({cell_address}): ERROR: {e}")