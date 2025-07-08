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
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from models import Employee, VacationInfo, VacationPeriod, VacationStatus, BlockReport
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
    
    def create_employee_file(self, employee: Employee, output_path: str) -> bool:
        """
        Создает файл сотрудника на основе шаблона
        
        Args:
            employee: данные сотрудника
            output_path: путь для сохранения файла
            
        Returns:
            bool: успешность операции
        """
        try:
            # Копируем шаблон
            template_path = Path(self.config.employee_template)
            if not template_path.exists():
                self.logger.error(f"Шаблон сотрудника не найден: {template_path}")
                return False
            
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            shutil.copy2(template_path, output_path)
            
            # Открываем скопированный файл для заполнения
            workbook = openpyxl.load_workbook(output_path)
            worksheet = workbook.active
            
            # Заполняем основные данные сотрудника
            self._fill_employee_data(worksheet, employee)
            
            # Заполняем строки планирования отпусков (строки 9-23)
            self._fill_vacation_rows(worksheet, employee)
            
            # Сохраняем файл
            workbook.save(output_path)
            workbook.close()
            
            self.logger.debug(f"Создан файл сотрудника: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка создания файла сотрудника {employee.full_name}: {e}")
            return False
    
    def read_vacation_info_from_file(self, file_path: str) -> Optional[VacationInfo]:
        """
        Читает информацию об отпусках из файла сотрудника
        
        Args:
            file_path: путь к файлу сотрудника
            
        Returns:
            VacationInfo или None при ошибке
        """
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            # Читаем базовую информацию о сотруднике из строк 9-23
            employee = Employee()
            
            # Ищем первую заполненную строку для получения базовой информации
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
            
            # Читаем подразделения из шапки файла (C2:C5)
            employee.department1 = str(self._get_cell_value(worksheet, "C2") or "").strip()
            employee.department2 = str(self._get_cell_value(worksheet, "C3") or "").strip()
            employee.department3 = str(self._get_cell_value(worksheet, "C4") or "").strip()
            employee.department4 = str(self._get_cell_value(worksheet, "C5") or "").strip()
            
            # Читаем периоды отпусков из строк 9-23
            periods = []
            
            for row in range(9, 24):
                start_date_value = self._get_cell_value(worksheet, f"E{row}")
                end_date_value = self._get_cell_value(worksheet, f"F{row}")
                days_value = self._get_cell_value(worksheet, f"G{row}")
                
                if not start_date_value or not end_date_value:
                    continue
                
                try:
                    # Парсим даты
                    start_date = self._parse_date(start_date_value)
                    end_date = self._parse_date(end_date_value)
                    
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
                    
                    period = VacationPeriod(
                        start_date=start_date,
                        end_date=end_date,
                        days=days
                    )
                    periods.append(period)
                    
                except Exception as e:
                    self.logger.warning(f"Ошибка обработки периода в строке {row}: {e}")
                    continue
            
            # Читаем результаты валидации
            validation_h2 = str(self._get_cell_value(worksheet, "H2") or "").strip()
            validation_i2 = str(self._get_cell_value(worksheet, "I2") or "").strip()
            validation_j2 = self._get_cell_value(worksheet, "J2") or 0
            
            # Создаем VacationInfo
            vacation_info = VacationInfo(
                employee=employee,
                periods=periods
            )
            
            # Определяем статус на основе валидаций
            vacation_info.validation_errors = []
            
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
                vacation_info.status = VacationStatus.OK
            else:
                vacation_info.status = VacationStatus.ERROR
            
            workbook.close()
            self.logger.debug(f"Прочитана информация об отпусках: {employee.full_name}")
            return vacation_info
            
        except Exception as e:
            self.logger.error(f"Ошибка чтения файла {file_path}: {e}")
            return None
    
    def create_block_report(self, block_name: str, vacation_infos: List[VacationInfo], output_path: str) -> bool:
        """
        Создает отчет по блоку с календарной матрицей
        
        Args:
            block_name: название блока
            vacation_infos: список сотрудников с информацией об отпусках
            output_path: путь для сохранения отчета
            
        Returns:
            bool: успешность операции
        """
        try:
            # Копируем шаблон отчета
            template_path = Path(self.config.block_report_template)
            if not template_path.exists():
                self.logger.error(f"Шаблон отчета по блоку не найден: {template_path}")
                return False
            
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            shutil.copy2(template_path, output_path)
            
            # Открываем файл для заполнения
            workbook = openpyxl.load_workbook(output_path)
            
            # Заполняем лист Report
            self._fill_report_sheet(workbook, block_name, vacation_infos)
            
            # Заполняем лист Print
            self._fill_print_sheet(workbook, block_name, vacation_infos)
            
            workbook.save(output_path)
            workbook.close()
            
            self.logger.info(f"Создан отчет по блоку: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка создания отчета по блоку {block_name}: {e}")
            return False
    
    def read_block_report_data(self, report_path: str) -> Optional[Dict]:
        """
        Читает данные из отчета по блоку для общего отчета
        
        Args:
            report_path: путь к отчету по блоку
            
        Returns:
            Dict с данными отчета или None при ошибке
        """
        try:
            workbook = openpyxl.load_workbook(report_path, data_only=True)
            
            # Ищем лист Report
            if 'Report' not in workbook.sheetnames:
                self.logger.error(f"Лист 'Report' не найден в файле {report_path}")
                return None
            
            worksheet = workbook['Report']
            
            # Читаем данные из шапки
            block_name = str(self._get_cell_value(worksheet, "A3") or "").strip()
            update_date_raw = str(self._get_cell_value(worksheet, "A4") or "").strip()
            total_employees_raw = str(self._get_cell_value(worksheet, "A5") or "").strip()
            completed_raw = str(self._get_cell_value(worksheet, "A6") or "").strip()
            
            # Парсим дату обновления: "Дата обновления: 07.07.2025 16:03"
            update_date = ""
            if "Дата обновления:" in update_date_raw:
                update_date = update_date_raw.replace("Дата обновления:", "").strip()
            
            # Парсим количество сотрудников: "Количество сотрудников: 4"
            total_employees = 0
            if "Количество сотрудников:" in total_employees_raw:
                try:
                    total_employees = int(total_employees_raw.split(":")[1].strip())
                except (ValueError, IndexError):
                    pass
            
            # Парсим завершивших планирование: "Закончили планирование: 3 (75%)"
            completed_employees = 0
            percentage = 0
            if "Закончили планирование:" in completed_raw:
                try:
                    # Извлекаем число перед скобкой
                    parts = completed_raw.split(":")[1].strip().split("(")
                    completed_employees = int(parts[0].strip())
                    
                    # Извлекаем процент из скобок
                    if len(parts) > 1:
                        percentage_str = parts[1].replace("%)", "").strip()
                        percentage = int(percentage_str)
                except (ValueError, IndexError):
                    pass
            
            workbook.close()
            
            remaining_employees = total_employees - completed_employees
            
            return {
                'block_name': block_name,
                'total_employees': total_employees,
                'completed_employees': completed_employees,
                'remaining_employees': remaining_employees,
                'percentage': percentage,
                'update_date': update_date
            }
            
        except Exception as e:
            self.logger.error(f"Ошибка чтения отчета по блоку {report_path}: {e}")
            return None
    
    def create_general_report_from_blocks(self, block_data: List[Dict], output_path: str) -> bool:
        """
        Создает общий отчет на основе данных из отчетов по блокам
        
        Args:
            block_data: список данных по блокам
            output_path: путь для сохранения отчета
            
        Returns:
            bool: успешность операции
        """
        try:
            # Копируем шаблон общего отчета
            template_path = Path(self.config.general_report_template)
            if not template_path.exists():
                self.logger.error(f"Шаблон общего отчета не найден: {template_path}")
                return False
            
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            shutil.copy2(template_path, output_path)
            
            # Открываем файл для заполнения
            workbook = openpyxl.load_workbook(output_path)
            worksheet = workbook.active
            
            # Заполняем первую строку данных (строка 6)
            if len(block_data) > 0:
                first_data = block_data[0]
                worksheet["A6"] = 1
                worksheet["B6"] = first_data['block_name']
                worksheet["C6"] = f"{first_data['percentage']}%"
                worksheet["D6"] = first_data['total_employees']
                worksheet["E6"] = first_data['completed_employees']
                worksheet["F6"] = first_data['remaining_employees']
                worksheet["G6"] = first_data['update_date']
            
            # Вставляем дополнительные строки ПОСЛЕ строки 6, если нужно
            if len(block_data) > 1:
                # Вставляем (количество блоков - 1) строк после строки 6
                worksheet.insert_rows(7, len(block_data) - 1)
                
                # Заполняем остальные строки данных
                for i in range(1, len(block_data)):
                    row = 6 + i
                    data = block_data[i]
                    
                    worksheet[f"A{row}"] = i + 1
                    worksheet[f"B{row}"] = data['block_name']
                    worksheet[f"C{row}"] = f"{data['percentage']}%"
                    worksheet[f"D{row}"] = data['total_employees']
                    worksheet[f"E{row}"] = data['completed_employees']
                    worksheet[f"F{row}"] = data['remaining_employees']
                    worksheet[f"G{row}"] = data['update_date']
            
            # Добавляем границы ко всем строкам данных
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for i in range(len(block_data)):
                row = 6 + i
                for col in range(1, 8):  # A-G
                    worksheet.cell(row=row, column=col).border = thin_border
            
            # Находим строку ИТОГО (она сдвинулась после вставки строк)
            total_row = 6 + len(block_data) + 1  # +1 для пустой строки
            
            # Обновляем формулы в строке ИТОГО
            if len(block_data) > 0:
                data_end_row = 6 + len(block_data) - 1
                worksheet[f"C{total_row}"] = f'=ROUND(E{total_row}/D{total_row}*100,0)&"%"'
                worksheet[f"D{total_row}"] = f'=SUM(D6:D{data_end_row})'
                worksheet[f"E{total_row}"] = f'=SUM(E6:E{data_end_row})'
                worksheet[f"F{total_row}"] = f'=SUM(F6:F{data_end_row})'
            
            workbook.save(output_path)
            workbook.close()
            
            self.logger.info(f"Создан общий отчет: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка создания общего отчета: {e}")
            return False
    
    def _fill_report_sheet(self, workbook, block_name: str, vacation_infos: List[VacationInfo]):
        """Заполняет лист Report"""
        if 'Report' not in workbook.sheetnames:
            self.logger.error("Лист 'Report' не найден в шаблоне")
            return
        
        worksheet = workbook['Report']
        current_time = datetime.now()
        
        # Шапка A3:A6
        worksheet["A3"] = block_name
        worksheet["A4"] = f"Дата обновления: {current_time.strftime('%d.%m.%Y %H:%M')}"
        worksheet["A5"] = f"Количество сотрудников: {len(vacation_infos)}"
        
        # Подсчет завершивших планирование
        completed = sum(1 for vi in vacation_infos if vi.status == VacationStatus.OK)
        percentage = (completed / len(vacation_infos) * 100) if vacation_infos else 0
        worksheet["A6"] = f"Закончили планирование: {completed} ({percentage:.0f}%)"
        
        # Заполняем таблицу сотрудников (начиная с строки 9)
        for i, vacation_info in enumerate(vacation_infos):
            row = i + 9
            emp = vacation_info.employee
            
            worksheet[f"A{row}"] = i + 1  # №
            worksheet[f"B{row}"] = emp.full_name  # ФИО
            worksheet[f"C{row}"] = emp.tab_number  # Таб. Номер
            worksheet[f"D{row}"] = getattr(emp, 'position', '')  # Должность
            worksheet[f"E{row}"] = emp.department1  # Подразделение 1
            worksheet[f"F{row}"] = emp.department2  # Подразделение 2
            worksheet[f"G{row}"] = emp.department3  # Подразделение 3
            worksheet[f"H{row}"] = emp.department4  # Подразделение 4
            
            # Статус планирования
            if vacation_info.status == VacationStatus.OK:
                worksheet[f"I{row}"] = "Ок"
            else:
                errors = getattr(vacation_info, 'validation_errors', [])
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
        self._fill_calendar_matrix(worksheet, vacation_infos)
    
    def _fill_calendar_matrix(self, worksheet, vacation_infos: List[VacationInfo]):
        """Заполняет календарную матрицу на листе Report"""
        try:
            # Начинаем календарь после таблицы сотрудников
            # Предполагаем, что календарь начинается с столбца L (12)
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
                            day_col = self._get_calendar_column(current_date, start_col)
                            if day_col:
                                worksheet.cell(row=emp_row, column=day_col, value=1)
                        
                        # Исправленная логика перехода к следующему дню
                        from datetime import timedelta
                        current_date = current_date + timedelta(days=1)
                        
                        if current_date > period.end_date:
                            break
                            
        except Exception as e:
            self.logger.error(f"Ошибка заполнения календарной матрицы: {e}")
    
    def _get_calendar_column(self, target_date: date, start_col: int) -> Optional[int]:
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
    
    def _fill_print_sheet(self, workbook, block_name: str, vacation_infos: List[VacationInfo]):
        """Заполняет лист Print в нормализованном виде"""
        if 'Print' not in workbook.sheetnames:
            self.logger.error("Лист 'Print' не найден в шаблоне")
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
                self._add_print_headers(worksheet, current_row)
                current_row += 1
                records_on_page = 0
                is_first_page = False
            
            # Заполняем строку данных
            emp = record['employee']
            
            worksheet[f"A{current_row}"] = record_idx + 1  # № п/п
            worksheet[f"B{current_row}"] = emp.tab_number  # Табельный номер
            worksheet[f"C{current_row}"] = emp.full_name  # ФИО
            worksheet[f"D{current_row}"] = getattr(emp, 'position', '')  # Должность
            
            if record['start_date']:
                worksheet[f"E{current_row}"] = record['start_date'].strftime('%d.%m.%Y')  # Дата начала
                worksheet[f"F{current_row}"] = record['end_date'].strftime('%d.%m.%Y')  # Дата окончания
                worksheet[f"G{current_row}"] = record['days']  # Продолжительность
            
            # Остальные столбцы пока пустые (Подпись, Дата ознакомления, Примечание)
            
            current_row += 1
            records_on_page += 1
    
    def _add_print_headers(self, worksheet, row: int):
        """Добавляет заголовки таблицы для печати"""
        headers = [
            "№ п/п", "Табельный номер", "ФИО", "Должность",
            "Дата начала отпуска", "Дата окончания отпуска",
            "Продолжительность (календарных дней)",
            "Подпись работника", "Дата ознакомления работника", "Примечание"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            worksheet.cell(row=row, column=col_idx, value=header)
    
    def _fill_employee_data(self, worksheet, employee: Employee):
        """Заполняет основные данные сотрудника в шаблоне"""
        # Основные данные в шапке формы
        worksheet["C2"] = employee.department1  # Подразделение 1
        worksheet["C3"] = employee.department2  # Подразделение 2
        worksheet["C4"] = employee.department3  # Подразделение 3
        
        if employee.department4:
            worksheet["C5"] = employee.department4
    
    def _fill_vacation_rows(self, worksheet, employee: Employee):
        """Заполняет строки планирования отпусков (9-23)"""
        # Заполняем строки 9-23 базовой информацией
        for row in range(9, 24):
            worksheet[f"B{row}"] = employee.tab_number  # Табельный номер
            worksheet[f"C{row}"] = employee.full_name   # ФИО
            worksheet[f"D{row}"] = getattr(employee, 'position', '')  # Должность
    
    def _get_cell_value(self, worksheet, cell_address: str):
        """Безопасно получает значение ячейки"""
        try:
            cell = worksheet[cell_address]
            return cell.value
        except Exception:
            return None
    
    def _parse_date(self, date_value) -> Optional[date]:
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
        formats = [
            "%d.%m.%Y",
            "%d.%m.%y", 
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d/%m/%y"
        ]
        
        for fmt in formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt).date()
                return parsed_date
            except ValueError:
                continue
        
        self.logger.warning(f"Не удалось распарсить дату: {date_value}")
        return None
    
    def generate_output_filename(self, employee: Employee) -> str:
        """Генерирует имя файла для сотрудника в формате ФИО (табНомер).xlsx"""
        clean_fio = self._clean_filename(employee.full_name)
        clean_tab_num = self._clean_filename(employee.tab_number)
        return f"{clean_fio} ({clean_tab_num}).xlsx"
    
    def generate_block_report_filename(self, block_name: str) -> str:
        """Генерирует имя файла отчета по блоку с временной меткой"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        clean_block_name = self._clean_filename(block_name)
        return f"Отчет по блоку_{clean_block_name}_{timestamp}.xlsx"
    
    def read_vacation_info_from_file(self, file_path: str) -> Optional[VacationInfo]:
            """
            Читает информацию об отпусках из файла сотрудника
            
            Args:
                file_path: путь к файлу сотрудника
                
            Returns:
                VacationInfo или None при ошибке
            """
            try:
                workbook = openpyxl.load_workbook(file_path, data_only=True)
                worksheet = workbook.active
                
                # Читаем базовую информацию о сотруднике из строк 9-23
                employee = Employee()
                
                # Ищем первую заполненную строку для получения базовой информации
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
                
                # Читаем подразделения из шапки файла (C2:C5)
                employee.department1 = str(self._get_cell_value(worksheet, "C2") or "").strip()
                employee.department2 = str(self._get_cell_value(worksheet, "C3") or "").strip()
                employee.department3 = str(self._get_cell_value(worksheet, "C4") or "").strip()
                employee.department4 = str(self._get_cell_value(worksheet, "C5") or "").strip()
                
                # Читаем периоды отпусков из строк 9-23
                periods = []
                
                for row in range(9, 24):
                    start_date_value = self._get_cell_value(worksheet, f"E{row}")
                    end_date_value = self._get_cell_value(worksheet, f"F{row}")
                    days_value = self._get_cell_value(worksheet, f"G{row}")
                    
                    if not start_date_value or not end_date_value:
                        continue
                    
                    try:
                        # Парсим даты
                        start_date = self._parse_date(start_date_value)
                        end_date = self._parse_date(end_date_value)
                        
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
                        
                        period = VacationPeriod(
                            start_date=start_date,
                            end_date=end_date,
                            days=days
                        )
                        periods.append(period)
                        
                    except Exception as e:
                        self.logger.warning(f"Ошибка обработки периода в строке {row}: {e}")
                        continue
                
                # Читаем результаты валидации
                validation_h2 = str(self._get_cell_value(worksheet, "H2") or "").strip()
                validation_i2 = str(self._get_cell_value(worksheet, "I2") or "").strip()
                validation_j2 = self._get_cell_value(worksheet, "J2") or 0
                
                # Создаем VacationInfo
                vacation_info = VacationInfo(
                    employee=employee,
                    periods=periods
                )
                
                # Определяем статус на основе валидаций
                vacation_info.validation_errors = []
                
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
                    vacation_info.status = VacationStatus.OK
                else:
                    vacation_info.status = VacationStatus.ERROR
                
                workbook.close()
                self.logger.debug(f"Прочитана информация об отпусках: {employee.full_name}")
                return vacation_info
                
            except Exception as e:
                self.logger.error(f"Ошибка чтения файла {file_path}: {e}")
                return None
    
    def _clean_filename(self, filename: str) -> str:
        """Очищает имя файла от недопустимых символов"""
        if not filename:
            return "unnamed"
        
        # Заменяем недопустимые символы
        invalid_chars = r'[\\/:*?"<>|]'
        clean_name = re.sub(invalid_chars, '_', filename)
        
        # Убираем лишние пробелы и ограничиваем длину
        clean_name = clean_name.strip()
        if len(clean_name) > 100:
            clean_name = clean_name[:100]
        
        return clean_name or "unnamed"