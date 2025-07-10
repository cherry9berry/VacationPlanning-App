#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Окно создания файлов сотрудников
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
from pathlib import Path
from typing import Optional, Dict
import re
from datetime import datetime


from config import Config
from core.processor import VacationProcessor
from models import ProcessingProgress, ProcessingStatus


class CreateFilesWindow:
    """Окно для создания файлов отпусков сотрудников"""
    

    def __init__(self, parent: tk.Tk, config: Config, main_window):
        """Конструктор окна создания файлов"""
        self.parent = parent
        self.config = config
        self.main_window = main_window
        self.logger = logging.getLogger(__name__)
        self.processor = VacationProcessor(config)
        
        # Состояние
        self.staff_file_path = ""
        self.output_dir_path = ""
        self.is_processing = False
        self.validation_result = None
        self.existing_files_info = {}
        self.new_employees_count = 0
        self.skip_employees_count = 0
        
        # НОВЫЕ ПЕРЕМЕННЫЕ для отслеживания повторных выборов
        self.file_reselected = False
        self.dir_reselected = False
        
        # Создаем окно
        self.window = None
        self.setup_ui()

    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Создание файлов отпусков")
        self.window.geometry("750x600")
        self.window.resizable(True, True)
        
        # Обработчик закрытия
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Основной фрейм
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка сетки
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Выбор файлов
        self.setup_file_selection(main_frame)
        
        # Информация/прогресс (занимает одно место)
        self.setup_info_progress_area(main_frame)
        
        # Кнопки управления (внизу по центру)
        self.setup_control_buttons(main_frame)
    
    def setup_file_selection(self, parent):
        """Настройка области выбора файлов"""
        files_frame = ttk.LabelFrame(parent, text="Выбор файлов", padding="10")
        files_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E))
        files_frame.columnconfigure(1, weight=1)
        
        # Файл с сотрудниками
        ttk.Label(files_frame, text="Файл с сотрудниками:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        self.staff_file_var = tk.StringVar()
        self.staff_file_entry = ttk.Entry(
            files_frame, 
            textvariable=self.staff_file_var,
            state="readonly"
        )
        self.staff_file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        ttk.Button(
            files_frame,
            text="Выбрать",
            command=self.select_staff_file
        ).grid(row=0, column=2, pady=5)
        
        # Целевая папка
        ttk.Label(files_frame, text="Целевая папка:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        
        self.output_dir_var = tk.StringVar()
        self.output_dir_entry = ttk.Entry(
            files_frame,
            textvariable=self.output_dir_var,
            state="readonly"
        )
        self.output_dir_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        self.output_dir_btn = ttk.Button(
            files_frame,
            text="Выбрать",
            command=self.select_output_dir,
            state=tk.DISABLED
        )
        self.output_dir_btn.grid(row=1, column=2, pady=5)
    

    def setup_info_progress_area(self, parent):
        """Настройка области информации/прогресса"""
        # Информация (показывается по умолчанию)
        self.info_frame = ttk.LabelFrame(parent, text="Информация", padding="10")
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E, tk.N, tk.S))
        self.info_frame.columnconfigure(0, weight=1)
        self.info_frame.rowconfigure(0, weight=1)
        
        # Текстовая область с прокруткой и возможностью копирования
        self.info_text = tk.Text(
            self.info_frame,
            height=12,
            wrap=tk.WORD,
            font=("TkDefaultFont", 9),
            state=tk.NORMAL,
            cursor="arrow"
        )
        
        # Делаем текст доступным для выделения и копирования
        def on_key(event):
            # Разрешаем только копирование
            if event.state & 0x4:  # Ctrl нажат
                if event.keysym.lower() in ['c', 'a']:
                    return  # Разрешаем Ctrl+C и Ctrl+A
            return "break"  # Блокируем все остальные клавиши
        
        self.info_text.bind('<Key>', on_key)
        self.info_text.bind('<Control-a>', lambda e: self.info_text.tag_add("sel", "1.0", "end"))
        
        # Контекстное меню для копирования
        def show_context_menu(event):
            try:
                context_menu = tk.Menu(self.info_text, tearoff=0)
                context_menu.add_command(label="Выделить всё", command=lambda: self.info_text.tag_add("sel", "1.0", "end"))
                context_menu.add_command(label="Копировать", command=lambda: self.copy_selected_text())
                context_menu.tk_popup(event.x_root, event.y_root)
            except:
                pass
        
        self.info_text.bind('<Button-3>', show_context_menu)
        
        info_scrollbar = ttk.Scrollbar(self.info_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scrollbar.set)
        
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Инициализация
        self.add_info("Выберите файл с сотрудниками для начала работы")
        
        # Прогресс обработки (скрыт по умолчанию)
        self.progress_frame = ttk.LabelFrame(parent, text="Прогресс обработки", padding="10")
        self.progress_frame.columnconfigure(0, weight=1)
        
        # Общий прогресс и время
        self.overall_progress_label = ttk.Label(self.progress_frame, text="Готов к началу", font=("TkDefaultFont", 10, "bold"))
        self.overall_progress_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.time_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 9))
        self.time_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # Прогресс по отделам (основной)
        self.departments_label = ttk.Label(self.progress_frame, text="Отделы:", font=("TkDefaultFont", 9, "bold"))
        self.departments_label.grid(row=2, column=0, sticky=tk.W, pady=(10, 2))
        
        self.departments_progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode='determinate',
            length=400
        )
        self.departments_progress_bar.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=2)
        
        self.departments_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.departments_detail_label.grid(row=4, column=0, sticky=tk.W, pady=2)
        
        # Прогресс по сотрудникам в текущем отделе (вторичный)
        self.employees_label = ttk.Label(self.progress_frame, text="Файлы в текущем отделе:", font=("TkDefaultFont", 9))
        self.employees_label.grid(row=5, column=0, sticky=tk.W, pady=(10, 2))
        
        self.employees_progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode='determinate',
            length=400
        )
        self.employees_progress_bar.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=2)
        
        self.employees_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.employees_detail_label.grid(row=7, column=0, sticky=tk.W, pady=2)

    
    def setup_control_buttons(self, parent):
        """Настройка кнопок управления"""
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=(10, 15))
        
        # Центрирование кнопок
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(2, weight=1)
        
        # Кнопка создания файлов
        self.create_btn = ttk.Button(
            buttons_frame,
            text="Создать файлы",
            command=self.create_files,
            state=tk.DISABLED
        )
        self.create_btn.grid(row=0, column=1)
    
    def show(self):
        """Показывает окно"""
        if self.window:
            self.window.deiconify()
            self.window.lift()
            self.window.focus()

    def select_staff_file(self):
        """Выбор файла с сотрудниками"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл с сотрудниками",
            filetypes=[
                ("Excel файлы", "*.xlsx *.xls"),
                ("Все файлы", "*.*")
            ]
        )
        
        if file_path:
            self.staff_file_path = file_path
            self.staff_file_var.set(file_path)
            
            # Полный сброс состояния при смене файла
            self._reset_state()
            
            self.add_info(f"Выбран файл: {Path(file_path).name}")
            
            # Активируем кнопку выбора папки
            self.output_dir_btn.config(state=tk.NORMAL)
            
            # ВАЖНО: сбрасываем кнопку создания в обычное состояние
            self.create_btn.config(state=tk.DISABLED, text="Создать файлы", command=self.create_files)
            
            # Автоматически запускаем валидацию
            self.validate_file()
        
        # Возвращаем фокус на окно создания файлов
        self.window.lift()
        self.window.focus_force()

    def _reset_state(self):
        """Полный сброс состояния окна"""
        self.output_dir_path = ""
        self.output_dir_var.set("")
        self.new_employees_count = 0
        self.skip_employees_count = 0
        self.validation_result = None
        self.existing_files_info = {}
        if hasattr(self, '_employees'):
            delattr(self, '_employees')

    def select_output_dir(self):
        """Выбор целевой папки"""
        dir_path = filedialog.askdirectory(
            title="Выберите целевую папку для создания файлов сотрудников"
        )
        
        if dir_path:
            self.output_dir_path = dir_path
            self.output_dir_var.set(dir_path)
            self.add_info(f"Выбрана целевая папка: {dir_path}")
            
            # ВАЖНО: сбрасываем кнопку создания в обычное состояние
            self.create_btn.config(state=tk.DISABLED, text="Создать файлы", command=self.create_files)
            
            # Проверяем существующие файлы
            self.add_info("")
            self.add_info("Анализ целевой папки...")
            self.check_existing_files(dir_path)
            
            # Проверяем возможность активации кнопки создания
            self.check_create_button_state()
        
        # Возвращаем фокус на окно создания файлов
        self.window.lift()
        self.window.focus_force()

    def check_existing_files(self, dir_path):
        """Проверяет существующие файлы в папке и выводит подробную статистику"""
        try:
            base_path = Path(dir_path)
            if not base_path.exists():
                self.add_info("Папка не существует")
                return
            
            # Проверяем что данные сотрудников загружены
            if not (self.validation_result and hasattr(self, '_employees') and self._employees):
                self.add_info("Ошибка: данные сотрудников не загружены")
                return
            
            # Получаем ожидаемые отделы из валидации
            expected_departments = set()
            for emp in self._employees:
                if emp.department1:
                    clean_dept = self._clean_directory_name(emp.department1)
                    expected_departments.add((emp.department1, clean_dept))
            
            existing_departments = []
            new_departments = []
            total_existing_employees = 0
            existing_employees_by_dept = {}
            departments_with_files = 0
            
            # Сканируем папки
            for item in base_path.iterdir():
                if item.is_dir():
                    # Проверяем, является ли папка ожидаемым отделом
                    for orig_dept, clean_dept in expected_departments:
                        if item.name == clean_dept:
                            existing_departments.append(orig_dept)
                            
                            # Подсчитываем файлы в этой папке
                            dept_files = 0
                            for file_item in item.iterdir():
                                if file_item.is_file() and file_item.suffix.lower() == '.xlsx':
                                    # Исключаем отчеты
                                    if not file_item.name.startswith('!'):
                                        dept_files += 1
                            
                            existing_employees_by_dept[orig_dept] = dept_files
                            total_existing_employees += dept_files
                            if dept_files > 0:
                                departments_with_files += 1
                            break
            
            # Определяем новые отделы
            all_departments_from_file = set(emp.department1 for emp in self._employees if emp.department1)
            new_departments = list(all_departments_from_file - set(existing_departments))
            
            # Подсчитываем новых сотрудников
            self.new_employees_count = max(0, len(self._employees) - total_existing_employees)
            self.skip_employees_count = total_existing_employees
            
            # Выводим информацию в новом формате
            if existing_departments or new_departments:
                self.add_info("Файлы по сотрудникам не будут переписаны, но \"долив\" в существующие папки сотрудников будет выполнен", "warning")
                
                # Списки отделов в одну строку через запятую
                if existing_departments:
                    existing_list = ", ".join(existing_departments)
                    self.add_info(f"Существующие отделы: {existing_list}")
                
                if new_departments:
                    new_list = ", ".join(new_departments)
                    self.add_info(f"Новые отделы: {new_list}")
                
                if self.new_employees_count > 0:
                    self.add_info(f"Будет добавлено {self.new_employees_count} новых сотр., остальные {self.skip_employees_count} сотр. уже есть")
                    if self.skip_employees_count > 0:
                        self.add_info("Если требуется их переписать - удалите старые файлы вручную")
                    
                    # Статистика
                    self.add_info("")
                    self.add_info("ПЛАН ОБРАБОТКИ:", "success")
                    self.add_info(f"  • Будет создано: {self.new_employees_count} сотр.")
                    self.add_info(f"  • Будет пропущено: {self.skip_employees_count} сотр.")
                    self.add_info(f"  • Всего сотрудников: {len(self._employees)}")
                    
                    # Рассчитываем ожидаемое время
                    if self.validation_result:
                        estimated_time = max(0.1, self.new_employees_count * self.config.get("processing_time_per_file", 0.3))
                        self.add_info(f"  • Ожидаемое время: {estimated_time:.1f} сек")
                    
                    # Следующие шаги
                    self.add_info("")
                    self.add_info("Для продолжения нажмите кнопку 'Создать файлы'", "success")
                else:
                    self.add_info("Все записи уже есть в папке - новых файлов создаваться не будет")
                    self.new_employees_count = 0
            else:
                # Новая структура
                self.new_employees_count = len(self._employees)
                self.skip_employees_count = 0
                
                # Все отделы новые
                all_departments = list(set(emp.department1 for emp in self._employees if emp.department1))
                new_list = ", ".join(all_departments)
                
                self.add_info("В папке нет подразделений из вашего файла - будет создана новая структура")
                self.add_info(f"Новые отделы: {new_list}")
                
                self.add_info("")
                self.add_info("ПЛАН ОБРАБОТКИ:", "success")
                self.add_info(f"  • Будет создано: {self.new_employees_count} сотр.")
                self.add_info(f"  • Будет пропущено: {self.skip_employees_count} сотр.")
                self.add_info(f"  • Всего сотрудников: {len(self._employees)}")
                
                # Рассчитываем ожидаемое время для всех файлов
                estimated_time = max(0.1, self.new_employees_count * self.config.get("processing_time_per_file", 0.3))
                self.add_info(f"  • Ожидаемое время: {estimated_time:.1f} сек")
                
                # Следующие шаги
                self.add_info("")
                self.add_info("Для продолжения нажмите кнопку 'Создать файлы'", "success")
            
            # Сохраняем информацию для использования в процессоре
            self.existing_files_info = {
                'departments': existing_departments,
                'files_count': total_existing_employees,
                'by_department': existing_employees_by_dept
            }
                
        except Exception as e:
            self.add_info(f"Ошибка проверки папки: {e}", "error")
    
    def validate_file(self):
        """Валидация выбранного файла"""
        if not self.staff_file_path:
            return
        
        self.add_info("Начало валидации файла...")
        
        # Запускаем валидацию в отдельном потоке
        def validate_thread():
            try:
                validation_result, employees = self.processor.validator.validate_staff_file(self.staff_file_path)
                
                # Сохраняем результат для использования
                self.validation_result = validation_result
                
                # Обновляем UI в главном потоке
                self.window.after(0, self.on_validation_complete, employees)
                
            except Exception as e:
                self.window.after(0, self.on_validation_error, str(e))
        
        threading.Thread(target=validate_thread, daemon=True).start()
    
    def on_validation_complete(self, employees):
        """Обработчик завершения валидации"""
        # Сохраняем сотрудников для проверки отделов и уникальности
        self._employees = employees
        
        if self.validation_result.is_valid:
            self.add_info("")
            self.add_info("ВАЛИДАЦИЯ УСПЕШНО ЗАВЕРШЕНА", "success")
            
            # Проверяем уникальность табельных номеров и отделов
            unique_check = self.check_employee_uniqueness(employees)
            if not unique_check['is_valid']:
                self.add_info("")
                self.add_info("НАЙДЕНЫ ПРОБЛЕМЫ:", "warning")
                for error in unique_check['errors']:
                    self.add_info(f"  • {error}", "warning")
            
            self.add_info("")
            self.add_info("СТАТИСТИКА ФАЙЛА:", "success")
            stats_lines = self.format_validation_stats(self.validation_result, employees).split('\n')
            for line in stats_lines:
                if line.strip():
                    self.add_info(f"  {line.strip()}")
            
            # Если уже выбрана папка, перепроверяем ее
            if self.output_dir_path:
                self.add_info("")
                self.add_info("Анализ целевой папки...")
                self.check_existing_files(self.output_dir_path)
            else:
                self.add_info("")
                self.add_info("Выберите целевую папку для создания файлов сотрудников")
            
            # Проверяем возможность активации кнопки создания
            self.check_create_button_state()
            
        else:
            self.add_info("")
            self.add_info("ВАЛИДАЦИЯ ВЫЯВИЛА ОШИБКИ", "error")
            
            self.add_info("")
            self.add_info("ОШИБКИ ВАЛИДАЦИИ:", "error")
            for error in self.validation_result.errors:
                self.add_info(f"  • {error}", "error")
            
            if self.validation_result.warnings:
                self.add_info("")
                self.add_info("ПРЕДУПРЕЖДЕНИЯ:", "warning")
                for warning in self.validation_result.warnings:
                    self.add_info(f"  • {warning}", "warning")
    
    def check_employee_uniqueness(self, employees):
        """Проверяет уникальность сотрудников по табельному номеру и отделам"""
        result = {'is_valid': True, 'errors': []}
        
        # Проверяем уникальность табельных номеров
        tab_numbers = {}
        for emp in employees:
            if emp.tab_number in tab_numbers:
                tab_numbers[emp.tab_number].append(emp.full_name)
            else:
                tab_numbers[emp.tab_number] = [emp.full_name]
        
        for tab_num, names in tab_numbers.items():
            if len(names) > 1:
                result['is_valid'] = False
                result['errors'].append(f"Дублирующийся табельный номер {tab_num}: {', '.join(names)}")
        
        # Проверяем, что один сотрудник не находится в разных отделах
        employee_departments = {}
        for emp in employees:
            key = f"{emp.full_name}_{emp.tab_number}"
            if key in employee_departments:
                if employee_departments[key] != emp.department1:
                    result['is_valid'] = False
                    result['errors'].append(f"Сотрудник {emp.full_name} ({emp.tab_number}) находится в разных подразделениях: {employee_departments[key]} и {emp.department1}")
            else:
                employee_departments[key] = emp.department1
        
        return result
    
    def on_validation_error(self, error_message):
        """Обработчик ошибки валидации"""
        self.add_info(f"Ошибка валидации: {error_message}", "error")
        messagebox.showerror("Ошибка валидации", error_message)
    
    def format_validation_stats(self, validation_result, employees):
        """Форматирует статистику валидации"""
        stats = f"СТАТИСТИКА ФАЙЛА:\n"
        stats += f"• Всего сотрудников: {validation_result.employee_count}\n"
        stats += f"• Уникальных табельных номеров: {validation_result.unique_tab_numbers}\n"
        
        if validation_result.warnings:
            stats += f"• Предупреждений: {len(validation_result.warnings)}"
        
        return stats
    
    def _clean_directory_name(self, name: str) -> str:
        """Очищает имя папки от недопустимых символов"""
        if not name:
            return "unnamed"
        
        invalid_chars = r'[<>:"/\\|?*]'
        clean_name = re.sub(invalid_chars, '_', name)
        clean_name = clean_name.strip('. ')
        
        if len(clean_name) > 100:
            clean_name = clean_name[:100]
        
        return clean_name or "unnamed"

    def check_create_button_state(self):
        """Проверяет и обновляет состояние кнопки создания файлов"""
        # Кнопка активна только если:
        # 1. Валидация прошла успешно
        # 2. Выбрана целевая папка
        # 3. Есть сотрудники для обработки (new_employees_count > 0)
        # 4. Не идет процесс обработки
        # 5. Кнопка еще не в состоянии "Закрыть"
        if (self.validation_result and 
            self.validation_result.is_valid and 
            self.output_dir_path and 
            hasattr(self, 'new_employees_count') and
            self.new_employees_count > 0 and
            not self.is_processing and
            self.create_btn['text'] != "Закрыть"):
            
            self.create_btn.config(state=tk.NORMAL, text="Создать файлы", command=self.create_files)
        elif self.create_btn['text'] != "Закрыть":
            self.create_btn.config(state=tk.DISABLED)
    
    def create_files(self):
        """Создание файлов сотрудников"""
        if not self.validation_result or not self.validation_result.is_valid:
            messagebox.showwarning("Предупреждение", "Сначала выполните валидацию файла")
            return
        
        if not self.output_dir_path:
            messagebox.showwarning("Предупреждение", "Выберите целевую папку")
            return
        
        if not hasattr(self, 'new_employees_count') or self.new_employees_count <= 0:
            messagebox.showwarning("Предупреждение", "Нет новых записей для создания")
            return
        
        self.start_processing()
    
    def start_processing(self):
        """Начинает процесс создания файлов"""
        self.is_processing = True
        self.create_btn.config(state=tk.DISABLED)
        
        # Переключаемся на отображение прогресса
        self.show_progress_view()
        
        self.add_info("Начало создания файлов...")
        
        def processing_thread():
            try:
                # Используем исправленный метод процессора
                operation_log = self.processor.create_employee_files_to_existing(
                    self.staff_file_path,
                    self.output_dir_path,
                    self.on_progress_update,
                    self.on_department_progress_update,
                    self.on_file_progress_update
                )
                
                # Завершение в главном потоке
                self.window.after(0, self.on_processing_complete, operation_log)
                
            except Exception as e:
                self.window.after(0, self.on_processing_error, str(e))
        
        threading.Thread(target=processing_thread, daemon=True).start()
    
    def show_progress_view(self):
        """Показывает область прогресса вместо информации"""
        self.info_frame.grid_remove()
        self.progress_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def show_info_view(self):
        """Показывает область информации вместо прогресса"""
        self.progress_frame.grid_remove()
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def on_progress_update(self, progress):
        """Обработчик обновления общего прогресса"""
        def update_ui():
            # Общий процент
            if progress.total_files > 0:
                overall_percent = (progress.processed_files / progress.total_files) * 100
                self.overall_progress_label.config(text=f"Общий прогресс: {overall_percent:.1f}%")
            
            # Время
            elapsed = (datetime.now() - progress.start_time).total_seconds()
            if progress.processed_files > 0 and progress.total_files > 0:
                speed = progress.processed_files / elapsed  # файлов в секунду
                remaining_files = progress.total_files - progress.processed_files
                remaining_time = remaining_files / speed if speed > 0 else 0
                self.time_label.config(
                    text=f"Прошло: {elapsed:.0f} сек, Осталось: {remaining_time:.0f} сек"
                )
            else:
                self.time_label.config(text=f"Прошло: {elapsed:.0f} сек")
        
        self.window.after(0, update_ui)
    
    def on_department_progress_update(self, current_dept, total_depts, dept_name):
        """Обработчик обновления прогресса по отделам"""
        def update_ui():
            if total_depts > 0:
                dept_percent = (current_dept / total_depts) * 100
                self.departments_progress_bar['value'] = dept_percent
                self.departments_detail_label.config(
                    text=f"Отдел {current_dept}/{total_depts}: {dept_name}"
                )
        
        self.window.after(0, update_ui)
    
    def on_file_progress_update(self, current_file, total_files, file_info):
        """Обработчик обновления прогресса по файлам в текущем отделе"""
        def update_ui():
            if total_files > 0:
                file_percent = (current_file / total_files) * 100
                self.employees_progress_bar['value'] = file_percent
                self.employees_detail_label.config(
                    text=f"Файл {current_file}/{total_files}: {file_info}"
                )
        
        self.window.after(0, update_ui)
    
    def on_processing_complete(self, operation_log):
        """Обработчик завершения создания файлов"""
        self.is_processing = False
        
        if operation_log.status == ProcessingStatus.SUCCESS:
            # Добавляем результат в существующую информацию
            self.add_info_to_existing("")
            self.add_info_to_existing("=" * 50)
            self.add_info_to_existing("СОЗДАНИЕ ФАЙЛОВ УСПЕШНО ЗАВЕРШЕНО!", "success")
            self.add_info_to_existing(f"Время выполнения: {operation_log.duration:.1f} сек")
            self.add_info_to_existing("=" * 50)
            self.add_info_to_existing("")
            
            # Добавляем информацию о результатах операции
            for entry in operation_log.entries:
                if entry.level == "INFO":
                    if "Создано:" in entry.message or "создано" in entry.message.lower():
                        message = entry.message
                        if "из" in message and "сотрудников" in message:
                            message = message.replace("сотрудников", "сотр.")
                        self.add_info_to_existing(f"ИТОГ: {message}", "success")
                    else:
                        self.add_info_to_existing(f"  • {entry.message}")
            
            # Возвращаемся к отображению информации
            self.show_info_view()
            
            # ИСПРАВЛЕНИЕ: Всегда показываем кнопку "Закрыть"
            self.create_btn.config(text="Закрыть", command=self.on_closing, state=tk.NORMAL)
            
        else:
            self.add_info_to_existing("")
            self.add_info_to_existing("СОЗДАНИЕ ФАЙЛОВ ЗАВЕРШЕНО С ОШИБКАМИ!", "error")
            
            # Показываем ошибки
            for entry in operation_log.entries:
                if entry.level == "ERROR":
                    self.add_info_to_existing(f"ОШИБКА: {entry.message}", "error")
            
            # Возвращаемся к отображению информации
            self.show_info_view()
            self.create_btn.config(text="Закрыть", command=self.on_closing, state=tk.NORMAL)
            
            # Показываем messagebox с ошибкой
            messagebox.showerror("Ошибка создания файлов", "Создание файлов завершено с ошибками. См. подробности в окне.")


    def on_processing_error(self, error_message):
        """Обработчик ошибки создания файлов"""
        self.is_processing = False
        self.create_btn.config(text="Закрыть", command=self.on_closing, state=tk.NORMAL)
        
        self.add_info(f"Критическая ошибка: {error_message}", "error")
        messagebox.showerror("Критическая ошибка", error_message)
        
        # Возвращаемся к отображению информации
        self.show_info_view()
    
    def copy_selected_text(self):
        """Копирует выделенный текст в буфер обмена"""
        try:
            selected_text = self.info_text.selection_get()
            self.info_text.clipboard_clear()
            self.info_text.clipboard_append(selected_text)
        except tk.TclError:
            # Если ничего не выделено, копируем весь текст
            all_text = self.info_text.get("1.0", "end-1c")
            self.info_text.clipboard_clear()
            self.info_text.clipboard_append(all_text)
    
    def add_info(self, message: str, level: str = "info"):
        """Добавляет информационное сообщение"""
        from datetime import datetime
        
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Определяем цвет и стиль для разных уровней
        # Красный или зеленый цвет и всегда жирный шрифт для важных сообщений
        if level in ["success", "error", "warning"]:
            colors = {
                "warning": "#FF8C00",
                "error": "red",
                "success": "green"
            }
            color = colors[level]
            font_style = ("TkDefaultFont", 9, "bold")
        else:
            color = "black"
            font_style = ("TkDefaultFont", 9)
        
        # Вставляем текст
        if message.strip():  # Только если сообщение не пустое
            self.info_text.insert(tk.END, f"[{timestamp}] {message}\n")
        else:
            self.info_text.insert(tk.END, "\n")
        
        # Применяем цвет и стиль к последней строке
        if level in ["success", "error", "warning"]:
            start_line = self.info_text.index(tk.END + "-2l linestart")
            end_line = self.info_text.index(tk.END + "-1l lineend")
            
            tag_name = f"color_{level}_{timestamp}"
            self.info_text.tag_add(tag_name, start_line, end_line)
            self.info_text.tag_config(tag_name, foreground=color, font=font_style)
        
        # Прокручиваем в конец
        self.info_text.see(tk.END)
        
        # Обновляем интерфейс
        self.parent.update_idletasks()
    
    def add_info_to_existing(self, message: str, level: str = "info"):
        """Добавляет информацию в существующий текст без временной метки"""
        # Определяем цвет и стиль для разных уровней
        # Красный или зеленый цвет и всегда жирный шрифт для важных сообщений
        if level in ["success", "error", "warning"]:
            colors = {
                "warning": "#FF8C00",
                "error": "red", 
                "success": "green"
            }
            color = colors[level]
            font_style = ("TkDefaultFont", 9, "bold")
        else:
            color = "black"
            font_style = ("TkDefaultFont", 9)
        
        # Вставляем текст
        if message.strip():  # Только если сообщение не пустое
            self.info_text.insert(tk.END, f"{message}\n")
        else:
            self.info_text.insert(tk.END, "\n")
        
        # Применяем цвет и стиль к последней строке
        if level in ["success", "error", "warning"]:
            start_line = self.info_text.index(tk.END + "-2l linestart")
            end_line = self.info_text.index(tk.END + "-1l lineend")
            
            tag_name = f"color_{level}_no_time"
            self.info_text.tag_add(tag_name, start_line, end_line)
            self.info_text.tag_config(tag_name, foreground=color, font=font_style)
        
        # Прокручиваем в конец
        self.info_text.see(tk.END)
        
        # Обновляем интерфейс
        self.parent.update_idletasks()
    
    def add_log(self, message: str, level: str = "info"):
        """Добавляет сообщение в лог (совместимость со старым кодом)"""
        self.add_info(message, level)
        
        # Также отправляем в главное окно
        if self.main_window:
            if hasattr(self.main_window, 'add_info'):
                self.main_window.add_info(f"Создание файлов: {message}", level)
    

    def restart_process(self):
        """Перезапуск процесса создания файлов"""
        # Сбрасываем состояние обработки
        self.is_processing = False
        
        # Очищаем информацию
        self.info_text.delete(1.0, tk.END)
        self.add_info("Готов к повторному запуску")
        
        # Если есть валидные данные - проверяем возможность создания
        if (self.validation_result and 
            self.validation_result.is_valid and 
            self.output_dir_path):
            
            # Перепроверяем папку
            self.add_info("")
            self.add_info("Повторный анализ целевой папки...")
            self.check_existing_files(self.output_dir_path)
            self.check_create_button_state()
        else:
            self.add_info("Выберите файл с сотрудниками и целевую папку для начала работы")
            self.create_btn.config(state=tk.DISABLED, text="Создать файлы", command=self.create_files)

    def on_closing(self):
        """Обработчик закрытия окна"""
        if self.is_processing:
            result = messagebox.askyesno(
                "Подтверждение",
                "Идет процесс создания файлов. Действительно закрыть окно?"
            )
            if not result:
                return
        
        self.window.destroy()
        if self.main_window:
            self.main_window.on_window_closed("create_files")