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
import os


from config import Config
from core.processor import VacationProcessor
from core.events import event_bus, EventType
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
        
        # Для отката созданных файлов и папок
        self.created_files = []
        self.created_dirs = []
        self._rollback_in_progress = False
        self._rollback_lock = threading.Lock()
        
        # Создаем окно
        self.window = None
        self.created_count_label = tk.Label(self.window, text="Создано: 0")
        self.created_count_label.grid(row=0, column=0, sticky='w')
        # Инициализация переменной прогресса для совместимости с on_processing_complete
        self.progress_var = tk.IntVar(value=0)
        self.stop_processing = False
        
        # Подписка на события
        self._setup_event_listeners()
        
        self.setup_ui()

        # Инициализация дополнительных переменных (но НЕ перезаписываем create_btn!)
        self.skipped_count_label = None
        self.error_count_label = None
        self.status_label = None
        self.create_button = None
        self.back_button = None
        self.created_count_label = None

    def _setup_event_listeners(self):
        """Настройка подписки на события"""
        event_bus.subscribe(EventType.FILE_CREATED, self._on_file_created)
        event_bus.subscribe(EventType.DIRECTORY_CREATED, self._on_directory_created)
        event_bus.subscribe(EventType.ERROR_OCCURRED, self._on_error_occurred)
    
    def _on_file_created(self, event):
        """Обработчик события создания файла"""
        file_path = event.data.get("file_path")
        if file_path and hasattr(self, 'created_files'):
            self.created_files.append(file_path)
    
    def _on_directory_created(self, event):
        """Обработчик события создания папки"""
        directory_path = event.data.get("directory_path")
        if directory_path and hasattr(self, 'created_dirs'):
            self.created_dirs.append(directory_path)
    
    def _on_error_occurred(self, event):
        """Обработчик события ошибки"""
        error = event.data.get("error")
        if error:
            self.logger.error(f"Ошибка из системы событий: {error}")

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
        main_frame.grid(row=0, column=0, sticky="nsew")
        
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
        files_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="ew")
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
        self.staff_file_entry.grid(row=0, column=1, sticky="ew", padx=(10, 5), pady=5)
        
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
        self.output_dir_entry.grid(row=1, column=1, sticky="ew", padx=(10, 5), pady=5)
        
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
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky="nsew")
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
        
        self.info_text.grid(row=0, column=0, sticky="nsew")
        info_scrollbar.grid(row=0, column=1, sticky="ns")
        
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
        self.departments_progress_bar.grid(row=3, column=0, sticky="ew", pady=2)
        
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
        self.employees_progress_bar.grid(row=6, column=0, sticky="ew", pady=2)
        
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
            if self.create_btn is not None:
                self.create_btn.config(state=tk.DISABLED, text="Создать файлы", command=self.create_files)
            
            # Автоматически запускаем валидацию
            self.validate_file()
        
        # Возвращаем фокус на окно создания файлов
        if self.window:
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
            if self.create_btn is not None:
                self.create_btn.config(state=tk.DISABLED, text="Создать файлы", command=self.create_files)
            
            # Проверяем существующие файлы
            self.add_info("")
            self.add_info("Анализ целевой папки...")
            self.check_existing_files(dir_path)
            
            # Проверяем возможность активации кнопки создания
            self.check_create_button_state()
        
        # Возвращаем фокус на окно создания файлов
        if self.window:
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
                if emp['Подразделение 1']:
                    clean_dept = self._clean_directory_name(emp['Подразделение 1'])
                    expected_departments.add((emp['Подразделение 1'], clean_dept))
            
            # --- Формируем список сотрудников для создания файлов ---
            employees_to_create = []
            for emp in self._employees:
                dept = emp['Подразделение 1']
                if not dept:
                    continue
                clean_dept = self._clean_directory_name(dept)
                dept_path = base_path / clean_dept
                filename = self.processor.excel_handler.generate_output_filename(emp)
                file_path = dept_path / filename
                if not file_path.exists():
                    employees_to_create.append(emp)
            self._employees_to_create = employees_to_create
            
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
            all_departments_from_file = set(emp['Подразделение 1'] for emp in self._employees if emp['Подразделение 1'])
            new_departments = list(all_departments_from_file - set(existing_departments))
            
            # Подсчитываем новых сотрудников
            self.new_employees_count = len(self._employees_to_create)
            self.skip_employees_count = len(self._employees) - self.new_employees_count
            
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
                        processing_time = self.config.get("processing_time_per_file", 0.3)
                        if processing_time is not None:
                            estimated_time = max(0.1, self.new_employees_count * processing_time)
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
                all_departments = list(set(emp['Подразделение 1'] for emp in self._employees if emp['Подразделение 1']))
                new_list = ", ".join(all_departments)
                
                self.add_info("В папке нет подразделений из вашего файла - будет создана новая структура")
                self.add_info(f"Новые отделы: {new_list}")
                
                self.add_info("")
                self.add_info("ПЛАН ОБРАБОТКИ:", "success")
                self.add_info(f"  • Будет создано: {self.new_employees_count} сотр.")
                self.add_info(f"  • Будет пропущено: {self.skip_employees_count} сотр.")
                self.add_info(f"  • Всего сотрудников: {len(self._employees)}")
                
                # Рассчитываем ожидаемое время для всех файлов
                processing_time = self.config.get("processing_time_per_file", 0.3)
                if processing_time is not None:
                    estimated_time = max(0.1, self.new_employees_count * processing_time)
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
                if self.window:
                    self.window.after(0, self.on_validation_complete, employees)
                
            except Exception as e:
                if self.window:
                    self.window.after(0, self.on_validation_error, str(e))
        
        threading.Thread(target=validate_thread, daemon=True).start()
        
    def on_validation_complete(self, employees):
            """Обработчик завершения валидации"""
            # Сохраняем сотрудников (уже отфильтрованных в validator)
            self._employees = employees
            
            if self.validation_result and self.validation_result.is_valid:
                self.add_info("")
                self.add_info("ВАЛИДАЦИЯ УСПЕШНО ЗАВЕРШЕНА", "success")
                
                # ИСПРАВЛЕНО: Теперь список employees уже отфильтрован и не содержит дублирующихся табельных номеров
                # Поэтому дополнительная проверка не нужна, но покажем предупреждения из валидации
                warnings = getattr(self.validation_result, 'warnings', None) if self.validation_result else None
                if warnings:
                    self.add_info("")
                    self.add_info("НАЙДЕНЫ ПРОБЛЕМЫ:", "warning")
                    for warning in warnings:
                        self.add_info(f"  • {warning}")
                
                self.add_info("")
                self.add_info("СТАТИСТИКА ФАЙЛА:", "success")
                
                # ИСПРАВЛЕНО: Используем правильную статистику
                total_after_filter = len(employees)  # Количество после фильтрации
                unique_tab_numbers = len(set(emp.get('Табельный номер', '') for emp in employees if emp.get('Табельный номер')))
                warnings_count = len(self.validation_result.warnings) if self.validation_result and self.validation_result.warnings is not None else 0
                
                self.add_info(f"  • Всего сотрудников после фильтрации: {total_after_filter}")
                self.add_info(f"  • Уникальных табельных номеров: {unique_tab_numbers}")
                if warnings_count > 0:
                    self.add_info(f"  • Предупреждений: {warnings_count}")
                
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
                errors = getattr(self.validation_result, 'errors', None) if self.validation_result else None
                if errors:
                    for error in errors:
                        self.add_info(f"  • {error}", "error")
                
                warnings = getattr(self.validation_result, 'warnings', None) if self.validation_result else None
                if warnings:
                    self.add_info("")
                    self.add_info("ПРЕДУПРЕЖДЕНИЯ:", "warning")
                    for warning in warnings:
                        self.add_info(f"  • {warning}", "warning")

    def check_employee_uniqueness(self, employees):
        """УДАЛЕНО: Этот метод больше не нужен, так как фильтрация происходит в validator"""
        # Метод оставлен для совместимости, но всегда возвращает успех
        # поскольку employees уже отфильтрован в validator
        return {'is_valid': True, 'errors': []}

    def format_validation_stats(self, validation_result, employees):
        """ИСПРАВЛЕНО: Форматирует статистику валидации с учетом фильтрации"""
        total_after_filter = len(employees)
        unique_tab_numbers = len(set(emp.get('Табельный номер', '') for emp in employees if emp.get('Табельный номер')))
        warnings = getattr(validation_result, 'warnings', None) if validation_result else None
        warnings_count = len(warnings) if warnings else 0
        
        stats = f"• Всего сотрудников после фильтрации: {total_after_filter}\n"
        stats += f"• Уникальных табельных номеров: {unique_tab_numbers}\n"
        
        if warnings_count > 0:
            stats += f"• Предупреждений: {warnings_count}"
        
        return stats

    def on_validation_error(self, error_message):
        """Обработчик ошибки валидации"""
        self.add_info(f"Ошибка валидации: {error_message}", "error")
        messagebox.showerror("Ошибка валидации", error_message, parent=self.window)
    
    
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
            self.create_btn is not None and self.create_btn['text'] != "Закрыть"):
            
            if self.create_btn is not None:
                self.create_btn.config(state=tk.NORMAL, text="Создать файлы", command=self.create_files)
        elif self.create_btn is not None and self.create_btn['text'] == "Закрыть":
            if self.create_btn is not None:
                self.create_btn.config(
                    text="Создать файлы",
                    state=tk.DISABLED,  # или tk.NORMAL, если все условия выполнены
                    command=self.create_files
                )
        else:
            if self.create_btn is not None:
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
            print(f"is_processing: {self.is_processing}")
            print(f"validation_result: {self.validation_result}")
            print(f"validation_result.is_valid: {self.validation_result.is_valid if self.validation_result else 'None'}")
            print(f"staff_file_path: {self.staff_file_path}")
            print(f"output_dir_path: {self.output_dir_path}")
            print(f"new_employees_count: {getattr(self, 'new_employees_count', 'НЕТ АТРИБУТА')}")
            
            self.is_processing = True
            if self.create_btn is not None:
                self.create_btn.config(state=tk.DISABLED)
            
            # Переключаемся на отображение прогресса
            self.show_progress_view()
            
            self.add_info("Начало создания файлов...")
            
            def processing_thread():
                try:
                    # Передаем только сотрудников, для которых нужно создавать файлы
                    employees_to_create = getattr(self, '_employees_to_create', None)
                    operation_log = self.processor.create_employee_files_to_existing(
                        self.staff_file_path,
                        self.output_dir_path,
                        self.on_progress_update,
                        self.on_department_progress_update,
                        self.on_file_progress_update,
                        employees_to_create=employees_to_create
                    )
                    
                    # Завершение в главном потоке
                    def after_processing():
                        if self.stop_processing:
                            self.rollback_created_files()
                        if self.window is not None:
                            self.window.after(0, self.on_processing_complete, operation_log)
                    if self.window is not None:
                        self.window.after(0, after_processing)
                    
                except Exception as e:
                    self.logger.error(f"Ошибка в потоке обработки: {e}")
                    import traceback
                    traceback.print_exc()
                    if self.window is not None:
                        self.window.after(0, self.on_processing_error, str(e))
            
            thread = threading.Thread(target=processing_thread, daemon=True)
            thread.start()

    def show_progress_view(self):
        """Показывает область прогресса вместо информации"""
        self.info_frame.grid_remove()
        self.progress_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky="nsew")
    
    def show_info_view(self):
        """Показывает область информации вместо прогресса"""
        self.progress_frame.grid_remove()
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky="nsew")

    def on_progress_update(self, progress):
            """ИСПРАВЛЕНО: Обработчик обновления общего прогресса"""
            def update_ui():
                try:
                    # Проверяем что окно еще существует
                    if self.window is None or not self.window.winfo_exists():
                        return
                        
                    # Общий процент
                    if progress.total_files > 0:
                        overall_percent = (progress.processed_files / progress.total_files) * 100
                        self.overall_progress_label.config(text=f"Общий прогресс: {overall_percent:.1f}%")
                    
                    # Время
                    if hasattr(progress, 'start_time') and progress.start_time:
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
                            
                except tk.TclError:
                    # Окно уже закрыто, игнорируем
                    pass
            
            # ИСПРАВЛЕНО: Используем правильный метод для обновления в главном потоке
            try:
                if self.window is not None and self.window.winfo_exists():
                    self.window.after(0, update_ui)
            except tk.TclError:
                pass

    def on_department_progress_update(self, current_dept, total_depts, dept_name):
            """ИСПРАВЛЕНО: Обработчик обновления прогресса по отделам"""
            def update_ui():
                try:
                    if self.window is None or not self.window.winfo_exists():
                        return
                        
                    if total_depts > 0:
                        dept_percent = (current_dept / total_depts) * 100
                        self.departments_progress_bar['value'] = dept_percent
                        self.departments_detail_label.config(
                            text=f"Отдел {current_dept}/{total_depts}: {dept_name}"
                        )
                except tk.TclError:
                    pass
            
            try:
                if self.window is not None and self.window.winfo_exists():
                    self.window.after(0, update_ui)
            except tk.TclError:
                pass

    def on_file_progress_update(self, current_file, total_files, file_info):
            """ИСПРАВЛЕНО: Обработчик обновления прогресса по файлам в текущем отделе"""
            def update_ui():
                try:
                    if self.window is None or not self.window.winfo_exists():
                        return
                        
                    if total_files > 0:
                        file_percent = (current_file / total_files) * 100
                        self.employees_progress_bar['value'] = file_percent
                        self.employees_detail_label.config(
                            text=f"Файл {current_file}/{total_files}: {file_info}"
                        )
                except tk.TclError:
                    pass
            
            try:
                if self.window is not None and self.window.winfo_exists():
                    self.window.after(0, update_ui)
            except tk.TclError:
                pass

    def on_processing_complete(self, result):
        """Обработчик завершения создания файлов"""
        
        try:
            if self.window is None or not self.window.winfo_exists():
                return
            
            # Останавливаем прогресс бар
            self.progress_var.set(100)
            
            # Извлекаем данные из OperationLog
            if hasattr(result, 'entries'):
                # Это OperationLog объект
                log_entries = result.entries
                processing_time = result.duration or 0.0
                
                # Подсчитываем статистику из записей лога
                created_count = 0
                skipped_count = 0
                error_count = 0
                average_time_per_file = 0.0
                
                for entry in log_entries:
                    message = entry.get('message', '')
                    level = entry.get('level', 'INFO')
                    
                    if 'создано' in message.lower() and 'файлов' in message.lower():
                        # Ищем количество созданных файлов
                        import re
                        match = re.search(r'(\d+)', message)
                        if match:
                            created_count = int(match.group(1))
                    elif 'пропущено' in message.lower():
                        # Ищем количество пропущенных файлов
                        import re
                        match = re.search(r'(\d+)', message)
                        if match:
                            skipped_count = int(match.group(1))
                    elif 'среднее время на файл' in message.lower():
                        # Ищем среднее время на файл
                        import re
                        match = re.search(r'(\d+\.?\d*)с', message)
                        if match:
                            average_time_per_file = float(match.group(1))
                    elif level == 'ERROR':
                        error_count += 1
                
            else:
                # Старый формат (словарь или объект с атрибутами)
                if isinstance(result, dict):
                    created_count = result.get('created', 0)
                    skipped_count = result.get('skipped', 0)
                    error_count = result.get('errors', 0)
                    processing_time = result.get('processing_time', 0)
                    log_entries = result.get('log_entries', [])
                else:
                    created_count = getattr(result, 'created', 0)
                    skipped_count = getattr(result, 'skipped', 0)
                    error_count = getattr(result, 'errors', 0)
                    processing_time = getattr(result, 'processing_time', 0)
                    log_entries = getattr(result, 'log_entries', [])
            
            # Обновляем счетчики (если элементы существуют)
            if self.created_count_label is not None:
                self.created_count_label.config(text=f"Создано: {created_count}")
            if self.skipped_count_label is not None:
                self.skipped_count_label.config(text=f"Пропущено: {skipped_count}")
            if self.error_count_label is not None:
                self.error_count_label.config(text=f"Ошибок: {error_count}")
            
            # Показываем итоговое сообщение
            if error_count > 0:
                status_message = f"Завершено с ошибками: {created_count} создано, {skipped_count} пропущено, {error_count} ошибок"
                self.add_info_to_existing("=" * 50)
                self.add_info_to_existing("СОЗДАНИЕ ФАЙЛОВ ЗАВЕРШЕНО С ОШИБКАМИ!")
                self.add_info_to_existing(f"Время выполнения: {processing_time:.1f} сек")
                if average_time_per_file > 0:
                    self.add_info_to_existing(f"Среднее время на файл: {average_time_per_file:.2f} сек")
                self.add_info_to_existing("=" * 50)
            else:
                status_message = f"Успешно завершено: {created_count} создано, {skipped_count} пропущено"
                self.add_info_to_existing("=" * 50)
                self.add_info_to_existing("СОЗДАНИЕ ФАЙЛОВ УСПЕШНО ЗАВЕРШЕНО!")
                self.add_info_to_existing(f"Время выполнения: {processing_time:.1f} сек")
                if average_time_per_file > 0:
                    self.add_info_to_existing(f"Среднее время на файл: {average_time_per_file:.2f} сек")
                self.add_info_to_existing("=" * 50)
            
            # Обрабатываем лог записи
            if log_entries:
                for entry in log_entries:
                    if isinstance(entry, dict):
                        # Если entry - словарь
                        level = entry.get('level', 'INFO')
                        message = entry.get('message', '')
                        if level == "ERROR":
                            self.add_info_to_existing(f"ОШИБКА: {message}")
                        elif level == "WARNING":
                            self.add_info_to_existing(f"ПРЕДУПРЕЖДЕНИЕ: {message}")
                        else:
                            self.add_info_to_existing(message)
                    else:
                        # Если entry - объект с атрибутами
                        level = getattr(entry, 'level', 'INFO')
                        message = getattr(entry, 'message', str(entry))
                        if level == "ERROR":
                            self.add_info_to_existing(f"ОШИБКА: {message}")
                        elif level == "WARNING":
                            self.add_info_to_existing(f"ПРЕДУПРЕЖДЕНИЕ: {message}")
                        else:
                            self.add_info_to_existing(message)
            
            # Обновляем статус (если элемент существует)
            if self.status_label is not None:
                self.status_label.config(text=status_message)
            
            # Включаем кнопки (если элементы существуют)
            if self.create_button is not None:
                self.create_button.config(state=tk.NORMAL)
            if self.back_button is not None:
                self.back_button.config(state=tk.NORMAL)
            
            # Устанавливаем флаг завершения
            self.is_processing = False
            
            if self.create_btn is not None:
                self.create_btn.config(
                    text="Закрыть",
                    state=tk.NORMAL,
                    command=self.on_closing
                )
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            
            # Безопасное завершение
            try:
                self.add_info_to_existing(f"Ошибка при завершении: {e}")
                if self.status_label is not None:
                    self.status_label.config(text="Ошибка при завершении обработки")
                if self.create_button is not None:
                    self.create_button.config(state=tk.NORMAL)
                if self.back_button is not None:
                    self.back_button.config(state=tk.NORMAL)
                self.is_processing = False
            except:
                pass

    def on_processing_error(self, error_message):
        """ИСПРАВЛЕНО: Обработчик ошибки обработки"""
        try:
            if self.window is None or not self.window.winfo_exists():
                return
        except tk.TclError:
            return
            
        self.is_processing = False
        
        self.add_info_to_existing("")
        self.add_info_to_existing("КРИТИЧЕСКАЯ ОШИБКА ПРИ СОЗДАНИИ ФАЙЛОВ!", "error")
        self.add_info_to_existing(f"Ошибка: {error_message}", "error")
        
        # Возвращаемся к отображению информации
        self.show_info_view()
        
        # Показываем кнопку "Перезапустить"
        if self.create_btn is not None:
            self.create_btn.config(text="Перезапустить", command=self.restart_process, state=tk.NORMAL)
        
        messagebox.showerror("Ошибка", f"Критическая ошибка при создании файлов:\n{error_message}")


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
        """ИСПРАВЛЕНО: Добавляет информацию к существующему тексту"""
        try:
            if self.window is None or not self.window.winfo_exists():
                return
        except tk.TclError:
            return
            
        # Переключаемся обратно на info view если мы в progress view
        current_frame = None
        try:
            if self.info_frame.winfo_viewable():
                current_frame = "info"
            elif self.progress_frame.winfo_viewable():
                current_frame = "progress"
                self.show_info_view()  # Переключаемся на info
        except tk.TclError:
            pass
        
        # Добавляем сообщение
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
            if self.create_btn is not None:
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
            self.stop_processing = True
            # Не вызываем rollback здесь! Откат будет вызван после завершения потока
        if self.window is not None:
            self.window.destroy()
        if self.main_window:
            self.main_window.on_window_closed("create_files")

    # --- ОТКАТ СОЗДАННЫХ ФАЙЛОВ ---
    def rollback_created_files(self):
        """Удаляет все созданные в этом запуске файлы и новые папки (если они пусты)"""
        with self._rollback_lock:
            self._rollback_in_progress = True
            # Удаляем файлы
            for file_path in reversed(self.created_files):
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception:
                    pass
            # Удаляем новые папки, если они пусты
            for dir_path in reversed(self.created_dirs):
                try:
                    if os.path.isdir(dir_path) and not os.listdir(dir_path):
                        os.rmdir(dir_path)
                except Exception:
                    pass
            self.created_files.clear()
            self.created_dirs.clear()
            self._rollback_in_progress = False