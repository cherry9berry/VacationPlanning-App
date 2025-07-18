#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Главное окно приложения
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
from pathlib import Path
from typing import Optional

from config import Config
from core.processor import VacationProcessor
from core.events import EventBus, EventType
from gui.create_files_window import CreateFilesWindow
from gui.reports_window import ReportsWindow


class MainWindow:
    """Главное окно приложения Vacation Tool"""
    
    def __init__(self, root: tk.Tk, config: Config):
        self.root = root
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.processor = VacationProcessor(config)
        
        # Дочерние окна
        self.create_files_window: Optional[CreateFilesWindow] = None
        self.reports_window: Optional[ReportsWindow] = None
        
        # Переменная для отслеживания состояния шаблонов
        self.templates_ok = False
        
        # Подписываемся на события
        self.setup_event_listeners()
        
        self.setup_ui()
        self.check_templates()
        
        # Настраиваем периодическую проверку шаблонов
        self.schedule_template_check()
    
    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        # Запрещаем изменение размера окна
        self.root.resizable(False, False)
        
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="16")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка сетки
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)  # Инструкции растягиваются
        
        # Информация о шаблонах
        self.templates_frame = ttk.LabelFrame(main_frame, text="Статус шаблонов", padding="10")
        self.templates_frame.grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky=(tk.W, tk.E))
        self.templates_frame.columnconfigure(1, weight=1)
        
        self.template_status = {}
        templates = [
            ("Шаблон сотрудника:", self.config.employee_template),
            ("Шаблон отчета по блоку:", self.config.block_report_template),
            ("Шаблон общего отчета:", self.config.general_report_template)
        ]
        
        for i, (label_text, template_path) in enumerate(templates):
            # Метка
            label = ttk.Label(self.templates_frame, text=label_text)
            label.grid(row=i, column=0, sticky=tk.W, pady=0)
            
            # Путь к файлу
            path_label = ttk.Label(self.templates_frame, text=template_path, foreground="blue")
            path_label.grid(row=i, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=0)
            
            # Статус
            status_label = ttk.Label(self.templates_frame, text="", foreground="red")
            status_label.grid(row=i, column=2, sticky=tk.E, padx=(10, 0), pady=0)

            self.template_status[template_path] = status_label
        
        # Кнопки основных функций
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(0, 15), sticky=(tk.W, tk.E))
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(1, weight=1)
        
        # Настраиваем стиль для больших кнопок
        style = ttk.Style()
        style.configure("BigButton.TButton", padding=(10, 15))
        
        # Кнопка создания файлов
        self.create_files_btn = ttk.Button(
            buttons_frame,
            text="Создание файлов",
            command=self.open_create_files_window,
            width=25,
            style="BigButton.TButton"
        )
        self.create_files_btn.grid(row=0, column=0, padx=(0, 10), pady=0, sticky=(tk.W, tk.E))
        
        # Кнопка работы с отчетами
        self.reports_btn = ttk.Button(
            buttons_frame,
            text="Работа с отчетами",
            command=self.open_reports_window,
            width=25,
            style="BigButton.TButton"
        )
        self.reports_btn.grid(row=0, column=1, padx=(10, 0), pady=0, sticky=(tk.W, tk.E))
        
        # Инструкции и статус
        instructions_frame = ttk.LabelFrame(main_frame, text="Инструкции", padding="10")
        instructions_frame.grid(row=2, column=0, columnspan=2, pady=(0, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        instructions_frame.columnconfigure(0, weight=1)
        instructions_frame.rowconfigure(0, weight=1)
        
        # Текст инструкций с возможностью копирования
        self.instructions_text = tk.Text(
            instructions_frame,
            wrap=tk.WORD,
            font=("TkDefaultFont", 9),
            state=tk.NORMAL,
            bg=self.root.cget('bg'),
            relief=tk.FLAT,
            cursor="arrow"
        )
        
        # Настройка копирования
        def on_key(event):
            # Разрешаем только копирование
            if event.state & 0x4:  # Ctrl нажат
                if event.keysym.lower() in ['c', 'a']:
                    return  # Разрешаем Ctrl+C и Ctrl+A
            return "break"  # Блокируем все остальные клавиши
        
        self.instructions_text.bind('<Key>', on_key)
        self.instructions_text.bind('<Control-a>', lambda e: self.instructions_text.tag_add("sel", "1.0", "end"))
        
        # Контекстное меню для копирования
        def show_context_menu(event):
            try:
                context_menu = tk.Menu(self.instructions_text, tearoff=0)
                context_menu.add_command(label="Выделить всё", command=lambda: self.instructions_text.tag_add("sel", "1.0", "end"))
                context_menu.add_command(label="Копировать", command=lambda: self.copy_selected_text())
                context_menu.tk_popup(event.x_root, event.y_root)
            except:
                pass
        
        self.instructions_text.bind('<Button-3>', show_context_menu)  # Правая кнопка мыши
        
        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(instructions_frame, orient=tk.VERTICAL, command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=scrollbar.set)
        
        self.instructions_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Вставляем инструкции
        self.insert_instructions()
        self.instructions_text.config(state=tk.DISABLED)
    
    def setup_event_listeners(self):
        """Настраивает подписку на события"""
        event_bus = EventBus()
        
        # Подписываемся на события создания файлов
        event_bus.subscribe(EventType.FILE_CREATED, self._on_file_created)
        event_bus.subscribe(EventType.DIRECTORY_CREATED, self._on_directory_created)
        event_bus.subscribe(EventType.ERROR_OCCURRED, self._on_error_occurred)
        event_bus.subscribe(EventType.PROGRESS_UPDATE, self._on_progress_updated)
    
    def _on_file_created(self, event):
        """Обработчик события создания файла"""
        file_path = event.data.get("file_path")
        employee = event.data.get("employee", {})
        skipped = event.data.get("skipped", False)
        
        if skipped:
            self.logger.info(f"Файл пропущен (уже существует): {file_path}")
        else:
            self.logger.info(f"Файл создан: {file_path}")
    
    def _on_directory_created(self, event):
        """Обработчик события создания папки"""
        directory_path = event.data.get("directory_path")
        self.logger.info(f"Папка создана: {directory_path}")
    
    def _on_error_occurred(self, event):
        """Обработчик события ошибки"""
        error = event.data.get("error")
        employee = event.data.get("employee", {})
        
        error_msg = f"Ошибка: {error}"
        if employee:
            error_msg += f" (Сотрудник: {employee.get('ФИО работника', 'Неизвестно')})"
        
        self.logger.error(error_msg)
        messagebox.showerror("Ошибка", error_msg)
    
    def _on_progress_updated(self, event):
        """Обработчик события обновления прогресса"""
        progress = event.data.get("progress")
        if progress:
            # Здесь можно обновить прогресс-бар или статус
            self.logger.debug(f"Прогресс: {progress.current_operation}")
    
    def copy_selected_text(self):
        """Копирует выделенный текст в буфер обмена"""
        try:
            selected_text = self.instructions_text.selection_get()
            self.instructions_text.clipboard_clear()
            self.instructions_text.clipboard_append(selected_text)
        except tk.TclError:
            # Если ничего не выделено, копируем весь текст
            all_text = self.instructions_text.get("1.0", "end-1c")
            self.instructions_text.clipboard_clear()
            self.instructions_text.clipboard_append(all_text)
    
    def check_templates(self):
        """Проверяет наличие шаблонов и обновляет состояние кнопок"""
        validation_result = self.processor.validator.validate_templates()
        
        all_found = True
        missing_templates = []
        
        for template_path, status_label in self.template_status.items():
            if Path(template_path).exists():
                status_label.config(text="Найден", foreground="green")
            else:
                status_label.config(text="Отсутствует", foreground="red")
                all_found = False
                missing_templates.append(Path(template_path).name)
        
        # Обновляем состояние только если изменилось
        if all_found != self.templates_ok:
            self.templates_ok = all_found
            
            if all_found:
                self.templates_frame.config(text="Статус шаблонов: все найдены, функции доступны")
                self.create_files_btn.config(state=tk.NORMAL)
                self.reports_btn.config(state=tk.NORMAL)
            else:
                missing_list = ", ".join(missing_templates)
                self.templates_frame.config(text=f"Статус шаблонов: отсутствуют {missing_list}")
                self.create_files_btn.config(state=tk.DISABLED)
                self.reports_btn.config(state=tk.DISABLED)
    
    def insert_instructions(self):
        """Вставляет базовые инструкции"""
        instructions_text = """1. СОЗДАНИЕ ФАЙЛОВ ОТПУСКОВ:
   • Подготовьте Excel файл с данными сотрудников
   • Заголовки должны быть в 5-й строке файла
   • Обязательные столбцы: "ФИО работника", "Табельный номер", "Подразделение 1", "Подразделение 2", "Подразделение 3", "Подразделение 4"
   • Нажмите "Создание файлов" для начала работы

2. РАБОТА С ОТЧЕТАМИ:
   • Обновление отчетов по подразделениям на основе заполненных файлов сотрудников
   • Создание общего отчета по всей компании
   • Нажмите "Работа с отчетами" для начала

3. ТРЕБОВАНИЯ К ШАБЛОНАМ:
   • Все шаблоны должны находиться в папке templates/
   • При отсутствии шаблонов функции будут недоступны
   • Шаблоны должны иметь название: employee_template.xlsx, block_report_template.xlsx, global_report_template.xlsx"""
        
        self.instructions_text.insert(tk.END, instructions_text)
    
    def open_create_files_window(self):
        """Открывает окно создания файлов"""
        try:
            # Проверяем, есть ли уже открытое окно
            if self.create_files_window and self.create_files_window.window and self.create_files_window.window.winfo_exists():
                # Окно уже открыто, просто активируем его
                self.create_files_window.show()
                return
            
            # Создаем новое окно
            self.create_files_window = CreateFilesWindow(self.root, self.config, self)
            self.create_files_window.show()
            
        except Exception as e:
            self.logger.error(f"Ошибка открытия окна создания файлов: {e}")
            messagebox.showerror("Ошибка", f"Не удалось открыть окно создания файлов: {e}")
    
    def open_reports_window(self):
        """Открывает окно работы с отчетами"""
        try:
            # Проверяем, есть ли уже открытое окно
            if self.reports_window and self.reports_window.window and self.reports_window.window.winfo_exists():
                # Окно уже открыто, просто активируем его
                self.reports_window.show()
                return
            
            # Создаем новое окно
            self.reports_window = ReportsWindow(self.root, self.config, self)
            self.reports_window.show()
            
        except Exception as e:
            self.logger.error(f"Ошибка открытия окна отчетов: {e}")
            messagebox.showerror("Ошибка", f"Не удалось открыть окно отчетов: {e}")
    
    def schedule_template_check(self):
        """Планирует периодическую проверку шаблонов"""
        self.check_templates()
        # Проверяем каждые 5 секунд
        self.root.after(5000, self.schedule_template_check)
    
    def on_window_closed(self, window_type: str):
        """Обработчик закрытия дочерних окон"""
        if window_type == "create_files":
            self.create_files_window = None
            # Обновляем статус шаблонов при возврате
            self.check_templates()
        elif window_type == "reports":
            self.reports_window = None

    def add_info(self, message: str, level: str = "info"):
        """Добавляет информацию (для совместимости с дочерними окнами)"""
        # Эта функция нужна для совместимости с create_files_window
        pass