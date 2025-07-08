#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Окно работы с отчетами
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime
import time

from config import Config
from core.processor import VacationProcessor
from models import ProcessingProgress, ProcessingStatus


class ReportTab:
    """Базовый класс для вкладок отчетов"""
    
    def __init__(self, parent_frame, config, processor, tab_type):
        self.frame = parent_frame
        self.config = config
        self.processor = processor
        self.tab_type = tab_type  # "departments" или "general"
        
        # Состояние
        self.target_path = ""
        self.scan_data = {}
        self.is_processing = False
        self.selected_departments = []
        
        self.setup_tab_ui()
    
    def setup_tab_ui(self):
        """Настройка интерфейса вкладки"""
        self.frame.columnconfigure(1, weight=1)
        self.frame.rowconfigure(1, weight=1)
        
        # 1. Блок выбора
        self.setup_file_selection()
        
        # 2. Информация/прогресс
        self.setup_info_progress_area()
        
        # 3. Кнопки
        self.setup_control_buttons()
    
    def setup_file_selection(self):
        """Настройка области выбора"""
        title = "Выбор папок подразделений" if self.tab_type == "departments" else "Выбор целевой папки"
        label_text = "Папки подразделений:" if self.tab_type == "departments" else "Целевая папка:"
        
        files_frame = ttk.LabelFrame(self.frame, text=title, padding="10")
        files_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E))
        files_frame.columnconfigure(1, weight=1)
        
        ttk.Label(files_frame, text=label_text).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(files_frame, textvariable=self.path_var, state="readonly")
        self.path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        self.select_btn = ttk.Button(files_frame, text="Выбрать", command=self.select_path)
        self.select_btn.grid(row=0, column=2, pady=5)
    
    def setup_info_progress_area(self):
        """Настройка области информации/прогресса"""
        # Информация (по умолчанию)
        self.info_frame = ttk.LabelFrame(self.frame, text="Информация", padding="10")
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E, tk.N, tk.S))
        self.info_frame.columnconfigure(0, weight=1)
        self.info_frame.rowconfigure(0, weight=1)
        
        self.info_text = tk.Text(self.info_frame, height=12, wrap=tk.WORD, font=("TkDefaultFont", 9), 
                                state=tk.NORMAL, cursor="arrow")
        
        # Настройка копирования
        self.setup_text_copy_behavior(self.info_text)
        
        info_scrollbar = ttk.Scrollbar(self.info_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scrollbar.set)
        
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Прогресс (скрыт)
        self.progress_frame = ttk.LabelFrame(self.frame, text="Прогресс обработки", padding="10")
        self.progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(self.progress_frame, text="Готов к началу")
        self.progress_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        # Прогресс подразделений
        ttk.Label(self.progress_frame, text="Подразделения:", font=("TkDefaultFont", 9)).grid(row=1, column=0, sticky=tk.W, pady=(10, 2))
        self.dept_progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.dept_progress_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=2)
        self.dept_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.dept_detail_label.grid(row=3, column=0, sticky=tk.W, pady=2)
        
        # Прогресс файлов
        ttk.Label(self.progress_frame, text="Файлы:", font=("TkDefaultFont", 9)).grid(row=4, column=0, sticky=tk.W, pady=(10, 2))
        self.files_progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.files_progress_bar.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=2)
        self.files_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.files_detail_label.grid(row=6, column=0, sticky=tk.W, pady=2)
        
        self.speed_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.speed_label.grid(row=7, column=0, sticky=tk.W, pady=2)
        
        # Инициализация
        initial_msg = "Выберите папки подразделений для создания отчетов" if self.tab_type == "departments" else "Выберите целевую папку для создания общего отчета"
        self.add_info(initial_msg)
    
    def setup_control_buttons(self):
        """Настройка кнопок управления"""
        buttons_frame = ttk.Frame(self.frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=(10, 15))
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(2, weight=1)
        
        btn_text = "Создать отчеты по подразделениям" if self.tab_type == "departments" else "Создать общий отчет"
        self.action_btn = ttk.Button(buttons_frame, text=btn_text, command=self.start_processing, state=tk.DISABLED)
        self.action_btn.grid(row=0, column=1)
    
    def setup_text_copy_behavior(self, text_widget):
        """Настройка поведения копирования для текстового виджета"""
        def on_key(event):
            if event.state & 0x4 and event.keysym.lower() in ['c', 'a']:
                return
            return "break"
        
        text_widget.bind('<Key>', on_key)
        text_widget.bind('<Control-a>', lambda e: text_widget.tag_add("sel", "1.0", "end"))
        
        def show_context_menu(event):
            try:
                context_menu = tk.Menu(text_widget, tearoff=0)
                context_menu.add_command(label="Выделить всё", command=lambda: text_widget.tag_add("sel", "1.0", "end"))
                context_menu.add_command(label="Копировать", command=lambda: self.copy_selected_text(text_widget))
                context_menu.tk_popup(event.x_root, event.y_root)
            except:
                pass
        
        text_widget.bind('<Button-3>', show_context_menu)
    
    def copy_selected_text(self, text_widget):
        """Копирует выделенный текст"""
        try:
            selected_text = text_widget.selection_get()
            text_widget.clipboard_clear()
            text_widget.clipboard_append(selected_text)
        except tk.TclError:
            all_text = text_widget.get("1.0", "end-1c")
            text_widget.clipboard_clear()
            text_widget.clipboard_append(all_text)
    
    def select_path(self):
        """Выбор пути"""
        dir_path = filedialog.askdirectory(title="Выберите папку с подразделениями")
        
        if not dir_path:
            return
        
        self.target_path = dir_path
        self.path_var.set(dir_path)
        self.add_info(f"Выбрана папка: {dir_path}")
        self.add_info("Сканирование папки...")
        
        def scan_thread():
            try:
                time.sleep(0.5)
                departments_info = self.processor.scan_target_directory(dir_path)
                
                if departments_info:
                    # Формируем список для диалога выбора
                    potential_departments = []
                    for dept_name, files_count in departments_info.items():
                        potential_departments.append({
                            'name': dept_name,
                            'path': str(Path(dir_path) / dept_name),
                            'files_count': files_count
                        })
                    
                    # Показываем диалог выбора
                    selected_departments = self.show_departments_selection_dialog(potential_departments)
                    
                    if selected_departments:
                        self.selected_departments = selected_departments
                        # Обновляем scan_data только выбранными
                        scan_data = {dept['name']: dept['files_count'] for dept in selected_departments}
                        self.frame.after(0, self.on_scan_complete, scan_data)
                    else:
                        return  # Пользователь отменил выбор
                else:
                    self.frame.after(0, self.on_scan_complete, {})
            except Exception as e:
                self.frame.after(0, self.on_scan_error, str(e))
        
        threading.Thread(target=scan_thread, daemon=True).start()

    def show_departments_selection_dialog(self, departments):
        """Показывает диалог для выбора подразделений"""
        # Находим родительское окно
        parent_window = self.frame.winfo_toplevel()
        
        # Создаем модальное окно
        dialog = tk.Toplevel(parent_window)
        dialog.title("Выбор подразделений")
        dialog.geometry("600x500")
        dialog.resizable(True, True)
        dialog.transient(parent_window)
        dialog.grab_set()
        
        # Центрируем относительно родительского окна
        dialog.geometry("+%d+%d" % (
            parent_window.winfo_rootx() + 50,
            parent_window.winfo_rooty() + 50
        ))
        
        result = []
        
        # Основной фрейм
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            header_frame,
            text="Выберите подразделения для создания отчетов:",
            font=("TkDefaultFont", 11, "bold")
        ).pack(anchor=tk.W)
        
        ttk.Label(
            header_frame,
            text=f"Найдено подразделений: {len(departments)}",
            font=("TkDefaultFont", 9)
        ).pack(anchor=tk.W, pady=(5, 0))
        
        # Область с чекбоксами
        list_frame = ttk.LabelFrame(main_frame, text="Подразделения", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Canvas для скроллинга
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Переменные для чекбоксов
        dept_vars = {}
        
        # Создаем чекбоксы
        checkboxes = []
        for i, dept in enumerate(departments):
            var = tk.BooleanVar()
            var.set(True)  # По умолчанию все выбраны
            dept_vars[i] = var
            
            # Фрейм для чекбокса
            cb_frame = ttk.Frame(scrollable_frame)
            cb_frame.pack(fill=tk.X, pady=2)
            
            checkbox = ttk.Checkbutton(
                cb_frame,
                text=f"{dept['name']}",
                variable=var,
                width=40
            )
            checkbox.pack(side=tk.LEFT)
            checkboxes.append(checkbox)
            
            # Информация о файлах
            info_label = ttk.Label(
                cb_frame,
                text=f"({dept['files_count']} файлов)",
                font=("TkDefaultFont", 8),
                foreground="gray"
            )
            info_label.pack(side=tk.RIGHT)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Кнопки управления выбором
        selection_frame = ttk.Frame(main_frame)
        selection_frame.pack(fill=tk.X, pady=(0, 15))
        
        def select_all():
            for var in dept_vars.values():
                var.set(True)
        
        def deselect_all():
            for var in dept_vars.values():
                var.set(False)
        
        def get_selected_count():
            return sum(1 for var in dept_vars.values() if var.get())
        
        def update_selection_info():
            selected_count = get_selected_count()
            total_files = sum(departments[i]['files_count'] for i, var in dept_vars.items() if var.get())
            selection_info.config(text=f"Выбрано: {selected_count} подразделений, {total_files} файлов")
        
        # Кнопки выбора
        selection_buttons_frame = ttk.Frame(selection_frame)
        selection_buttons_frame.pack(side=tk.LEFT)
        
        ttk.Button(selection_buttons_frame, text="Выбрать все", command=lambda: [select_all(), update_selection_info()]).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(selection_buttons_frame, text="Снять выбор", command=lambda: [deselect_all(), update_selection_info()]).pack(side=tk.LEFT, padx=5)
        
        # Информация о выборе
        selection_info = ttk.Label(selection_frame, text="", font=("TkDefaultFont", 9, "bold"))
        selection_info.pack(side=tk.RIGHT)
        
        # Обновляем информацию при изменении чекбоксов
        for var in dept_vars.values():
            var.trace('w', lambda *args: update_selection_info())
        
        # Изначально обновляем информацию
        update_selection_info()
        
        # Кнопки диалога
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)
        
        def on_ok():
            nonlocal result
            selected_indices = [i for i, var in dept_vars.items() if var.get()]
            if not selected_indices:
                messagebox.showwarning("Предупреждение", "Выберите хотя бы одно подразделение")
                return
            result = [departments[i] for i in selected_indices]
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        # Кнопки справа
        button_right_frame = ttk.Frame(buttons_frame)
        button_right_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_right_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_right_frame, text="Отмена", command=on_cancel).pack(side=tk.LEFT)
        
        # Привязываем Enter и Escape
        dialog.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())
        
        # Фокус на первый чекбокс
        if checkboxes:
            checkboxes[0].focus_set()
        
        # Ждем закрытия диалога
        dialog.wait_window()
        
        return result
    
    def on_scan_complete(self, departments_info):
        """Обработчик завершения сканирования"""
        self.scan_data = departments_info
        
        if departments_info:
            total_departments = len(departments_info)
            total_files = sum(departments_info.values())
            
            self.add_info("")
            self.add_info("АНАЛИЗ ЗАВЕРШЕН", "success")
            self.add_info("")
            
            # Показываем выбранные подразделения
            self.add_info("Выбранные подразделения:", "success")
            for dept in self.selected_departments:
                self.add_info(f"  • {dept['name']}: {dept['files_count']} файлов")
            
            self.add_info("")
            self.add_info("Статистика обработки:", "success")
            self.add_info(f"  • Подразделений: {total_departments}")
            self.add_info(f"  • Файлов сотрудников: {total_files}")
            
            # Проверка шаблона
            template_key = "block_report_template" if self.tab_type == "departments" else "general_report_template"
            template_path = Path(getattr(self.config, template_key))
            template_name = "отчета по подразделениям" if self.tab_type == "departments" else "общего отчета"
            
            if template_path.exists():
                self.add_info(f"  • Шаблон {template_name}: найден")
                can_proceed = True
            else:
                self.add_info(f"  • Шаблон {template_name}: НЕ НАЙДЕН", "error")
                self.add_info(f"    Файл: {template_path}", "error")
                can_proceed = False
            
            # Расчет времени (0.3 сек на файл)
            estimated_time = total_files * 0.3
            self.add_info(f"  • Ожидаемое время: {estimated_time:.1f} сек")
            
            self.add_info("")
            if can_proceed:
                if self.tab_type == "departments":
                    self.add_info("Будет создан отчет в каждом выбранном подразделении", "success")
                    self.add_info("Нажмите 'Создать отчеты по подразделениям' для начала", "success")
                else:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    self.add_info(f"Будет создан файл: ОБЩИЙ_ОТЧЕТ_{timestamp}.xlsx", "success")
                    self.add_info("Нажмите 'Создать общий отчет' для начала", "success")
                
                self.action_btn.config(state=tk.NORMAL)
            else:
                self.add_info("Исправьте проблему с шаблоном для продолжения", "error")
                self.action_btn.config(state=tk.DISABLED)
        else:
            self.add_info("В выбранной папке не найдено подразделений с файлами сотрудников", "warning")
            self.action_btn.config(state=tk.DISABLED)
    
    def on_scan_error(self, error_message):
        """Обработчик ошибки сканирования"""
        self.add_info(f"Ошибка сканирования: {error_message}", "error")
        messagebox.showerror("Ошибка сканирования", error_message)
    
    def start_processing(self):
        """Начало обработки"""
        if not self.scan_data:
            messagebox.showwarning("Предупреждение", "Сначала выполните сканирование")
            return
        
        self.is_processing = True
        self.action_btn.config(state=tk.DISABLED)
        self.show_progress_view()
        
        def processing_thread():
            try:
                if self.tab_type == "departments":
                    # Обновление отчетов по подразделениям
                    operation_log = self.processor.update_block_reports(
                        self.selected_departments,
                        self.on_progress_update
                    )
                else:
                    # Создание общего отчета
                    operation_log = self.processor.create_general_report(
                        self.selected_departments,
                        self.target_path,
                        self.on_progress_update
                    )
                
                self.frame.after(0, self.on_processing_complete, operation_log)
            except Exception as e:
                self.frame.after(0, self.on_processing_error, str(e))
        
        threading.Thread(target=processing_thread, daemon=True).start()
    
    def on_progress_update(self, progress):
        """Обработчик обновления прогресса"""
        def update_ui():
            # Обновляем прогресс по подразделениям
            if progress.total_blocks > 0:
                dept_percent = (progress.processed_blocks / progress.total_blocks) * 100
                self.dept_progress_bar['value'] = dept_percent
                self.dept_detail_label.config(
                    text=f"Подразделение {progress.processed_blocks}/{progress.total_blocks}: {progress.current_block or 'Готовится...'}"
                )
            
            # Обновляем прогресс по файлам
            if progress.total_files > 0:
                files_percent = (progress.processed_files / progress.total_files) * 100
                self.files_progress_bar['value'] = files_percent
                
                if progress.current_file:
                    self.files_detail_label.config(
                        text=f"Файл {progress.processed_files}/{progress.total_files}: {progress.current_file}"
                    )
            
            # Обновляем основную метку
            self.progress_label.config(text=progress.current_operation)
            
            # Показываем скорость и оставшееся время
            if progress.speed > 0 and progress.processed_files > 0:
                seconds_per_file = 1.0 / progress.speed
                remaining_files = max(0, progress.total_files - progress.processed_files)
                remaining_time = max(0, remaining_files * seconds_per_file)
                self.speed_label.config(
                    text=f"Скорость: {seconds_per_file:.2f} сек/файл, "
                         f"Осталось: {remaining_time:.0f} сек"
                )
            elif progress.processed_files == 0:
                self.speed_label.config(text="Подготовка...")
            else:
                self.speed_label.config(text="Завершение...")
        
        self.frame.after(0, update_ui)
    
    def show_progress_view(self):
        """Показать прогресс"""
        self.info_frame.grid_remove()
        self.progress_frame.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def show_info_view(self):
        """Показать информацию"""
        self.progress_frame.grid_remove()
        self.info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def on_processing_complete(self, operation_log):
        """Завершение обработки"""
        self.is_processing = False
        
        if operation_log.status == ProcessingStatus.SUCCESS:
            self.add_info_to_existing("")
            self.add_info_to_existing("=" * 50)
            self.add_info_to_existing("СОЗДАНИЕ ОТЧЕТОВ УСПЕШНО ЗАВЕРШЕНО!", "success")
            self.add_info_to_existing(f"Время выполнения: {operation_log.duration:.1f} сек")
            self.add_info_to_existing("=" * 50)
            
            for entry in operation_log.entries:
                if entry.level == "INFO":
                    self.add_info_to_existing(f"ИТОГ: {entry.message}", "success")
        else:
            self.add_info_to_existing("")
            self.add_info_to_existing("СОЗДАНИЕ ОТЧЕТОВ ЗАВЕРШЕНО С ОШИБКАМИ", "error")
            for entry in operation_log.entries:
                if entry.level == "ERROR":
                    self.add_info_to_existing(f"Ошибка: {entry.message}", "error")
        
        self.show_info_view()
        self.action_btn.config(text="Закрыть", command=self.close_window, state=tk.NORMAL)
    
    def on_processing_error(self, error_message):
        """Ошибка обработки"""
        self.is_processing = False
        self.add_info(f"Критическая ошибка: {error_message}", "error")
        messagebox.showerror("Критическая ошибка", error_message)
        self.show_info_view()
        self.action_btn.config(text="Закрыть", command=self.close_window, state=tk.NORMAL)
    
    def close_window(self):
        """Закрытие окна"""
        # Будет реализовано в главном классе
        pass
    
    def add_info(self, message: str, level: str = "info"):
        """Добавляет информационное сообщение"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level in ["success", "error", "warning"]:
            colors = {"warning": "#FF8C00", "error": "red", "success": "green"}
            color = colors[level]
            font_style = ("TkDefaultFont", 9, "bold")
        else:
            color = "black"
            font_style = ("TkDefaultFont", 9)
        
        if message.strip():
            self.info_text.insert(tk.END, f"[{timestamp}] {message}\n")
        else:
            self.info_text.insert(tk.END, "\n")
        
        if level in ["success", "error", "warning"]:
            start_line = self.info_text.index(tk.END + "-2l linestart")
            end_line = self.info_text.index(tk.END + "-1l lineend")
            tag_name = f"color_{level}_{timestamp}"
            self.info_text.tag_add(tag_name, start_line, end_line)
            self.info_text.tag_config(tag_name, foreground=color, font=font_style)
        
        self.info_text.see(tk.END)
    
    def add_info_to_existing(self, message: str, level: str = "info"):
        """Добавляет информацию без временной метки"""
        if level in ["success", "error", "warning"]:
            colors = {"warning": "#FF8C00", "error": "red", "success": "green"}
            color = colors[level]
            font_style = ("TkDefaultFont", 9, "bold")
        else:
            color = "black"
            font_style = ("TkDefaultFont", 9)
        
        if message.strip():
            self.info_text.insert(tk.END, f"{message}\n")
        else:
            self.info_text.insert(tk.END, "\n")
        
        if level in ["success", "error", "warning"]:
            start_line = self.info_text.index(tk.END + "-2l linestart")
            end_line = self.info_text.index(tk.END + "-1l lineend")
            tag_name = f"color_{level}_no_time"
            self.info_text.tag_add(tag_name, start_line, end_line)
            self.info_text.tag_config(tag_name, foreground=color, font=font_style)
        
        self.info_text.see(tk.END)


class ReportsWindow:
    """Окно для работы с отчетами по отпускам"""
    
    def __init__(self, parent: tk.Tk, config: Config, main_window):
        self.parent = parent
        self.config = config
        self.main_window = main_window
        self.logger = logging.getLogger(__name__)
        self.processor = VacationProcessor(config)
        
        self.window = None
        self.dept_tab = None
        self.general_tab = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Работа с отчетами")
        self.window.geometry("750x600")
        self.window.resizable(True, True)
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Notebook с вкладками
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Создаем вкладки
        dept_frame = ttk.Frame(self.notebook, padding="15")
        general_frame = ttk.Frame(self.notebook, padding="15")
        
        self.notebook.add(dept_frame, text="По подразделениям")
        self.notebook.add(general_frame, text="Общий")
        
        # Создаем экземпляры вкладок
        self.dept_tab = ReportTab(dept_frame, self.config, self.processor, "departments")
        self.general_tab = ReportTab(general_frame, self.config, self.processor, "general")
        
        # Привязываем закрытие окна
        self.dept_tab.close_window = self.on_closing
        self.general_tab.close_window = self.on_closing
    
    def show(self):
        """Показывает окно"""
        if self.window:
            self.window.deiconify()
            self.window.lift()
            self.window.focus()
    
    def on_closing(self):
        """Обработчик закрытия окна"""
        if (self.dept_tab and self.dept_tab.is_processing) or (self.general_tab and self.general_tab.is_processing):
            result = messagebox.askyesno("Подтверждение", "Идет процесс обработки. Действительно закрыть окно?")
            if not result:
                return
        
        self.window.destroy()
        if self.main_window:
            self.main_window.on_window_closed("reports")