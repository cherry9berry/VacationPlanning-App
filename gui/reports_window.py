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
        """Конструктор вкладки отчетов"""
        self.frame = parent_frame
        self.config = config
        self.processor = processor
        self.tab_type = tab_type  # "departments" или "general"
        
        # Состояние
        self.target_path = ""
        self.scan_data = {}
        self.is_processing = False
        self.selected_departments = []
        
        # НОВЫЕ ПЕРЕМЕННЫЕ для отслеживания повторных выборов
        self.path_reselected = False
        
        self.setup_tab_ui()
    
    def add_info(self, message: str, level: str = "info"):
        """Добавляет информационное сообщение"""
        # ИСПРАВЛЕНИЕ: Проверяем что виджет еще существует
        try:
            if not self.info_text.winfo_exists():
                return
        except tk.TclError:
            return
        
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

    def restart_process(self):
        """Перезапуск процесса создания отчетов"""
        # Сбрасываем состояние обработки
        self.is_processing = False
        
        # Очищаем информацию
        self.info_text.delete(1.0, tk.END)
        initial_msg = "Готов к повторному созданию отчетов по подразделениям" if self.tab_type == "departments" else "Готов к повторному созданию общего отчета"
        self.add_info(initial_msg)
        
        # Если есть валидные данные - проверяем возможность создания
        if self.scan_data and self.selected_departments:
            self.add_info("")
            self.add_info("Данные для обработки готовы")
            
            # Показываем выбранные подразделения
            total_departments = len(self.scan_data)
            total_files = sum(self.scan_data.values())
            
            self.add_info("Выбранные подразделения:", "success")
            for dept in self.selected_departments:
                self.add_info(f"  • {dept['name']}: {dept['files_count']} файлов")
            
            self.add_info("")
            self.add_info("Статистика обработки:", "success")
            self.add_info(f"  • Подразделений: {total_departments}")
            self.add_info(f"  • Файлов сотрудников: {total_files}")
            
            # Активируем кнопку
            btn_text = "Создать отчеты по подразделениям" if self.tab_type == "departments" else "Создать общий отчет"
            self.action_btn.config(text=btn_text, command=self.start_processing, state=tk.NORMAL)
        else:
            initial_msg = "Выберите папки подразделений для создания отчетов" if self.tab_type == "departments" else "Выберите целевую папку для создания общего отчета"
            self.add_info(initial_msg)
            btn_text = "Создать отчеты по подразделениям" if self.tab_type == "departments" else "Создать общий отчет"
            self.action_btn.config(text=btn_text, command=self.start_processing, state=tk.DISABLED)

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
        
        # Общий прогресс и время
        self.overall_progress_label = ttk.Label(self.progress_frame, text="Готов к началу", font=("TkDefaultFont", 10, "bold"))
        self.overall_progress_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.time_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 9))
        self.time_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # Прогресс подразделений
        label_text = "Подразделения:" if self.tab_type == "departments" else "Отделы:"
        ttk.Label(self.progress_frame, text=label_text, font=("TkDefaultFont", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=(10, 2))
        self.dept_progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.dept_progress_bar.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=2)
        self.dept_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.dept_detail_label.grid(row=4, column=0, sticky=tk.W, pady=2)
        
        # Прогресс файлов
        files_label_text = "Файлы в текущем отделе:" if self.tab_type == "departments" else "Обработка отчетов:"
        ttk.Label(self.progress_frame, text=files_label_text, font=("TkDefaultFont", 9)).grid(row=5, column=0, sticky=tk.W, pady=(10, 2))
        self.files_progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.files_progress_bar.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=2)
        self.files_detail_label = ttk.Label(self.progress_frame, text="", font=("TkDefaultFont", 8))
        self.files_detail_label.grid(row=7, column=0, sticky=tk.W, pady=2)
        
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
            self.frame.winfo_toplevel().lift()
            self.frame.winfo_toplevel().focus_force()
            return
        
        self.target_path = dir_path
        self.path_var.set(dir_path)
        self.add_info(f"Выбрана папка: {dir_path}")
        self.add_info("Сканирование папки...")
        
        # ИСПРАВЛЕНИЕ: Сбрасываем кнопку в ЗАКРЫТЬ если была завершена операция
        if self.action_btn['text'] == "Закрыть":
            # Отмечаем что был перевыбор и восстанавливаем кнопку действия qwe
            self.path_reselected = True
            btn_text = "Создать отчеты по подразделениям" if self.tab_type == "departments" else "Создать общий отчет"
            self.action_btn.config(state=tk.DISABLED, text=btn_text, command=self.start_processing)
        else:
            # Обычный сброс
            btn_text = "Создать отчеты по подразделениям" if self.tab_type == "departments" else "Создать общий отчет"
            self.action_btn.config(state=tk.DISABLED, text=btn_text, command=self.start_processing)
        
        def scan_thread():
            try:
                time.sleep(0.5)
                departments_info = self.processor.scan_target_directory(dir_path)
                
                if departments_info:
                    potential_departments = []
                    for dept_name, files_count in departments_info.items():
                        potential_departments.append({
                            'name': dept_name,
                            'path': str(Path(dir_path) / dept_name),
                            'files_count': files_count
                        })
                    
                    selected_departments = self.show_departments_selection_dialog(potential_departments)
                    
                    if selected_departments:
                        self.selected_departments = selected_departments
                        scan_data = {dept['name']: dept['files_count'] for dept in selected_departments}
                        self.frame.after(0, self.on_scan_complete, scan_data)
                    else:
                        return
                else:
                    self.frame.after(0, self.on_scan_complete, {})
            except Exception as e:
                self.frame.after(0, self.on_scan_error, str(e))
        
        threading.Thread(target=scan_thread, daemon=True).start()
        self.frame.winfo_toplevel().lift()
        self.frame.winfo_toplevel().focus_force()

    def show_departments_selection_dialog(self, departments):
        """Показывает диалог для выбора подразделений"""
        # Находим родительское окно
        parent_window = self.frame.winfo_toplevel()
        
        # Создаем модальное окно
        dialog = tk.Toplevel(parent_window)
        dialog.title("Выбор подразделений")
        dialog.geometry("650x550")
        dialog.resizable(True, True)
        dialog.transient(parent_window)
        dialog.grab_set()
        
        # Центрируем относительно родительского окна
        dialog.geometry("+%d+%d" % (
            parent_window.winfo_rootx() + 50,
            parent_window.winfo_rooty() + 50
        ))
        
        result = []
        
        # Основной фрейм с отступами
        main_frame = ttk.Frame(dialog, padding="20")
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
        
        # ИСПРАВЛЕНИЕ: Область с чекбоксами с правильной структурой
        list_frame = ttk.LabelFrame(main_frame, text="Подразделения", padding="15")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Контейнер для Canvas и Scrollbar
        canvas_container = ttk.Frame(list_frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)
        
        # Canvas для скроллинга
        canvas = tk.Canvas(canvas_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=canvas.yview)
        
        # ИСПРАВЛЕНИЕ: Создаем scrollable_frame внутри canvas с отступами
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # ИСПРАВЛЕНИЕ: Правильное размещение элементов
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Переменные для чекбоксов
        dept_vars = {}
        
        # Создаем чекбоксы с правильными отступами
        checkboxes = []
        for i, dept in enumerate(departments):
            var = tk.BooleanVar()
            var.set(True)  # По умолчанию все выбраны
            dept_vars[i] = var
            
            # ИСПРАВЛЕНИЕ: Фрейм для чекбокса с отступами
            cb_frame = ttk.Frame(scrollable_frame)
            cb_frame.pack(fill=tk.X, pady=3, padx=10)
            
            checkbox = ttk.Checkbutton(
                cb_frame,
                text=f"{dept['name']}",
                variable=var,
                width=50
            )
            checkbox.pack(side=tk.LEFT, anchor=tk.W)
            checkboxes.append(checkbox)
            
            # Информация о файлах
            info_label = ttk.Label(
                cb_frame,
                text=f"({dept['files_count']} файлов)",
                font=("TkDefaultFont", 8),
                foreground="gray"
            )
            info_label.pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))
        
        # ИСПРАВЛЕНИЕ: Обработка прокрутки колесиком мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)
        
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
        
        def on_dialog_close():
            """Обработчик закрытия диалога - возвращаем фокус на окно отчетов"""
            dialog.destroy()
            # ИСПРАВЛЕНИЕ: Возвращаем фокус на окно отчетов
            parent_window.lift()
            parent_window.focus_force()
        
        # Привязываем обработчик закрытия
        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        
        # Кнопки справа
        button_right_frame = ttk.Frame(buttons_frame)
        button_right_frame.pack(side=tk.RIGHT)
        
        def ok_and_focus():
            on_ok()
            # ИСПРАВЛЕНИЕ: Возвращаем фокус на окно отчетов после OK
            parent_window.lift()
            parent_window.focus_force()
        
        def cancel_and_focus():
            on_cancel()
            # ИСПРАВЛЕНИЕ: Возвращаем фокус на окно отчетов после Cancel
            parent_window.lift()
            parent_window.focus_force()
        
        ttk.Button(button_right_frame, text="OK", command=ok_and_focus).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_right_frame, text="Отмена", command=cancel_and_focus).pack(side=tk.LEFT)
        
        # Привязываем Enter и Escape
        dialog.bind('<Return>', lambda e: ok_and_focus())
        dialog.bind('<Escape>', lambda e: cancel_and_focus())
        
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
                
                # При повторном анализе отмечаем что была активность
                if hasattr(self, 'path_reselected'):
                    self.path_reselected = True
                
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
            # ИСПРАВЛЕНИЕ: Проверяем что окно еще существует
            try:
                if not self.frame.winfo_exists():
                    return
            except tk.TclError:
                return
            
            # Общий процент
            if self.tab_type == "departments":
                if progress.total_files > 0:
                    overall_percent = (progress.processed_files / progress.total_files) * 100
                    self.overall_progress_label.config(text=f"Общий прогресс: {overall_percent:.1f}%")
            else:
                if progress.total_blocks > 0:
                    overall_percent = (progress.processed_blocks / progress.total_blocks) * 100
                    self.overall_progress_label.config(text=f"Общий прогресс: {overall_percent:.1f}%")
            
            # Время
            elapsed = (datetime.now() - progress.start_time).total_seconds() if progress.start_time else 0
            
            if self.tab_type == "departments":
                if progress.processed_files > 0 and progress.total_files > 0:
                    speed = progress.processed_files / elapsed if elapsed > 0 else 0
                    remaining_files = progress.total_files - progress.processed_files
                    remaining_time = remaining_files / speed if speed > 0 else 0
                    self.time_label.config(text=f"Прошло: {elapsed:.0f} сек, Осталось: {remaining_time:.0f} сек")
                else:
                    self.time_label.config(text=f"Прошло: {elapsed:.0f} сек")
            else:
                if progress.processed_blocks > 0 and progress.total_blocks > 0:
                    speed = progress.processed_blocks / elapsed if elapsed > 0 else 0
                    remaining_blocks = progress.total_blocks - progress.processed_blocks
                    remaining_time = remaining_blocks / speed if speed > 0 else 0
                    self.time_label.config(text=f"Прошло: {elapsed:.0f} сек, Осталось: {remaining_time:.0f} сек")
                else:
                    self.time_label.config(text=f"Прошло: {elapsed:.0f} сек")
            
            # ИСПРАВЛЕНИЕ: Верхний прогресс-бар (отделы) - возвращаем правильную логику
            if progress.total_blocks > 0:
                # Для отчетов по блокам - показываем текущий обрабатываемый отдел
                if self.tab_type == "departments":
                    dept_percent = (progress.processed_blocks / progress.total_blocks) * 100
                    current_dept_display = progress.processed_blocks + 1 if progress.processed_blocks < progress.total_blocks else progress.total_blocks
                    self.dept_progress_bar['value'] = dept_percent
                    self.dept_detail_label.config(
                        text=f"Отдел {current_dept_display}/{progress.total_blocks}: {progress.current_block or 'Готовится...'}"
                    )
                else:
                    # Для общего отчета - показываем завершенные отделы
                    dept_percent = (progress.processed_blocks / progress.total_blocks) * 100
                    self.dept_progress_bar['value'] = dept_percent
                    self.dept_detail_label.config(
                        text=f"Отдел {progress.processed_blocks}/{progress.total_blocks}: {progress.current_block or 'Готовится...'}"
                    )
            
            # ИСПРАВЛЕННЫЙ НИЖНИЙ ПРОГРЕСС-БАР
            if self.tab_type == "departments":
                # ДЛЯ ОТЧЕТОВ ПО ПОДРАЗДЕЛЕНИЯМ - реальные файлы в текущем отделе
                if hasattr(self, 'selected_departments') and self.selected_departments:
                    current_dept_index = progress.processed_blocks
                    
                    if 0 <= current_dept_index < len(self.selected_departments):
                        current_dept = self.selected_departments[current_dept_index]
                        files_in_dept = current_dept['files_count']
                        
                        if files_in_dept > 0:
                            # ИСПРАВЛЕНИЕ: Правильный расчет файлов в текущем отделе
                            files_before_current = sum(
                                self.selected_departments[i]['files_count'] 
                                for i in range(current_dept_index)
                            )
                            
                            # Файлы обработанные в текущем отделе
                            files_in_current = max(0, progress.processed_files - files_before_current)
                            files_in_current = min(files_in_current, files_in_dept)
                            
                            files_percent = (files_in_current / files_in_dept) * 100
                            self.files_progress_bar['value'] = files_percent
                            self.files_detail_label.config(
                                text=f"Файл {files_in_current}/{files_in_dept} в отделе"
                            )
                        else:
                            self.files_progress_bar['value'] = 0
                            self.files_detail_label.config(text="Нет файлов в отделе")
                    else:
                        self.files_progress_bar['value'] = 0
                        self.files_detail_label.config(text="Инициализация...")
                else:
                    self.files_progress_bar['value'] = 0
                    self.files_detail_label.config(text="Подготовка...")
            else:
                # ДЛЯ ОБЩЕГО ОТЧЕТА - простая эмуляция на случайное время 1.5-3 сек на блок
                if hasattr(self, 'selected_departments') and self.selected_departments:
                    current_block_index = progress.processed_blocks
                    
                    if current_block_index < len(self.selected_departments):
                        dept_name = self.selected_departments[current_block_index]['name']
                        
                        # ИСПРАВЛЕНИЕ: Инициализируем переменные только один раз для всего процесса
                        if not hasattr(self, '_block_timings'):
                            self._block_timings = {}
                            import random
                            
                            # Предварительно генерируем время для каждого блока
                            for i in range(len(self.selected_departments)):
                                self._block_timings[i] = {
                                    'duration': random.uniform(1.5, 3.0),
                                    'start_time': None
                                }
                        
                        # Устанавливаем время начала для текущего блока если еще не установлено
                        if self._block_timings[current_block_index]['start_time'] is None:
                            self._block_timings[current_block_index]['start_time'] = time.time()
                        
                        block_start_time = self._block_timings[current_block_index]['start_time']
                        block_duration = self._block_timings[current_block_index]['duration']
                        time_in_block = time.time() - block_start_time
                        
                        # Прогресс в текущем блоке (от 0 до 100%)
                        block_progress = min(100, (time_in_block / block_duration) * 100)
                        
                        self.files_progress_bar['value'] = block_progress
                        self.files_detail_label.config(text=f"Обработка: {dept_name} ({block_progress:.0f}%)")
                    else:
                        # Все блоки завершены
                        self.files_progress_bar['value'] = 100
                        self.files_detail_label.config(text="Завершение...")
                else:
                    self.files_progress_bar['value'] = 0
                    self.files_detail_label.config(text="Подготовка...")
        
        # ИСПРАВЛЕНИЕ: Проверяем что фрейм еще существует перед обновлением
        try:
            if self.frame.winfo_exists():
                self.frame.after(0, update_ui)
        except tk.TclError:
            # Окно уже закрыто, игнорируем
            pass

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
        # ИСПРАВЛЕНИЕ: Проверяем что окно еще существует
        try:
            if not self.frame.winfo_exists():
                return
        except tk.TclError:
            return
        
        self.is_processing = False
        
        if operation_log.status == ProcessingStatus.SUCCESS:
            self.add_info_to_existing("")
            self.add_info_to_existing("=" * 50)
            self.add_info_to_existing("ОТЧЕТЫ УСПЕШНО СОЗДАНЫ!", "success")
            self.add_info_to_existing(f"Время: {operation_log.duration:.1f} сек")
            self.add_info_to_existing("=" * 50)
            
            for entry in operation_log.entries:
                if entry.level == "INFO":
                    # ИСПРАВЛЕНИЕ: Убираем зеленое выделение для всех ИТОГ сообщений по отделам
                    if ("Создан отчет:" in entry.message or 
                        "Данные собраны из отчета для" in entry.message or
                        "Скрипт скопирован в" in entry.message or
                        "найдено" in entry.message.lower() and "отчет" in entry.message.lower()):
                        self.add_info_to_existing(f"ИТОГ: {entry.message}")  # Обычный текст
                    else:
                        self.add_info_to_existing(f"ИТОГ: {entry.message}", "success")  # Зеленое только для итоговых сводок
        else:
            self.add_info_to_existing("")
            self.add_info_to_existing("ОШИБКА СОЗДАНИЯ ОТЧЕТОВ!", "error")
            for entry in operation_log.entries:
                if entry.level == "ERROR":
                    self.add_info_to_existing(f"ОШИБКА: {entry.message}", "error")
            
            # Показываем messagebox с ошибкой - привязываем к окну отчетов
            messagebox.showerror("Ошибка создания отчетов", "Создание отчетов завершено с ошибками. См. подробности в окне.", parent=self.frame.winfo_toplevel())
        
        self.show_info_view()
        
        # ИСПРАВЛЕНИЕ: ВСЕГДА показываем "Закрыть" после завершения
        self.action_btn.config(text="Закрыть", command=self.close_window, state=tk.NORMAL)

    def on_processing_error(self, error_message):
        """Ошибка обработки"""
        self.is_processing = False
        self.add_info(f"Критическая ошибка: {error_message}", "error")
        messagebox.showerror("Критическая ошибка", error_message, parent=self.frame.winfo_toplevel())
        self.show_info_view()
        self.action_btn.config(text="Закрыть", command=self.close_window, state=tk.NORMAL)
    
    def close_window(self):
        """Закрытие окна"""
        # Будет реализовано в главном классе
        pass
    
    def add_info_to_existing(self, message: str, level: str = "info"):
        """Добавляет информацию без временной метки"""
        # ИСПРАВЛЕНИЕ: Проверяем что виджет еще существует
        try:
            if not self.info_text.winfo_exists():
                return
        except tk.TclError:
            return
        
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