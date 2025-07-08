#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vacation Tool - Главный модуль
Приложение для управления отпусками сотрудников
"""

import sys
import os
import logging
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import warnings

# Убираем предупреждения
warnings.filterwarnings("ignore")

# Показываем консоль при запуске для обратной связи
if getattr(sys, 'frozen', False):
    # Запуск из exe - показываем что загружается
    print("Запуск Vacation Tool...")
    print("Загрузка интерфейса...")
    # Устанавливаем рабочую папку
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)
else:
    # Режим разработки
    application_path = os.path.dirname(os.path.abspath(__file__))

# Добавляем текущую папку в sys.path для импорта модулей
if __name__ == "__main__":
    current_dir = Path(__file__).parent
    sys.path.insert(0, str(current_dir))

from config import Config
from gui.main_window import MainWindow

def setup_logging():
    """Настройка логирования"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / "vacation_tool.log", encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )


def main():
    """Главная функция приложения"""
    try:
        # В exe показываем прогресс
        if getattr(sys, 'frozen', False):
            print("Настройка логирования...")
        
        # Настройка логирования
        setup_logging()
        logger = logging.getLogger(__name__)
        logger.info("Запуск Vacation Tool...")
        
        if getattr(sys, 'frozen', False):
            print(" Загрузка конфигурации...")
        
        # Загрузка конфигурации
        config = Config()
        config.load_or_create_default()
        
        if getattr(sys, 'frozen', False):
            print(" Создание интерфейса...")
        
        # Создание и запуск GUI
        root = tk.Tk()
        root.title("Vacation Tool v1.0 - Управление отпусками")
        root.geometry("750x500")
        root.resizable(False, False)
        
        # Создание главного окна
        app = MainWindow(root, config)
        
        # Обработка закрытия окна
        def on_closing():
            logger.info("Завершение работы приложения")
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        if getattr(sys, 'frozen', False):
            print("Готово! Открываем приложение...")
            # Небольшая задержка чтобы пользователь увидел что все загрузилось
            root.after(1500, lambda: None)  # Консоль закроется через 1.5 сек
        
        # Запуск главного цикла
        logger.info("GUI запущен")
        root.mainloop()
        
    except Exception as e:
        error_msg = f"Критическая ошибка: {e}"
        logging.error(error_msg, exc_info=True)
        
        # Показываем ошибку пользователю
        try:
            messagebox.showerror("Ошибка", error_msg)
        except:
            print(error_msg)
        
        sys.exit(1)


if __name__ == "__main__":
    main()