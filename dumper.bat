﻿@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: Dump Python Code - Создание дампа всех Python файлов
:: Автор: Vacation Tool Project
:: Версия: 1.5 (исправленная рабочая версия) - Модифицирована для дампа только .py файлов

echo ========================================
echo   Dump Python Code v1.5
echo   Создание дампа всех Python файлов (только .py)
echo ========================================
echo.

:: Получаем текущую дату и время для имени файла
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do (
    set "current_date=%%c%%b%%a"
)
for /f "tokens=1-2 delims=: " %%a in ('time /t') do (
    set "current_time=%%a%%b"
)

:: Убираем точки и пробелы из времени
set "current_time=%current_time:.=%"
set "current_time=%current_time: =%"

:: Имя выходного файла
set "output_file=python_code_dump_%current_date%_%current_time%.txt"

echo Создание дампа: %output_file%
echo.

:: Счетчики
set /a file_count=0
set /a total_lines=0

:: Создаем файл с UTF-8 BOM
powershell -Command "[System.IO.File]::WriteAllText('%output_file%', '', (New-Object System.Text.UTF8Encoding $true))" >nul

:: Создаем заголовок через временный файл
echo =========================================> "%temp%\header.tmp"
echo PYTHON CODE DUMP>> "%temp%\header.tmp"
echo Generated: %date% %time%>> "%temp%\header.tmp"
echo Directory: %cd%>> "%temp%\header.tmp"
echo =========================================>> "%temp%\header.tmp"
echo.>> "%temp%\header.tmp"

powershell -Command "Get-Content '%temp%\header.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"
del "%temp%\header.tmp" >nul

:: PROJECT STRUCTURE
echo.>> "%temp%\structure.tmp"
echo ###############################################>> "%temp%\structure.tmp"
echo ### PROJECT STRUCTURE (Only .py files and their directories)>> "%temp%\structure.tmp"
echo ###############################################>> "%temp%\structure.tmp"
echo.>> "%temp%\structure.tmp"
echo === PROJECT STRUCTURE ===>> "%temp%\structure.tmp"

powershell -Command "Get-Content '%temp%\structure.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"
del "%temp%\structure.tmp" >nul

:: Простое дерево (только .py файлы и их родительские папки)
:: Collect all unique directories containing .py files
set "py_dirs_temp_file=%temp%\py_dirs.tmp"
for /R %%f in (*.py) do (
    echo %%~dpf>> "!py_dirs_temp_file!"
)

:: Get unique directories and then list their contents
set "filtered_tree_temp_file=%temp%\filtered_tree.tmp"
(for /f "delims=" %%D in ('sort /unique "!py_dirs_temp_file!"') do (
    dir /s /b "%%D*.py"
)) > "!filtered_tree_temp_file!"

:: Append the filtered tree to the output file
powershell -Command "Get-Content '%filtered_tree_temp_file%' -Encoding UTF8 | ForEach-Object { $_.Replace('%cd%\', '') } | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"
del "!py_dirs_temp_file!" >nul 2>&1
del "!filtered_tree_temp_file!" >nul 2>&1


:: Разделитель для файлов
echo.>> "%temp%\separator.tmp"
echo.>> "%temp%\separator.tmp"
echo ###############################################>> "%temp%\separator.tmp"
echo ### PYTHON FILES CONTENT>> "%temp%\separator.tmp"
echo ###############################################>> "%temp%\separator.tmp"
echo.>> "%temp%\separator.tmp"

powershell -Command "Get-Content '%temp%\separator.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"
del "%temp%\separator.tmp" >nul

:: Обходим все Python файлы
echo Поиск Python файлов...
for /R %%f in (*.py) do (
    call :process_file "%%f"
)

:: Статистика
echo.>> "%temp%\stats.tmp"
echo ###############################################>> "%temp%\stats.tmp"
echo ### STATISTICS>> "%temp%\stats.tmp"
echo ###############################################>> "%temp%\stats.tmp"
echo.>> "%temp%\stats.tmp"
echo === STATISTICS ===>> "%temp%\stats.tmp"
echo Total Python files processed: %file_count%>> "%temp%\stats.tmp"
echo Total lines of code: %total_lines%>> "%temp%\stats.tmp"
echo Dump created: %date% %time%>> "%temp%\stats.tmp"
echo =============================================>> "%temp%\stats.tmp"

powershell -Command "Get-Content '%temp%\stats.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"
del "%temp%\stats.tmp" >nul

:: Итоговое сообщение
echo.
echo ========================================
echo Дамп создан успешно!
echo ========================================
echo Файл: %output_file%
echo Обработано файлов: %file_count%
echo Всего строк кода: %total_lines%
echo ========================================
echo.

:: Открываем папку с файлом
echo Открыть папку с дампом? (Y/N)
set /p "choice=Ваш выбор: "
if /i "%choice%"=="Y" explorer .

goto :end

:: Функция обработки файла
:process_file
set "filepath=%~1"
set "filename=%~nx1"
set "relative_path=%filepath:*%cd%\=%"

echo Обработка: %relative_path%

:: Подсчитываем строки
set /a current_lines=0
for /f %%i in ('find /c /v "" "%filepath%" 2^>nul') do set /a current_lines=%%i
if !current_lines! EQU 0 set /a current_lines=0
set /a total_lines+=current_lines

:: Создаем заголовок во временном файле
echo.> "%temp%\file_header.tmp"
echo =============================================>> "%temp%\file_header.tmp"
echo FILE: %relative_path%>> "%temp%\file_header.tmp"
echo =============================================>> "%temp%\file_header.tmp"
echo Lines: !current_lines!>> "%temp%\file_header.tmp"
echo.>> "%temp%\file_header.tmp"

:: Записываем заголовок
powershell -Command "Get-Content '%temp%\file_header.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"

:: Записываем содержимое файла
powershell -Command "Get-Content '%filepath%' -Encoding UTF8 -ErrorAction SilentlyContinue | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"

:: Добавляем пустые строки
echo.>> "%temp%\empty.tmp"
echo.>> "%temp%\empty.tmp"
powershell -Command "Get-Content '%temp%\empty.tmp' -Encoding UTF8 | Out-File -FilePath '%output_file%' -Encoding UTF8 -Append"

:: Удаляем временные файлы
del "%temp%\file_header.tmp" >nul 2>&1
del "%temp%\empty.tmp" >nul 2>&1

set /a file_count+=1
goto :eof

:end
echo Нажмите любую клавишу для выхода...
pause >nul