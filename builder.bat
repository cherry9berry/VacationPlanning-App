@echo off
chcp 65001 >nul
echo Сборка Vacation Tool в .exe...
echo.

REM Удаляем старые файлы
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"

echo Запуск PyInstaller...
pyinstaller --onefile ^
    --console ^
    --name "VacationTool" ^
    --add-data "config.json;." ^
    --exclude-module "matplotlib" ^
    --exclude-module "pandas" ^
    --exclude-module "numpy" ^
    --exclude-module "PIL" ^
    --exclude-module "scipy" ^
    --exclude-module "pytest" ^
    --optimize=2 ^
    main.py

echo.
echo Создание финальной структуры...
if not exist "release" mkdir "release"
copy "dist\VacationTool.exe" "release\"
copy "config.json" "release\"
if exist "templates" xcopy "templates" "release\templates\" /E /I /Y

echo.
echo ГОТОВО! Сборка завершена!
echo.
echo Результат в папке 'release':
if exist "release\VacationTool.exe" (
    echo    OK VacationTool.exe - СОЗДАН
) else (
    echo    ERROR VacationTool.exe - НЕ НАЙДЕН
)

if exist "release\config.json" (
    echo    OK config.json - СКОПИРОВАН
) else (
    echo    ERROR config.json - НЕ НАЙДЕН
)

if exist "release\templates" (
    echo    OK templates\ - СКОПИРОВАНА
) else (
    echo    ERROR templates\ - НЕ НАЙДЕНА
)
REM Очистка временных файлов после сборки
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"
echo.
echo Можно запускать: release\VacationTool.exe
echo Окно закроется через 5 секунд...

REM Ждем 5 секунд и закрываемся автоматически
timeout /t 5 /nobreak >nul