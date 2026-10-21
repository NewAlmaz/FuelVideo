@echo off
rem ============================================================================
rem  Сборка FuelVideo.exe (Windows).
rem  Использование: просто запустить build.bat двойным щелчком.
rem
rem  Windows 10/11: подойдёт Python 3.11/3.12.
rem  Windows 7:     используйте Python 3.8.10 и PyInstaller 5.13.2.
rem ============================================================================
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo [1/5] Проверка Python...
python --version
if errorlevel 1 (
    echo ОШИБКА: Python не найден в PATH. Установите Python 3.8+ и повторите.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo [2/5] Создание виртуального окружения venv...
    python -m venv venv
    if errorlevel 1 (
        echo ОШИБКА: не удалось создать venv.
        pause
        exit /b 1
    )
) else (
    echo [2/5] Виртуальное окружение уже существует, пропускаю.
)

echo [3/5] Установка зависимостей...
call "venv\Scripts\python.exe" -m pip install --upgrade pip
call "venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo ОШИБКА: не удалось установить зависимости.
    pause
    exit /b 1
)

echo [4/5] Сборка exe (PyInstaller)...
call "venv\Scripts\pyinstaller.exe" --noconfirm --onefile --windowed --name FuelVideo ^
    --add-data "settings.ini;." ^
    --add-data "azs.ini;." ^
    --hidden-import openpyxl ^
    main.py
if errorlevel 1 (
    echo ОШИБКА: сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo [5/5] Готово.
echo.
echo exe: %~dp0dist\FuelVideo.exe
echo Рядом с exe при первом запуске будут созданы settings.ini, azs.ini и папка logs.
echo Если положить свои settings.ini и azs.ini рядом с exe — они будут использованы.
echo.
pause
endlocal
