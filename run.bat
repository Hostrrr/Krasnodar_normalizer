@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "PY="
python --version >nul 2>&1
if not errorlevel 1 (
    set "PY=python"
) else (
    py -3 --version >nul 2>&1
    if not errorlevel 1 (
        set "PY=py -3"
    )
)

if not defined PY (
    echo Python не найден.
    echo Скачайте и установите Python 3: https://www.python.org/downloads/
    echo При установке отметьте пункт "Add python.exe to PATH".
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

if not exist ".deps_installed" (
    echo Первый запуск: устанавливаю зависимости из requirements.txt ...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Не удалось установить зависимости.
        pause
        exit /b 1
    )
    echo.>.deps_installed
)

echo Запускаю веб-интерфейс нормализатора...
%PY% -m streamlit run app.py
if errorlevel 1 (
    echo Не удалось запустить Streamlit.
    pause
    exit /b 1
)
