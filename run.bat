@echo off
setlocal EnableExtensions
title Krasnodar normalizer
cd /d "%~dp0"

echo.
echo ========================================
echo  Krasnodar city normalizer
echo ========================================
echo  Folder: %CD%
echo ========================================
echo.

chcp 65001 >nul 2>&1

set "PY="

py -3 -c "import sys; raise SystemExit(0 if sys.version_info>=(3,8) and 'WindowsApps' not in sys.executable else 1)" 2>nul
if not errorlevel 1 (
    set "PY=py -3"
    goto :found_python
)

python -c "import sys; raise SystemExit(0 if sys.version_info>=(3,8) and 'WindowsApps' not in sys.executable else 1)" 2>nul
if not errorlevel 1 (
    set "PY=python"
    goto :found_python
)

python3 -c "import sys; raise SystemExit(0 if sys.version_info>=(3,8) and 'WindowsApps' not in sys.executable else 1)" 2>nul
if not errorlevel 1 (
    set "PY=python3"
    goto :found_python
)

echo.
echo [ERROR] Python 3.8+ not found in PATH.
echo.
echo Install Python: https://www.python.org/downloads/
echo During setup tick: "Add python.exe to PATH"
echo Then close this window and double-click run.bat again.
echo.
echo Python не найден. Установите Python 3 и отметьте Add python.exe to PATH.
echo.
start https://www.python.org/downloads/
echo.
pause
exit /b 1

:found_python
echo Using: %PY%
%PY% --version
echo.

if not exist "app.py" (
    echo [ERROR] app.py not found.
    echo Unpack the ZIP first, then run run.bat from that folder.
    echo Распакуйте архив ZIP и запускайте run.bat из папки проекта.
    echo.
    pause
    exit /b 1
)

if not exist ".deps_installed" (
    echo First run: installing requirements ...
    echo Первый запуск: установка зависимостей ...
    echo.
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] pip install failed.
        echo Try:  %PY% -m ensurepip --upgrade
        echo Then run run.bat again.
        echo.
        pause
        exit /b 1
    )
    echo installed> ".deps_installed"
    echo.
)

echo Starting Streamlit ...
echo After start open in browser: http://localhost:8501
echo Stop: Ctrl+C
echo.
%PY% -m streamlit run app.py --server.headless false
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
    echo [ERROR] Streamlit exited with code %ERR%
) else (
    echo Streamlit stopped.
)
echo.
pause
