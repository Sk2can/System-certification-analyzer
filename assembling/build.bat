```bat
@echo off
chcp 65001 >nul
setlocal
title Сборка SCA

set "PROJECT_DIR=%~dp0.."
for %%I in ("%PROJECT_DIR%") do set "PROJECT_DIR=%%~fI"

set "PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
set "BUILDER=%~dp0build.py"

if not exist "%PYTHON%" (
    echo ОШИБКА: Не найден Python:
    echo "%PYTHON%"
    pause
    exit /b 1
)

if not exist "%BUILDER%" (
    echo ОШИБКА: Не найден build.py:
    echo "%BUILDER%"
    pause
    exit /b 1
)

"%PYTHON%" "%BUILDER%"

echo.
pause
```