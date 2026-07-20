@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"

if exist "input\60Bom.xlsx" (
    echo Found input\60Bom.xlsx
    echo Generating BOM Intelligence report with Merge Candidate sheet...
    "%PYTHON_EXE%" src\main.py "input\60Bom.xlsx" --output "output\BOM_Intelligence_Report.xlsx"
    if errorlevel 1 (
        echo.
        echo BOM Intelligence report failed. Starting desktop app instead...
        "%PYTHON_EXE%" src\desktop_app.py
    ) else (
        echo.
        echo Done: output\BOM_Intelligence_Report.xlsx
    )
) else (
    "%PYTHON_EXE%" src\desktop_app.py
)

if errorlevel 1 (
    echo.
    echo The BOM tool stopped with an error.
    pause
)