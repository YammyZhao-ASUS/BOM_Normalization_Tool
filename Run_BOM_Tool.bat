@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"

if exist "input\60Bom.xlsx" (
    echo Found input\60Bom.xlsx
    echo Generating Chinese normalization notes into original 60BOM...
    "%PYTHON_EXE%" src\annotate_60_bom.py "input\60Bom.xlsx" --output "output\60Bom_归一化标注.xlsx"
    if errorlevel 1 (
        echo.
        echo 60BOM annotation failed. Starting basic GUI instead...
        "%PYTHON_EXE%" src\gui.py
    ) else (
        echo.
        echo Done: output\60Bom_归一化标注.xlsx
    )
) else (
    "%PYTHON_EXE%" src\gui.py
)

if errorlevel 1 (
    echo.
    echo The BOM tool stopped with an error.
    pause
)