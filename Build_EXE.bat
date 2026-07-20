@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"

"%PYTHON_EXE%" -c "import openpyxl, pandas, xlrd, PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo One or more build dependencies are missing.
    echo Run: python -m pip install -r requirements-build.txt
    exit /b 1
)

"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean bom_tool.spec
if errorlevel 1 (
    echo.
    echo Build failed.
    exit /b 1
)

start "" /wait "dist\BOM_Intelligence_Platform.exe" --self-test
if errorlevel 1 (
    echo.
    echo Build completed, but the packaged self-test failed.
    echo See: %%TEMP%%\BOM_Intelligence_self_test.log
    exit /b 1
)

echo.
echo Built and verified: dist\BOM_Intelligence_Platform.exe