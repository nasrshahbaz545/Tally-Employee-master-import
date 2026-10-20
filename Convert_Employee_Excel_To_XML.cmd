@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Employee Master Excel to Tally XML Converter
echo ============================================================

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo ERROR: Python 3 is not installed or is not in PATH.
        echo Install Python 3, then run: python -m pip install openpyxl
        pause
        exit /b 1
    )
    set "PYTHON=python"
)

%PYTHON% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing required Python package: openpyxl
    %PYTHON% -m pip install openpyxl
    if errorlevel 1 (
        echo ERROR: Could not install openpyxl.
        pause
        exit /b 1
    )
)

if not exist "Employee_Master_Full_Template.xlsx" (
    echo ERROR: Employee_Master_Full_Template.xlsx was not found in this folder.
    pause
    exit /b 1
)

%PYTHON% "employee_full_xlsx_to_xml.py" "Employee_Master_Full_Template.xlsx" "Employee_Master_Full_Import.xml"
if errorlevel 1 (
    echo.
    echo Conversion failed. Read the error shown above.
    pause
    exit /b 1
)

echo.
echo SUCCESS: Employee_Master_Full_Import.xml was created.
echo Import it using Gateway of Tally ^> Import Data ^> Masters.
echo Always test in a backup company first.
pause
