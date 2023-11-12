@echo off
call venv\Scripts\activate

set PYTHON_SCRIPT=main.py
set EXCEL_FILE=list.xlsx

python %PYTHON_SCRIPT% %EXCEL_FILE%

call venv\Scripts\deactivate
pause