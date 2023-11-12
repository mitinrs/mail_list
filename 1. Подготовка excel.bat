@echo off
call venv\Scripts\activate

set PYTHON_SCRIPT=function_excel_modify.py
set EXCEL_FILE=list.xlsx

python %PYTHON_SCRIPT% %EXCEL_FILE%

call venv\Scripts\deactivate
pause