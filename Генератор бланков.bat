@echo off
call venv\Scripts\activate

set PYTHON_SCRIPT=generate_pdf.py
set EXCEL_FILE="list.xlsx"
set WORD_TEMPLATE="template.docx"
set WORKING_DIRECTORY="%CD%"

python %PYTHON_SCRIPT% %EXCEL_FILE% %WORD_TEMPLATE% %WORKING_DIRECTORY%

call venv\Scripts\deactivate
pause