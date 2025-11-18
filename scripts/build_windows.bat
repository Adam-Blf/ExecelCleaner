@echo off
python -m venv .venv
call .venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt
pyinstaller excel_cleaner.spec
echo Built at dist\ExcelCleaner\ExcelCleaner.exe
pause
