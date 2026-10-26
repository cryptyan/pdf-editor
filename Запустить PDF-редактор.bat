@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import pymupdf" 2>nul || (
  echo Устанавливаю PyMuPDF...
  python -m pip install pymupdf
)
rem PDF можно перетащить прямо на этот файл — он сразу откроется в редакторе
python pdf_editor.py %*
pause
