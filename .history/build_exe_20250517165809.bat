@echo off
REM Activate your virtual environment (if not already activated)
REM If you see (venv) in your prompt, you can skip the next line
REM call venv\Scripts\activate

set BUILD_TYPE=onedir
REM set BUILD_TYPE=onefile

if "%BUILD_TYPE%"=="onedir" (
    python -m PyInstaller --noconfirm --onedir --windowed ^
      --add-data "excel_copies;excel_copies" ^
      --add-data "output_word_files;output_word_files" ^
      --add-data "executive_summary_template.docx;." ^
      --add-data "iso_excel.xlsx;." ^
      --hidden-import=win32timezone ^
      --hidden-import=win32com ^
      --hidden-import=customtkinter ^
      mummy.py
) else (
    python -m PyInstaller --noconfirm --onefile --windowed ^
      --add-data "excel_copies;excel_copies" ^
      --add-data "output_word_files;output_word_files" ^
      --add-data "executive_summary_template.docx;." ^
      --add-data "iso_excel.xlsx;." ^
      --hidden-import=win32timezone ^
      --hidden-import=win32com ^
      --hidden-import=customtkinter ^
      mummy.py
)

pause