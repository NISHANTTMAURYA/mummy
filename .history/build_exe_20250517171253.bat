@echo off
REM Activate your virtual environment (if not already activated)
REM If you see (venv) in your prompt, you can skip the next line
REM call venv\Scripts\activate

REM Ensure _internal folder exists and copy all dependencies into it
if not exist _internal mkdir _internal
if not exist _internal\excel_copies mkdir _internal\excel_copies
if not exist _internal\output_word_files mkdir _internal\output_word_files
copy /Y iso_excel.xlsx _internal\iso_excel.xlsx
copy /Y executive_summary_template.docx _internal\executive_summary_template.docx
REM Optionally copy any other static files you want to bundle

set BUILD_TYPE=onedir
REM set BUILD_TYPE=onefile

if "%BUILD_TYPE%"=="onedir" (
    python -m PyInstaller --noconfirm --onedir --windowed ^
      --add-data "iso_excel.xlsx;." ^
      --add-data "executive_summary_template.docx;." ^
      --add-data "_internal;_internal" ^
      --hidden-import=win32timezone ^
      --hidden-import=win32com ^
      --hidden-import=customtkinter ^
      mummy.py
) else (
    python -m PyInstaller --noconfirm --onefile --windowed ^
      --add-data "_internal;_internal" ^
      --hidden-import=win32timezone ^
      --hidden-import=win32com ^
      --hidden-import=customtkinter ^
      mummy.py
)

pause