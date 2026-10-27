@echo off
setlocal
rem Go to the folder where this .bat is located
cd /d "%~dp0"

rem UTF-8 console + UTF-8 Python output
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

rem --- Find Python (python.exe or the py launcher) ---
set "PY_CMD="
where python >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD where py >nul 2>&1 && set "PY_CMD=py"
if not defined PY_CMD (
    echo [ERROR] Python not found. Install it from https://www.python.org/downloads/
    goto done
)

rem --- Auto-install dependencies on first run (img2pdf is required for lossless PDF) ---
"%PY_CMD%" -c "import img2pdf" >nul 2>&1
if errorlevel 1 (
    echo [SETUP] Installing required packages ^(pywin32, img2pdf, Pillow^)...
    "%PY_CMD%" -m pip install -r requirements.txt
    echo.
)

rem --- If PPTX files were dragged onto this .bat or passed as arguments ---
if not "%~1"=="" (
    "%PY_CMD%" pptx_to_pdf.py %*
    goto done
)

rem --- No arguments: the Python script will ask for the file path ---
"%PY_CMD%" pptx_to_pdf.py

:done
echo.
pause
endlocal