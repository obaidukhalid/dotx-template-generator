@echo off
REM Template Studio launcher for Windows
cd /d "%~dp0"

REM Preferred path: uv resolves and installs from uv.lock, then runs.
where uv >nul 2>nul
if not errorlevel 1 (
  uv run python app.py
  pause
  exit /b 0
)

REM Fallback for machines without uv: plain pip with the pinned set exported
REM from uv.lock into requirements.txt.
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.9 or newer from python.org
  echo and tick "Add Python to PATH" during setup.
  pause
  exit /b 1
)

python -c "import flask, docx" >nul 2>nul
if errorlevel 1 (
  echo Installing dependencies...
  python -m pip install -r requirements.txt
)

python app.py
pause
