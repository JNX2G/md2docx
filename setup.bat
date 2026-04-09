@echo off
REM Setup script for Windows

echo Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo ERROR: python not found. Install Python 3.11+ from python.org
    echo        and ensure "Add Python to PATH" is checked during installation.
    pause
    exit /b 1
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Upgrading pip...
pip install --upgrade pip

echo Installing dependencies...
pip install -r requirements.txt

echo Installing Playwright Chromium browser...
playwright install chromium

echo.
echo Setup complete.
echo Activate with:  .venv\Scripts\activate.bat
echo Start server:   uvicorn main:app --reload --port 8765
echo Open browser:   http://localhost:8765
pause
