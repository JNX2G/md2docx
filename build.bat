@echo off
REM Build a standalone executable with PyInstaller (Windows)
REM Must be run after setup.bat.

if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: .venv not found. Run setup.bat first.
    pause
    exit /b 1
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Installing PyInstaller...
pip install pyinstaller

echo Building executable...
pyinstaller --onefile --add-data "static;static" --collect-data docx --collect-all mistune --hidden-import uvicorn.logging --hidden-import uvicorn.loops.auto --hidden-import uvicorn.protocols.http.auto --hidden-import uvicorn.protocols.http.h11_impl --hidden-import uvicorn.protocols.http.httptools_impl --hidden-import uvicorn.protocols.websockets.auto --hidden-import uvicorn.lifespan.on --hidden-import uvicorn.lifespan.off --hidden-import lxml.etree --hidden-import lxml._elementpath --hidden-import lxml.builder --hidden-import multipart --name md2docx main.py

echo.
echo Build complete: dist\md2docx.exe
echo Mermaid diagrams will use mermaid.ink API (internet required).
echo.
pause
