#!/bin/bash
# Build a standalone executable with PyInstaller (macOS / Linux)
# Must be run with the venv active, or it will activate it automatically.
set -e

# Activate venv if not already active
if [ -z "$VIRTUAL_ENV" ]; then
  if [ -f ".venv/bin/activate" ]; then
    echo "Activating virtual environment..."
    source .venv/bin/activate
  else
    echo "ERROR: .venv not found. Run setup.sh first."
    exit 1
  fi
fi

echo "Installing PyInstaller..."
pip install pyinstaller

echo "Building executable..."
pyinstaller --onefile \
  --add-data "static:static" \
  --collect-data docx \
  --collect-all mistune \
  --hidden-import uvicorn.logging \
  --hidden-import uvicorn.loops.auto \
  --hidden-import uvicorn.protocols.http.auto \
  --hidden-import uvicorn.protocols.http.h11_impl \
  --hidden-import uvicorn.protocols.http.httptools_impl \
  --hidden-import uvicorn.protocols.websockets.auto \
  --hidden-import uvicorn.lifespan.on \
  --hidden-import uvicorn.lifespan.off \
  --hidden-import lxml.etree \
  --hidden-import lxml._elementpath \
  --hidden-import lxml.builder \
  --hidden-import multipart \
  --hidden-import email.mime.multipart \
  --hidden-import email.mime.text \
  --name md2docx \
  main.py

echo ""
echo "Build complete: dist/md2docx"
echo ""
echo "NOTE: Playwright (headless Chromium) is NOT bundled in the binary."
echo "      Mermaid diagrams will be rendered via mermaid.ink (internet required)"
echo "      or fall back to a placeholder image if offline."
echo ""
echo "Distribute the dist/md2docx binary — no Python required on the target machine."
