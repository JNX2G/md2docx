#!/bin/bash
# Setup script for macOS / Linux
set -e

echo "Creating virtual environment..."
python3 -m venv .venv

echo "Activating virtual environment..."
source .venv/bin/activate

echo "Upgrading pip..."
pip install --upgrade pip

echo "Installing dependencies..."
pip install -r requirements.txt

echo "Installing Playwright Chromium browser..."
playwright install chromium

echo ""
echo "Setup complete."
echo "Activate with:  source .venv/bin/activate"
echo "Start server:   uvicorn main:app --reload --port 8765"
echo "Open browser:   http://localhost:8765"
