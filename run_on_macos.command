#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

PYTHON_CMD=""
if command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_CMD="python"
fi

if [ -z "$PYTHON_CMD" ]; then
  echo "[ERROR] Python was not found. Please install Python 3 and try again."
  read -r -p "Press Enter to exit..."
  exit 1
fi

echo "Checking Tkinter..."
if ! "$PYTHON_CMD" -c "import tkinter" >/dev/null 2>&1; then
  echo "[ERROR] Tkinter is not available in this Python installation."
  echo "Install Python from https://www.python.org/downloads/macos/ or use a Python build that includes Tkinter."
  read -r -p "Press Enter to exit..."
  exit 1
fi

echo "Checking required Python packages..."
if ! "$PYTHON_CMD" -c "import importlib.util, sys; missing=[m for m in ['pandas','openpyxl','docxtpl','docx2pdf','tkinterdnd2'] if importlib.util.find_spec(m) is None]; print(','.join(missing)); sys.exit(1 if missing else 0)"; then
  echo "Missing packages found. Installing from requirements.txt..."
  "$PYTHON_CMD" -m pip install -r requirements.txt
fi

echo "Starting Contract Generator..."
"$PYTHON_CMD" app.py

read -r -p "Press Enter to exit..."
