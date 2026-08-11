#!/usr/bin/env bash
# Template Studio launcher for macOS and Linux
set -e

cd "$(dirname "$0")"

VENV_DIR=".venv"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found. Install Python 3.9 or newer."
  exit 1
fi

if [ ! -d "$VENV_DIR" ]; then
  echo "Creating virtual environment in $VENV_DIR ..."
  if ! python3 -m venv "$VENV_DIR"; then
    echo
    echo "Could not create the virtual environment."
    echo "On Debian and Ubuntu install the venv package first:"
    echo "  sudo apt install python3-venv"
    exit 1
  fi
fi

# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

if ! python -c "import flask, docx" >/dev/null 2>&1; then
  echo "Installing dependencies into the virtual environment ..."
  python -m pip install --upgrade pip >/dev/null
  python -m pip install -r requirements.txt
fi

python app.py
