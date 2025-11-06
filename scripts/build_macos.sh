#!/usr/bin/env bash
set -euo pipefail

# Build a macOS app bundle using PyInstaller
# Output: dist/SummaryQuickMerge.app

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$ROOT_DIR"

PY="python3"

if ! command -v "$PY" >/dev/null 2>&1; then
  echo "python3 not found. Install Python 3.10+ first." >&2
  exit 1
fi

# Ensure dependencies
"$PY" -m pip install --upgrade pip
"$PY" -m pip install -r requirements.txt

# Clean previous builds
rm -rf build dist "SummaryQuickMerge.spec" || true

# Build app bundle (prefer spec if present for correct data collection)
if [[ -f "SummaryQuickMerge.spec" ]]; then
  "$PY" -m PyInstaller SummaryQuickMerge.spec
else
  "$PY" -m PyInstaller \
    --windowed \
    --name "SummaryQuickMerge" \
    --collect-data docx \
    --collect-data docxcompose \
    --hidden-import lxml.etree \
    --hidden-import lxml._elementpath \
    gui.py
fi

echo "\nBuilt: dist/SummaryQuickMerge.app"

