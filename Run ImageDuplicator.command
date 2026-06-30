#!/bin/bash
# Double-click this file to launch ImageDuplicator.
# On first run it will install required packages (takes ~30 seconds).
# On later runs it starts up immediately.

# Always run from the folder this script lives in.
cd "$(dirname "$0")"

echo "================================================"
echo "  ImageDuplicator"
echo "================================================"
echo ""

# ── 1. Locate Python 3.11+ ───────────────────────────────────────────────────
PYTHON=""
for candidate in python3.13 python3.12 python3.11 python3 python; do
    if command -v "$candidate" &>/dev/null; then
        ver=$("$candidate" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null)
        major=$(echo "$ver" | cut -d. -f1)
        minor=$(echo "$ver" | cut -d. -f2)
        if [ "$major" -ge 3 ] && [ "$minor" -ge 11 ]; then
            PYTHON="$candidate"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "[X] Python 3.11 or newer is required but was not found."
    echo ""
    echo "    Download it from: https://www.python.org/downloads/"
    echo ""
    read -rp "Press Enter to close..."
    exit 1
fi

echo "[+] Python $("$PYTHON" --version 2>&1 | grep -oE '[0-9]+\.[0-9]+\.[0-9]+')"

# ── 2. Create virtual environment on first run ────────────────────────────────
VENV=".venv"
if [ ! -d "$VENV" ]; then
    echo "[+] First run — creating virtual environment..."
    "$PYTHON" -m venv "$VENV"
    if [ $? -ne 0 ]; then
        echo ""
        echo "[X] Could not create virtual environment."
        read -rp "Press Enter to close..."
        exit 1
    fi
fi

# ── 3. Activate virtual environment ──────────────────────────────────────────
source "$VENV/bin/activate"

# ── 4. Verify pip is available ───────────────────────────────────────────────
if ! pip --version &>/dev/null; then
    echo "[!] pip not found, attempting to bootstrap..."
    python -m ensurepip --upgrade
    if ! pip --version &>/dev/null; then
        echo ""
        echo "[X] Could not install pip."
        echo "    On Mac:   reinstall Python from https://www.python.org/downloads/"
        echo "    On Linux: run 'sudo apt install python3-pip'  (Ubuntu/Debian)"
        echo "              or  'sudo dnf install python3-pip'  (Fedora/RHEL)"
        echo ""
        read -rp "Press Enter to close..."
        exit 1
    fi
    echo "[+] pip installed successfully."
fi

# ── 5. Install / verify dependencies ─────────────────────────────────────────
echo "[+] Checking dependencies..."
pip install --quiet flask colorama filetype pandas openpyxl
if [ $? -ne 0 ]; then
    echo ""
    echo "[X] Failed to install dependencies. Check your internet connection and try again."
    read -rp "Press Enter to close..."
    exit 1
fi

echo "[+] All dependencies ready."
echo ""

# ── 6. Launch the app ────────────────────────────────────────────────────────
echo "    Your browser will open automatically."
echo "    Close this window (or press Ctrl+C) to stop the program."
echo ""
python app.py
EXIT_CODE=$?

# Exit code 130 = Ctrl+C (normal shutdown) — don't treat that as an error.
if [ $EXIT_CODE -ne 0 ] && [ $EXIT_CODE -ne 130 ]; then
    echo ""
    echo "[X] The program exited unexpectedly (exit code $EXIT_CODE)."
    echo "    Scroll up to read the error, then contact your administrator."
    echo ""
    read -rp "Press Enter to close..."
fi
