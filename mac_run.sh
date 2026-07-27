#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ---------------------------------------------
# Detect Python
# ---------------------------------------------
if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_CMD="python"
else
    echo "Python is not installed."
    exit 1
fi

VENV_DIR="venv"
MAIN_FILE="main.py"

command -v clear >/dev/null 2>&1 && clear

echo "==============================================="
echo "        Bulk Certificate Generator"
echo "==============================================="
echo

# -------------------------------------------------
# Progress Bar
# -------------------------------------------------

progress_bar() {

    local pid=$1
    local width=40
    local pos=0
    local dir=1

    tput civis 2>/dev/null || true

    while kill -0 "$pid" 2>/dev/null; do

        printf "\rInstalling Dependencies ["

        for ((i=0;i<width;i++)); do
            if ((i==pos)); then
                printf "#"
            else
                printf "-"
            fi
        done

        printf "]"

        pos=$((pos+dir))

        if ((pos>=width-1)); then
            dir=-1
        elif ((pos<=0)); then
            dir=1
        fi

        sleep 0.05

    done

    wait "$pid"
    status=$?

    printf "\rInstalling Dependencies [########################################] Done\n"

    tput cnorm 2>/dev/null || true

    return $status
}

echo "[1/5] Checking application..."

if [[ ! -f "$MAIN_FILE" ]]; then
    echo
    echo "Error: main.py not found."
    exit 1
fi

echo "[2/5] Checking Python..."

echo "[3/5] Checking virtual environment..."

if [[ ! -d "$VENV_DIR" ]]; then
    "$PYTHON_CMD" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

echo "[4/5] Checking dependencies..."

python -m pip install --upgrade pip --quiet >/dev/null 2>&1 &
progress_bar $!

python - <<'PY' &
import ast
import importlib
import subprocess
import sys

PACKAGE_MAP = {
    "PIL": "Pillow",
    "ttkbootstrap": "ttkbootstrap",
    "openpyxl": "openpyxl",
    "pypdf": "pypdf",
}

STDLIB = {
    "os","sys","re","io","json","queue","random","shutil",
    "sqlite3","platform","subprocess","threading",
    "traceback","pathlib","datetime","concurrent",
    "tkinter","typing","math","time","itertools",
    "collections","functools","logging"
}

with open("main.py","r",encoding="utf-8") as f:
    tree=ast.parse(f.read())

imports=set()

for node in ast.walk(tree):
    if isinstance(node,ast.Import):
        for mod in node.names:
            imports.add(mod.name.split(".")[0])
    elif isinstance(node,ast.ImportFrom):
        if node.module:
            imports.add(node.module.split(".")[0])

missing=[]

for module in sorted(imports):

    if module in STDLIB:
        continue

    try:
        importlib.import_module(module)
    except ImportError:
        missing.append(PACKAGE_MAP.get(module,module))

if missing:
    subprocess.check_call(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--quiet",
            *sorted(set(missing))
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
PY

progress_bar $!

echo
echo "[5/5] Launching Bulk Certificate Generator..."
echo

exec python "$MAIN_FILE"