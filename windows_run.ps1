# windows_run.ps1
# ================================================
# Bulk Certificate Generator Launcher
# ================================================

$ErrorActionPreference = "Stop"

Clear-Host

Write-Host "==============================================="
Write-Host "        Bulk Certificate Generator"
Write-Host "==============================================="
Write-Host ""

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$MainFile = "main.py"
$Venv = "venv"

# -------------------------------------------------
# Progress Animation
# -------------------------------------------------

function Show-ProgressWhileRunning {
    param(
        [System.Diagnostics.Process]$Process
    )

    $width = 40
    $pos = 0
    $dir = 1

    while (-not $Process.HasExited) {

        $bar = ""

        for ($i=0; $i -lt $width; $i++) {
            if ($i -eq $pos) {
                $bar += "#"
            }
            else {
                $bar += "-"
            }
        }

        Write-Host -NoNewline "`rInstalling Dependencies [$bar]"

        $pos += $dir

        if ($pos -ge ($width-1)) {
            $dir = -1
        }
        elseif ($pos -le 0) {
            $dir = 1
        }

        Start-Sleep -Milliseconds 50
    }

    Write-Host "`rInstalling Dependencies [########################################] Done"
}

# -------------------------------------------------

Write-Host "[1/5] Checking application..."

if (!(Test-Path $MainFile)) {
    Write-Host ""
    Write-Host "Error: main.py not found."
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host "[2/5] Checking Python..."

$python = Get-Command py -ErrorAction SilentlyContinue

if ($python) {
    $PythonCmd = "py"
}
else {

    $python = Get-Command python -ErrorAction SilentlyContinue

    if (!$python) {
        Write-Host ""
        Write-Host "Python is not installed."
        Read-Host "Press ENTER to exit"
        exit 1
    }

    $PythonCmd = "python"
}

Write-Host "[3/5] Checking virtual environment..."

if (!(Test-Path $Venv)) {
    & $PythonCmd -m venv $Venv
}

$Activate = Join-Path $Root "venv\Scripts\Activate.ps1"

& $Activate

Write-Host "[4/5] Checking dependencies..."

python -m pip install --upgrade pip --quiet *> $null

$TempScript = Join-Path $env:TEMP "bulk_certificate_generator_dependency_check.py"

@'
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
    tree = ast.parse(f.read())

imports=set()

for node in ast.walk(tree):
    if isinstance(node, ast.Import):
        for m in node.names:
            imports.add(m.name.split(".")[0])
    elif isinstance(node, ast.ImportFrom):
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

missing=list(dict.fromkeys(missing))

if missing:
    subprocess.check_call(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--quiet",
            *missing
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
'@ | Set-Content $TempScript -Encoding UTF8

$proc = Start-Process python `
    -ArgumentList "`"$TempScript`"" `
    -WindowStyle Hidden `
    -PassThru

Show-ProgressWhileRunning $proc

$proc.WaitForExit()

Remove-Item $TempScript -Force -ErrorAction SilentlyContinue

if ($proc.ExitCode -ne 0) {
    Write-Host ""
    Write-Host "Failed to install required dependencies."
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host ""
Write-Host "[5/5] Launching Bulk Certificate Generator..."
Write-Host ""

python main.py