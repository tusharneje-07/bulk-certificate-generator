@echo off
setlocal EnableDelayedExpansion

title Bulk Certificate Generator

cd /d "%~dp0"

cls
echo ===============================================
echo         Bulk Certificate Generator
echo ===============================================
echo.

echo [1/5] Checking application...

if not exist "main.py" (
    echo.
    echo Error: main.py not found.
    pause
    exit /b 1
)

echo [2/5] Checking Python...

where py >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=py
) else (
    where python >nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo Error: Python is not installed.
        pause
        exit /b 1
    )
    set PYTHON=python
)

echo [3/5] Checking virtual environment...

if not exist "venv\" (
    %PYTHON% -m venv venv
)

call venv\Scripts\activate.bat

echo [4/5] Checking dependencies...

python -m pip install --upgrade pip --quiet >nul 2>&1

echo Installing dependencies...

python -c ^
"import ast,importlib,subprocess,sys; \
PACKAGE_MAP={'PIL':'Pillow','ttkbootstrap':'ttkbootstrap','openpyxl':'openpyxl','pypdf':'pypdf'}; \
STDLIB={'os','sys','re','io','json','queue','random','shutil','sqlite3','platform','subprocess','threading','traceback','pathlib','datetime','concurrent','tkinter','typing','math','time','itertools','collections','functools','logging'}; \
tree=ast.parse(open('main.py','r',encoding='utf-8').read()); \
imports=set(); \
[imports.add(n.name.split('.')[0]) for node in ast.walk(tree) if isinstance(node,ast.Import) for n in node.names]; \
[imports.add(node.module.split('.')[0]) for node in ast.walk(tree) if isinstance(node,ast.ImportFrom) and node.module]; \
missing=[]; \
[(missing.append(PACKAGE_MAP.get(m,m))) for m in sorted(imports) if m not in STDLIB and (__import__('importlib').import_module(m) if importlib.util.find_spec(m) else missing.append(PACKAGE_MAP.get(m,m)))]; \
missing=list(dict.fromkeys(missing)); \
subprocess.check_call([sys.executable,'-m','pip','install','--quiet',*missing],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL) if missing else None"

if errorlevel 1 (
    echo.
    echo Failed to install required dependencies.
    pause
    exit /b 1
)

echo Done.
echo.

echo [5/5] Launching Bulk Certificate Generator...
echo.

python main.py

pause