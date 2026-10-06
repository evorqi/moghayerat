@echo off
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  where python >nul 2>&1
  if %errorlevel%==0 (
    set "PY=python"
  ) else (
    echo Python 3 was not found. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run this file again.
    pause
    exit /b 1
  )
)
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
  if errorlevel 1 (
    echo Could not create the virtual environment.
    pause
    exit /b 1
  )
  .venv\Scripts\python.exe -m pip install -r requirements.txt
)
echo Open http://127.0.0.1:8000
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:8000"
.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
