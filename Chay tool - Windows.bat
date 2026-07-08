@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Mini Frog - Tool quet Web URL

where python >nul 2>nul
if errorlevel 1 (
  echo ============================================================
  echo  Chua cai Python. Trinh duyet se mo trang tai Python.
  echo  LUU Y: khi cai, nho tick o "Add python.exe to PATH"
  echo  Cai xong, chay lai file nay.
  echo ============================================================
  start https://www.python.org/downloads/
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Lan dau chay: dang cai dat moi truong ^(1-2 phut, can mang^)...
  python -m venv .venv
  .venv\Scripts\python -m pip install -q --upgrade pip
  .venv\Scripts\pip install -q -r requirements.txt
  if errorlevel 1 (
    echo Cai thu vien that bai. Kiem tra mang roi chay lai.
    pause
    exit /b 1
  )
  echo Cai dat xong!
)

echo Dang khoi dong Mini Frog...
start /b .venv\Scripts\python -m uvicorn app:app --port 8765
timeout /t 3 >nul
start "" http://localhost:8765
echo.
echo Tool dang chay tai http://localhost:8765
echo DONG CUA SO NAY DE TAT TOOL.
pause >nul
