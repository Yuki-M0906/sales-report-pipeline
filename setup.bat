@echo off
REM Install the Python libraries this tool needs (run once per PC).
set "PY="
python -c "import sys" >nul 2>nul
if %errorlevel%==0 set "PY=python"
if not defined PY (
  py -3 -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PY=py -3"
)
if not defined PY (
  echo Python was not found. Please install Python 3.10 or later first, then run this file again.
  pause
  exit /b 1
)
echo ============================================================
echo  Installing required libraries (see requirements.txt)
echo ============================================================
%PY% -m pip install -r "%~dp0requirements.txt"
set RC=%errorlevel%
echo.
if %RC%==0 (
  echo Setup finished. You can now start the GUI with run_sample.bat or run_gui.bat.
) else (
  echo Setup FAILED. Please check the messages above.
)
pause
