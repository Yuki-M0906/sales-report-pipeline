@echo off
REM Start the GUI with your own settings (pipeline_config.json in this folder).
pushd "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 (
  start "" pythonw run_gui.py
  goto done
)
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw -3 run_gui.py
  goto done
)
python run_gui.py
if errorlevel 1 (
  echo.
  echo The GUI could not be started.
  echo If Python is not installed, install it first, then run setup.bat in this folder.
  pause
)
:done
popd
