@echo off
REM Start the GUI with the bundled sample data (sample\pipeline_config.json).
pushd "%~dp0"
if not exist "sample\pipeline_config.json" (
  echo Sample data was not found. Creating it...
  python tools\make_sample_data.py
)
if not exist "sample\work" mkdir "sample\work"
set "PIPELINE_CONFIG=%~dp0sample\pipeline_config.json"
set "PIPELINE_PROFILE=SAMPLE"
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
