@echo off
REM Put the sample data back to its initial state (everything under sample\ is rebuilt).
pushd "%~dp0"
python tools\make_sample_data.py
popd
pause
