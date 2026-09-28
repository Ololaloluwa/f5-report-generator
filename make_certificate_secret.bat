@echo off
setlocal
cd /d "%~dp0"

REM Makes Certificate.xlsx.b64.txt from the Certificate.xlsx in the folder
REM above, and opens it in Notepad: select all, copy, and paste it into
REM Render as the Secret File named  Certificate.xlsx.b64

set PY=python
if exist venv\Scripts\python.exe set PY=venv\Scripts\python.exe
%PY% encode_certificate.py
if errorlevel 1 (
    pause
    exit /b 1
)
notepad Certificate.xlsx.b64.txt
