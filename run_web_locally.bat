@echo off
setlocal
cd /d "%~dp0"

REM Runs the web version of the report generator on THIS computer only
REM (http://127.0.0.1:5000), for testing before it goes online.
REM It has its own private venv here in the webapp folder, separate from
REM the one run_weekly_report.bat uses.

if not exist venv (
    echo First-time setup - this only happens once, give it a minute...
    python -m venv venv
    if errorlevel 1 (
        echo Could not find Python. Install it from python.org/downloads first.
        pause
        exit /b 1
    )
)
venv\Scripts\python.exe -m pip install --quiet -r requirements.txt

REM Local test password - the real one is set in Render, not here.
set APP_PASSWORD=test

echo.
echo Starting the web app. Open http://127.0.0.1:5000 in your browser
echo and log in with the password:  test
echo Close this window to stop it.
echo.
start "" http://127.0.0.1:5000
venv\Scripts\python.exe app.py
pause
