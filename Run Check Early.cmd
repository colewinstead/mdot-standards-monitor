@echo off
setlocal
title MDOT Standards Monitor - Early Check
cd /d "%~dp0"

echo MDOT Standards Monitor
echo Running an early check...
echo.

if not exist "%~dp0monitor.py" (
    echo ERROR: monitor.py was not found in:
    echo %~dp0
    goto :finish
)

where py.exe >nul 2>&1
if not errorlevel 1 (
    py.exe -3.14 "%~dp0monitor.py" check
    set "MONITOR_EXIT=%ERRORLEVEL%"
    goto :result
)

where python.exe >nul 2>&1
if not errorlevel 1 (
    python.exe "%~dp0monitor.py" check
    set "MONITOR_EXIT=%ERRORLEVEL%"
    goto :result
)

echo ERROR: Python was not found. Install Python 3.11 or newer.
goto :finish

:result
echo.
if "%MONITOR_EXIT%"=="0" (
    echo Check completed successfully.
) else (
    echo Check ended with error code %MONITOR_EXIT%.
    echo Review the message above and the monitor log for details.
)

:finish
echo.
pause
endlocal
