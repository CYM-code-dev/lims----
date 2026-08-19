@echo off
rem One-time server setup: register auto-start windowless file service on port 8010
rem Usage: copy this file INTO the publish folder on the server, right-click "Run as administrator".
setlocal
set "DIR=%~dp0"
if "%DIR:~-1%"=="\" set "DIR=%DIR:~0,-1%"

where pythonw.exe >nul 2>&1
if errorlevel 1 (
  echo [ERROR] pythonw.exe not found. Install Python and add it to PATH first.
  pause
  exit /b 1
)
for /f "delims=" %%i in ('where pythonw.exe') do set "PYW=%%i"

schtasks /Create /F /TN "LIMS Update Service" /TR "\"%PYW%\" \"%DIR%\server_serve.py\"" /SC ONSTART /RU SYSTEM
netsh advfirewall firewall delete rule name="LIMS Update Service" >nul 2>&1
netsh advfirewall firewall add rule name="LIMS Update Service" dir=in action=allow protocol=TCP localport=8010

rem Idempotent re-run: kill whatever still listens on 8010 (e.g. a stale instance), then start fresh.
schtasks /End /TN "LIMS Update Service" >nul 2>&1
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8010 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
schtasks /Run /TN "LIMS Update Service"

echo.
echo DONE. Serving port 8010 from:
echo   %DIR%
echo Runs at boot as SYSTEM (no window). Reboot-safe.
echo.
echo Remove: schtasks /Delete /TN "LIMS Update Service" /F
pause
