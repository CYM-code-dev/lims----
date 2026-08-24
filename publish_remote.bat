@echo off
REM ============================================================
REM  LIMS data-entry remote publish. Run from the dev machine.
REM  1) build release (publish.py, PyInstaller x2)
REM  2) scp dist/publish to the server folder served on port 8010
REM  3) curl the served manifest to verify
REM  One-time server prep (after first upload, brings 8010 online):
REM    ssh Administrator@10.1.93.25 "cmd /c E:\server\publish\server_setup.bat < NUL"
REM  Clients auto-update from manifest.json on next launcher start.
REM ============================================================
setlocal
title LIMS remote publish

set "SRV=10.1.93.25"
set "SUSER=Administrator"
set "DEST=%SUSER%@%SRV%:E:/server/publish/"

echo === [1/3] build release (PyInstaller x2, takes a few minutes) ===
if not exist ".venv\Scripts\python.exe" ( echo [X] .venv missing & pause & exit /b 1 )
.venv\Scripts\python.exe publish.py
if errorlevel 1 ( echo [X] publish.py FAILED & pause & exit /b 1 )

echo === [2/3] upload to server E:\server\publish (folder served on 8010) ===
scp dist\publish\launcher.exe dist\publish\server_setup.bat dist\publish\server_serve.py dist\publish\app-*.zip dist\publish\manifest.json %DEST%
if errorlevel 1 ( echo [X] scp FAILED & pause & exit /b 1 )

echo === [3/3] served manifest (empty means 8010 not running) ===
curl -s -m 5 http://%SRV%:8010/manifest.json
echo.
pause
endlocal
