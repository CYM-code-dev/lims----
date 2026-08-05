@echo off
setlocal enabledelayedexpansion
(for %%a in (*_*_*.pdf) do (
    set "n=%%~na"
    set "part3=!n!"
    set "part3=!part3:*_=!"
    set "part3=!part3:*_=!"
    ren "%%a" "!part3!%%~xa" && echo ÒÑÖØÃüÃû£º%%a ¡ú !part3!%%~xa
)) > rename_log.txt
exit