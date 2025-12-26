@echo off

:: 1. Check if we are in an active venv to toggle it off
if defined VIRTUAL_ENV (
    echo [WAIT] Deactivating environment...
    call deactivate
    echo [OFF] Virtual environment closed.
    goto :end
)

:: 2. Search for venv folders
set "found="
for %%d in (venv .venv env) do (
    if exist "%%d\Scripts\activate.bat" (
        set "found=%%d"
        
        :: Check if the script was double-clicked (no parent cmd process)
        :: If double-clicked, we use cmd /k to stay interactive
        echo %cmdcmdline% | findstr /i /c:"%~nx0" >nul
        if not errorlevel 1 (
            echo [OK] Double-click detected. Opening interactive shell...
            start cmd /k "echo [OK] Found: %%d && call %%d\Scripts\activate.bat && echo [READY] Environment active. You can type commands now."
            goto :end
        ) else (
            :: If run from an existing terminal, just activate it
            echo [OK] Activating: %%d
            call "%%d\Scripts\activate.bat"
            goto :end
        )
    )
)

if not defined found (
    echo [ERROR] No venv, .venv, or env folder found.
    pause
)

:end