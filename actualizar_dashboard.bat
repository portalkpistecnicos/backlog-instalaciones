@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%actualizar_dashboard.ps1"
set RC=%ERRORLEVEL%
if %RC% NEQ 0 (
    echo.
    echo La actualizacion fallo. Revisa la carpeta "logs" para el detalle.
) else (
    echo.
    echo Dashboard actualizado correctamente.
)
exit /b %RC%
