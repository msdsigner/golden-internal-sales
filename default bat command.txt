@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo   Starting Golden Inventory System
echo ============================================================
echo.
echo [EXECUTION] Running Unified Golden Inventory Engine...
python golden_inventory_app.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] The automation encountered a problem.
    echo Check the messages above for details.
    pause
)
echo.
echo ============================================================
echo   PROCESS COMPLETE: Reports Generated ^& Web Data Live
echo ============================================================
pause
