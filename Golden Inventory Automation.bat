@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo          GOLDEN INVENTORY AUTOMATION
echo ============================================================
echo.

echo [1/2] Syncing inventory from Odoo...
python sync_odoo.py

if errorlevel 1 (
    echo.
    echo Odoo synchronization failed.
    pause
    exit /b
)

echo.
echo [2/2] Running Golden Inventory Engine...
python golden_inventory_app.py

if errorlevel 1 (
    echo.
    echo Inventory processing failed.
    pause
    exit /b
)

echo.
echo ============================================================
echo                AUTOMATION COMPLETE
echo ============================================================
pause