@echo off
echo ========================================
echo  DataCopilot - Build Script
echo ========================================

:: Clean previous builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist DataCopilot.spec del DataCopilot.spec

echo [1/3] Building exe...
pyinstaller --noconsole ^
            --onefile ^
            --name DataCopilot ^
            --add-data "engine;engine" ^
            main.py

if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo [2/3] Copying model folder...
xcopy /e /i /y model dist\model

echo [3/3] Done!
echo.
echo Output folder: dist\
echo   DataCopilot.exe   - main program
echo   model\            - AI model (required alongside exe)
echo.
echo To distribute: zip the entire dist\ folder.
pause
