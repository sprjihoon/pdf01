@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Excel Matcher - Build EXE
echo ================================================================
echo.

:: Install PyInstaller
echo [1/3] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

:: Clean previous build
echo [2/3] Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

:: Build EXE
echo [3/3] Building EXE file... (3-5 minutes)
echo.

pyinstaller --onefile --windowed --name PDFMatcher main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed
    pause
    exit /b 1
)

echo.
echo ================================================================
echo Done!
echo ================================================================
echo.
echo EXE Location: dist\PDFMatcher.exe
echo.
echo To use on other PC:
echo    1. Copy dist\PDFMatcher.exe file
echo    2. Double-click to run!
echo    3. No Python needed!
echo.
pause

