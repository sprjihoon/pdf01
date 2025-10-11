@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Matcher - EXE Build
echo ================================================================
echo.

:: PyInstaller check
echo [1/4] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
) else (
    echo PyInstaller already installed
)
echo.

:: Clean previous build
echo [2/4] Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
echo Clean complete
echo.

:: Build EXE
echo [3/4] Building EXE file... (3-5 minutes)
echo Please wait...
echo.

pyinstaller PDFMatcher.spec

if errorlevel 1 (
    echo.
    echo Build failed
    echo Please check errors and try again.
    pause
    exit /b 1
)

echo Build complete
echo.

:: Check result
echo [4/4] Checking build result...
if exist "dist\PDF정렬프로그램.exe" (
    echo.
    echo ================================================================
    echo SUCCESS! EXE file created!
    echo ================================================================
    echo.
    echo File location: dist\PDF정렬프로그램.exe
    echo.
    echo How to use:
    echo    1. Copy dist\PDF정렬프로그램.exe file
    echo    2. Paste to desired folder
    echo    3. Double-click to run!
    echo    4. No Python needed!
    echo.
    echo For distribution:
    echo    - Just copy PDF정렬프로그램.exe file
    echo    - Works on Windows 10/11
    echo    - No installation required
    echo.
    
    :: Open dist folder
    echo Open dist folder? (Y/N)
    choice /c YN /n /m "Choice: "
    if errorlevel 2 goto :end
    if errorlevel 1 explorer dist
) else (
    echo.
    echo EXE file creation failed
    echo Cannot find dist\PDF정렬프로그램.exe file.
)

:end
echo.
pause

