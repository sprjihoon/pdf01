@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Excel Matcher - EXE 빌드 스크립트
echo ================================================================
echo.

:: 가상환경 확인
if not exist ".venv\Scripts\activate.bat" (
    echo [오류] 가상환경이 없습니다. 먼저 가상환경을 생성하세요:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

:: 가상환경 활성화
echo [1/4] 가상환경 활성화 중...
call .venv\Scripts\activate.bat

:: PyInstaller 확인
echo [2/4] PyInstaller 확인 중...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다. 설치 중...
    pip install pyinstaller
)

:: 이전 빌드 삭제
echo [3/4] 이전 빌드 삭제 중...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

:: PyInstaller 실행
echo [4/4] EXE 파일 빌드 중...
echo 잠시만 기다려 주세요... (수 분 소요될 수 있습니다)
echo.

pyinstaller --onefile ^
    --noconsole ^
    --name PDFReorderApp ^
    --add-data ".venv\Lib\site-packages\pdfplumber;pdfplumber" ^
    --hidden-import=PySide6 ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=xlrd ^
    --hidden-import=pdfplumber ^
    --hidden-import=pypdf ^
    --hidden-import=rapidfuzz ^
    --collect-all PySide6 ^
    --collect-all pdfplumber ^
    main.py

if errorlevel 1 (
    echo.
    echo [오류] 빌드 중 오류가 발생했습니다.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo ✅ 빌드 완료!
echo ================================================================
echo.
echo 실행 파일 위치: dist\PDFReorderApp.exe
echo.
echo 배포 방법:
echo   1. dist\PDFReorderApp.exe 파일을 복사
echo   2. 사용자 PC에서 바로 실행 가능!
echo.
pause

