@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Excel Matcher - 실행 파일 만들기
echo ================================================================
echo.

:: PyInstaller 설치 확인
echo [1/3] PyInstaller 설치 확인...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller 설치 중...
    pip install pyinstaller
)

:: 이전 빌드 삭제
echo [2/3] 이전 빌드 삭제...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

:: PyInstaller 실행 (단일 파일)
echo [3/3] EXE 파일 생성 중... (3~5분 소요)
echo.

pyinstaller --onefile --windowed --name PDFMatcher main.py

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo.
echo ================================================================
echo ✅ 완료!
echo ================================================================
echo.
echo 📁 실행 파일 위치: dist\PDFMatcher.exe
echo.
echo 🚀 다른 PC에서 사용하기:
echo    1. dist\PDFMatcher.exe 파일만 복사
echo    2. 더블클릭으로 실행!
echo    3. Python 설치 필요 없음!
echo.
pause

