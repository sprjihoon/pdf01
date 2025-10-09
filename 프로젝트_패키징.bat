@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Excel Matcher - 프로젝트 패키징
echo ================================================================
echo.
echo 다른 PC로 이전할 수 있도록 필요한 파일만 압축합니다.
echo.

cd /d "%~dp0"

:: 임시 폴더 생성
set PACKAGE_NAME=PDFExcelMatcher_Package
if exist "%PACKAGE_NAME%" rmdir /s /q "%PACKAGE_NAME%"
mkdir "%PACKAGE_NAME%"

echo [1/3] 소스 파일 복사 중...
copy main.py "%PACKAGE_NAME%\" >nul
copy matcher.py "%PACKAGE_NAME%\" >nul
copy io_utils.py "%PACKAGE_NAME%\" >nul
copy requirements.txt "%PACKAGE_NAME%\" >nul

echo [2/3] 스크립트 파일 복사 중...
copy build.bat "%PACKAGE_NAME%\" >nul
copy PDF정렬_실행.bat "%PACKAGE_NAME%\" >nul

echo [3/3] 문서 파일 복사 중...
copy README_v2.md "%PACKAGE_NAME%\" >nul
copy START_v2.txt "%PACKAGE_NAME%\" >nul
copy .gitignore "%PACKAGE_NAME%\" >nul

:: 설치 가이드 생성
echo ================================================================ > "%PACKAGE_NAME%\00_설치방법.txt"
echo PDF Excel Matcher - 설치 가이드 >> "%PACKAGE_NAME%\00_설치방법.txt"
echo ================================================================ >> "%PACKAGE_NAME%\00_설치방법.txt"
echo. >> "%PACKAGE_NAME%\00_설치방법.txt"
echo 1단계: Python 설치 확인 >> "%PACKAGE_NAME%\00_설치방법.txt"
echo   python --version >> "%PACKAGE_NAME%\00_설치방법.txt"
echo   (Python 3.8 이상 필요) >> "%PACKAGE_NAME%\00_설치방법.txt"
echo. >> "%PACKAGE_NAME%\00_설치방법.txt"
echo 2단계: 패키지 설치 >> "%PACKAGE_NAME%\00_설치방법.txt"
echo   pip install -r requirements.txt >> "%PACKAGE_NAME%\00_설치방법.txt"
echo. >> "%PACKAGE_NAME%\00_설치방법.txt"
echo 3단계: 실행 >> "%PACKAGE_NAME%\00_설치방법.txt"
echo   python main.py >> "%PACKAGE_NAME%\00_설치방법.txt"
echo. >> "%PACKAGE_NAME%\00_설치방법.txt"
echo 또는 PDF정렬_실행.bat 더블클릭! >> "%PACKAGE_NAME%\00_설치방법.txt"
echo. >> "%PACKAGE_NAME%\00_설치방법.txt"
echo ================================================================ >> "%PACKAGE_NAME%\00_설치방법.txt"

echo.
echo ================================================================
echo ✅ 패키징 완료!
echo ================================================================
echo.
echo 생성된 폴더: %PACKAGE_NAME%
echo.
echo 다음 단계:
echo   1. "%PACKAGE_NAME%" 폴더를 압축 (ZIP)
echo   2. 다른 PC로 복사
echo   3. 압축 해제 후 00_설치방법.txt 참고
echo.
echo 또는 "%PACKAGE_NAME%" 폴더를 USB/클라우드로 복사하세요!
echo.
pause

