@echo off
chcp 65001 >nul
echo ================================================================
echo Git 저장소 초기화
echo ================================================================
echo.

cd /d "%~dp0"

:: Git이 설치되어 있는지 확인
git --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Git이 설치되어 있지 않습니다.
    echo.
    echo Git 다운로드: https://git-scm.com/download/win
    echo.
    pause
    exit /b 1
)

echo ✅ Git이 설치되어 있습니다.
echo.

:: Git 초기화
if not exist ".git" (
    echo [1/4] Git 저장소 초기화...
    git init
) else (
    echo [1/4] Git 저장소가 이미 존재합니다.
)

:: .gitignore 확인
if exist ".gitignore" (
    echo [2/4] .gitignore 파일 확인됨
) else (
    echo [2/4] .gitignore 파일 생성...
    echo # Python > .gitignore
    echo __pycache__/ >> .gitignore
    echo *.pyc >> .gitignore
    echo .venv/ >> .gitignore
    echo venv/ >> .gitignore
    echo *.pdf >> .gitignore
    echo !example*.pdf >> .gitignore
    echo *.xlsx >> .gitignore
    echo *.xls >> .gitignore
    echo !example*.xls* >> .gitignore
    echo ordered_*.pdf >> .gitignore
    echo *_match_report.csv >> .gitignore
)

:: 모든 파일 추가
echo [3/4] 파일 추가...
git add .

:: 커밋
echo [4/4] 커밋...
git commit -m "Initial commit: PDF Excel Matcher v2.0"

echo.
echo ================================================================
echo ✅ Git 저장소 초기화 완료!
echo ================================================================
echo.
echo 다음 단계:
echo.
echo [방법 A] GitHub 사용:
echo   1. GitHub에서 새 저장소 생성
echo   2. git remote add origin [저장소URL]
echo   3. git push -u origin master
echo.
echo [방법 B] 다른 PC로 직접 복사:
echo   1. 이 폴더를 USB/클라우드로 복사
echo   2. 다른 PC에서:
echo      git pull origin master
echo.
pause

