@echo off
chcp 65001 >nul
echo ================================================================
echo GitHub 업로드 (https://github.com/sprjihoon/pdf01)
================================================================
echo.

cd /d "%~dp0"

:: 홈 디렉토리의 .git 삭제 (문제 해결)
if exist "C:\Users\user\.git" (
    echo 상위 폴더의 Git 저장소를 정리합니다...
    rmdir /s /q "C:\Users\user\.git"
)

:: 현재 폴더에 Git 초기화
echo [1/7] Git 저장소 초기화...
git init

:: README.md가 이미 생성됨

:: .gitignore 생성
echo [2/7] .gitignore 생성...
(
echo # Python
echo __pycache__/
echo *.pyc
echo .venv/
echo venv/
echo.
echo # 결과 파일
echo ordered_*.pdf
echo *_match_report.csv
echo *.pdf
echo !example*.pdf
echo *.xlsx
echo *.xls
echo !example*.xls*
) > .gitignore

:: 파일 추가
echo [3/7] 파일 추가 중...
git add main.py
git add matcher.py
git add io_utils.py
git add requirements.txt
git add README.md
git add README_v2.md
git add START_v2.txt
git add build.bat
git add .gitignore

:: Git 사용자 설정 (필요시)
git config user.name "sprjihoon" 2>nul
git config user.email "sprjihoon@users.noreply.github.com" 2>nul

:: 커밋
echo [4/7] 커밋 중...
git commit -m "Initial commit: PDF Excel Matcher v2.0

- 4요소 매칭 (이름+전화번호+주소+주문번호)
- 전화번호 앞자리 0 누락 자동 복구
- 파일명 버전 관리
- CSV 리포트 생성"

:: 브랜치 이름 변경
echo [5/7] 브랜치를 main으로 변경...
git branch -M main

:: 원격 저장소 추가
echo [6/7] GitHub 저장소 연결...
git remote remove origin 2>nul
git remote add origin https://github.com/sprjihoon/pdf01.git

:: 업로드
echo [7/7] GitHub에 업로드 중...
echo.
echo ⚠️  GitHub 로그인 창이 나타날 수 있습니다.
echo.
git push -u origin main

if errorlevel 1 (
    echo.
    echo ================================================================
    echo ❌ 업로드 실패
    echo ================================================================
    echo.
    echo 다음을 시도해보세요:
    echo 1. GitHub Personal Access Token 생성
    echo    https://github.com/settings/tokens
    echo.
    echo 2. 다음 명령 실행:
    echo    git push -u origin main
    echo.
    echo 3. Username에 "sprjihoon" 입력
    echo 4. Password에 생성한 Token 입력
    echo.
) else (
    echo.
    echo ================================================================
    echo ✅ 업로드 완료!
    echo ================================================================
    echo.
    echo GitHub 저장소: https://github.com/sprjihoon/pdf01
    echo.
    echo 다른 PC에서 다운로드:
    echo   git clone https://github.com/sprjihoon/pdf01.git
    echo.
)

pause

