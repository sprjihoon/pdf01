@echo off
chcp 65001 >nul
echo ================================================================
echo GitHub에 프로젝트 업로드
echo ================================================================
echo.

cd /d "%~dp0"

:: Git 저장소 초기화 (폴더 내에서만)
if not exist ".git" (
    echo [1/6] Git 저장소 초기화...
    git init
) else (
    echo [1/6] Git 저장소가 이미 존재합니다.
)

:: README 생성
echo [2/6] README.md 생성...
(
echo # PDF Excel Matcher v2.0
echo.
echo 엑셀 구매자 순서대로 PDF 페이지를 자동 정렬하는 프로그램
echo.
echo ## 주요 기능
echo - 4요소 매칭: 이름 + 전화번호 + 주소 + 주문번호
echo - 스마트 정규화: 전화번호 앞자리 0 누락 자동 복구
echo - 버전 관리: ordered_YYYYMMDD_v2.pdf 자동 생성
echo - CSV 리포트: 상세한 매칭 결과
echo.
echo ## 설치
echo ```bash
echo pip install -r requirements.txt
echo ```
echo.
echo ## 실행
echo ```bash
echo python main.py
echo ```
echo.
echo ## 파일 구조
echo - main.py: GUI 애플리케이션
echo - matcher.py: 매칭 로직
echo - io_utils.py: 파일 처리
echo.
echo 자세한 내용은 README_v2.md 참고
) > README.md

:: .gitignore 확인
echo [3/6] .gitignore 확인...
if not exist ".gitignore" (
    echo .gitignore 생성...
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
    echo.
    echo # 개발 도구
    echo .vscode/
    echo .idea/
    ) > .gitignore
)

:: 파일 추가
echo [4/6] 파일 추가...
git add main.py matcher.py io_utils.py requirements.txt
git add README.md README_v2.md START_v2.txt
git add build.bat PDF정렬_실행.bat
git add .gitignore

:: 커밋
echo [5/6] 커밋...
git commit -m "Initial commit: PDF Excel Matcher v2.0 - 4요소 매칭 (이름+전화+주소+주문번호)"

:: 원격 저장소 설정
echo [6/6] GitHub 저장소 연결...
git branch -M main
git remote remove origin 2>nul
git remote add origin https://github.com/sprjihoon/pdf01.git

echo.
echo ================================================================
echo ✅ GitHub 업로드 준비 완료!
echo ================================================================
echo.
echo 이제 다음 명령으로 업로드하세요:
echo.
echo   git push -u origin main
echo.
echo GitHub 로그인이 필요합니다.
echo.
set /p PROCEED="지금 업로드하시겠습니까? (Y/N): "
if /i "%PROCEED%"=="Y" (
    echo.
    echo 업로드 중...
    git push -u origin main
    echo.
    echo ================================================================
    echo ✅ 업로드 완료!
    echo ================================================================
    echo.
    echo GitHub: https://github.com/sprjihoon/pdf01
) else (
    echo.
    echo 나중에 직접 업로드하려면:
    echo   git push -u origin main
)

echo.
pause

