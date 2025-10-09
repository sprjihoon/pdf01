@echo off
chcp 65001 >nul
echo ================================================================
echo PDF Excel Matcher - 빠른 시작 스크립트
echo ================================================================
echo.

:: 가상환경 확인 및 생성
if not exist ".venv\Scripts\activate.bat" (
    echo [1/3] 가상환경 생성 중...
    python -m venv .venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패. Python이 설치되어 있는지 확인하세요.
        pause
        exit /b 1
    )
) else (
    echo [1/3] 가상환경이 이미 존재합니다.
)

:: 가상환경 활성화
echo [2/3] 가상환경 활성화 중...
call .venv\Scripts\activate.bat

:: 패키지 설치 확인
echo [3/3] 필요한 패키지 설치 확인 중...
pip show PySide6 >nul 2>&1
if errorlevel 1 (
    echo 패키지를 설치합니다... (최초 1회, 2-5분 소요)
    pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
        pause
        exit /b 1
    )
) else (
    echo 모든 패키지가 이미 설치되어 있습니다.
)

echo.
echo ================================================================
echo ✅ 설정 완료! 프로그램을 실행합니다...
echo ================================================================
echo.

:: 프로그램 실행
python main.py

pause

