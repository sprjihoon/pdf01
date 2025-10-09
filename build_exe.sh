#!/bin/bash

echo "========================================"
echo "PDF Excel Matcher - EXE 빌드 스크립트"
echo "========================================"
echo ""

# 가상환경 확인
if [ ! -f ".venv/bin/activate" ]; then
    echo "[오류] 가상환경이 없습니다. 먼저 가상환경을 생성하세요:"
    echo "  python -m venv .venv"
    echo "  source .venv/bin/activate"
    echo "  pip install -r requirements.txt"
    exit 1
fi

# 가상환경 활성화
echo "[1/4] 가상환경 활성화 중..."
source .venv/bin/activate

# PyInstaller 설치 확인
echo "[2/4] PyInstaller 확인 중..."
if ! pip show pyinstaller > /dev/null 2>&1; then
    echo "PyInstaller가 설치되어 있지 않습니다. 설치 중..."
    pip install pyinstaller
fi

# 이전 빌드 삭제
echo "[3/4] 이전 빌드 삭제 중..."
rm -rf build dist *.spec

# PyInstaller 실행
echo "[4/4] 실행 파일 빌드 중..."
echo "잠시만 기다려 주세요... (수 분 소요될 수 있습니다)"
echo ""

pyinstaller --name=PDFExcelMatcher \
    --onedir \
    --windowed \
    --icon=NONE \
    --hidden-import=PySide6 \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=pdfplumber \
    --hidden-import=pypdf \
    --hidden-import=rapidfuzz \
    --collect-all PySide6 \
    --collect-all pdfplumber \
    main.py

if [ $? -ne 0 ]; then
    echo ""
    echo "[오류] 빌드 중 오류가 발생했습니다."
    exit 1
fi

echo ""
echo "========================================"
echo "✅ 빌드 완료!"
echo "========================================"
echo ""
echo "실행 파일 위치: dist/PDFExcelMatcher/PDFExcelMatcher"
echo ""
echo "배포 방법:"
echo "  1. dist/PDFExcelMatcher 폴더 전체를 복사"
echo "  2. 폴더 안의 PDFExcelMatcher를 실행"
echo ""

