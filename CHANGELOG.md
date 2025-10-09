# Changelog

## [1.0.0] - 2025-10-09

### 추가된 기능
- ✨ PySide6 기반 GUI 애플리케이션
- 📄 PDF 페이지 자동 정렬 기능
- 📊 엑셀 파일에서 구매자 정보 읽기
- 🔍 다양한 정규화 타입 지원 (일반, 전화번호, 이름, 주문번호)
- 🎯 유사도 기반 스마트 매칭 (rapidfuzz 사용)
- 📝 상세 매칭 리포트 생성
- 🚀 단일 실행 파일(.exe) 빌드 지원
- 🧪 예시 파일 생성 스크립트
- 📖 상세한 문서 (README, 설치 가이드)

### 주요 기능
- **자동 매칭**: 엑셀 키와 PDF 텍스트를 자동으로 매칭
- **스마트 정렬**: 매칭된 페이지는 엑셀 순서대로, 미매칭은 마지막에 배치
- **유연한 정규화**: 데이터 타입에 맞는 정규화 방식 선택 가능
- **직관적 GUI**: 사용하기 쉬운 인터페이스
- **상세 리포트**: 페이지별 매칭 결과 및 유사도 점수 제공

### 기술 스택
- Python 3.11
- PySide6 (GUI)
- pandas, openpyxl (엑셀 처리)
- pdfplumber, pypdf (PDF 처리)
- rapidfuzz (텍스트 유사도)
- PyInstaller (실행 파일 빌드)

### 파일 구조
```
pdf_excel_matcher/
├── main.py                    # 메인 프로그램
├── requirements.txt           # 패키지 목록
├── build_exe.bat             # Windows 빌드 스크립트
├── build_exe.sh              # Mac/Linux 빌드 스크립트
├── quick_start.bat           # 빠른 시작 스크립트
├── create_example_files.py   # 예시 파일 생성
├── README.md                 # 상세 문서
├── setup_guide.txt           # 설치 가이드
├── CHANGELOG.md              # 변경 이력
└── .gitignore               # Git 제외 파일
```

