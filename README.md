# PDF Excel Matcher v2.0

엑셀 구매자 순서대로 PDF 페이지를 자동 정렬하는 프로그램

## 🎯 주요 기능

- **4요소 동시 매칭**: 이름 + 전화번호 + 주소 + 주문번호
- **스마트 정규화**: 전화번호 앞자리 0 누락 자동 복구
- **버전 관리**: `ordered_YYYYMMDD_v2.pdf` 자동 생성
- **CSV 리포트**: 상세한 매칭 결과 제공

## 📦 설치

```bash
pip install -r requirements.txt
```

## 🚀 실행

```bash
python main.py
```

또는 `PDF정렬_실행.bat` 더블클릭!

## 📁 파일 구조

- `main.py`: PySide6 GUI 애플리케이션
- `matcher.py`: 텍스트 추출 및 매칭 로직
- `io_utils.py`: 파일 입출력 및 버전 관리

## 📖 상세 문서

자세한 사용 방법은 `README_v2.md`와 `START_v2.txt`를 참고하세요.

## 💡 매칭 로직

1. **엑셀 데이터 정규화**
   - 전화번호: `1026417075` → `01026417075` (앞 0 추가)
   - 이름: 공백 제거, 대문자 변환
   - 주소: 특수문자 제거
   - 주문번호: 영문+숫자만 추출

2. **PDF 텍스트 추출**
   - pdfplumber로 각 페이지 텍스트 추출
   - 정규표현식으로 정보 추출

3. **4요소 매칭**
   - 이름, 전화번호, 주소, 주문번호 모두 일치해야 매칭
   - 같은 사람의 여러 주문도 정확히 구분

4. **페이지 정렬**
   - 매칭된 페이지: 엑셀 순서대로
   - 미매칭 페이지: 원본 순서로 뒤에 배치

## 🔧 기술 스택

- Python 3.11
- PySide6 (GUI)
- pandas, openpyxl (Excel)
- pdfplumber, pypdf (PDF)
- rapidfuzz (유사도 매칭)

## 📝 라이선스

MIT License

## 👨‍💻 개발

Cursor AI로 개발된 프로젝트입니다.
