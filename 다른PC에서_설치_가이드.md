# 다른 PC에서 PDF 매칭 프로그램 설치 가이드

## 📋 필요한 것
1. Python 3.8 이상
2. Git (선택사항)
3. 인터넷 연결

---

## 🚀 설치 방법

### 방법 1: Git 사용 (추천)

1. **Git 설치** (이미 설치되어 있다면 건너뛰기)
   - https://git-scm.com/download/win 에서 다운로드

2. **프로젝트 다운로드**
   ```bash
   git clone https://github.com/sprjihoon/pdf01.git
   cd pdf01
   ```

3. **필요한 패키지 설치**
   ```bash
   pip install -r requirements.txt
   ```

4. **프로그램 실행**
   ```bash
   python main.py
   ```

---

### 방법 2: ZIP 파일 다운로드

1. **GitHub에서 다운로드**
   - https://github.com/sprjihoon/pdf01
   - 초록색 "Code" 버튼 클릭
   - "Download ZIP" 클릭
   - 압축 해제

2. **명령 프롬프트 열기**
   - 압축 해제한 폴더에서 Shift + 우클릭
   - "여기서 PowerShell 창 열기" 선택

3. **필요한 패키지 설치**
   ```bash
   pip install -r requirements.txt
   ```

4. **프로그램 실행**
   ```bash
   python main.py
   ```

---

## 📦 설치되는 패키지

- **PySide6**: GUI 프로그램
- **pandas**: 엑셀 파일 처리
- **openpyxl**: 엑셀 읽기/쓰기
- **pdfplumber**: PDF 텍스트 추출
- **pypdf**: PDF 페이지 재정렬
- **rapidfuzz**: 유사도 매칭

---

## 💡 사용 방법

1. 프로그램 실행 (`python main.py`)
2. **엑셀 파일** 선택 (구매자명, 전화번호, 주소, 주문번호 컬럼 필요)
3. **PDF 파일** 선택 (텍스트 기반 PDF)
4. **출력 폴더** 선택
5. **유사도 매칭** 체크 확인 (기본값 ON)
6. **"PDF 정렬 실행"** 버튼 클릭

---

## 🎯 지원하는 기능

### 이름 매칭
✅ 괄호 안 이름: `임재숙(양성희)`, `이순자[이화순]`  
✅ 공백 구분 이름: `윤익주 김민정`  
✅ 영문 대문자: `LI JINGSHI`, `BAI FENGJIU`  
✅ 외자 이름: `셈`  
✅ 긴 이름: `이경애본순박영아` (최대 10자)  
✅ 숫자 포함: `본순박0` → `본순박`도 매칭  

### 전화번호
✅ 7~11자리 정규화 지원  
✅ 키워드 기반 추출: "전화", "연락", "TEL" 등  
✅ 자동 보정:
   - 8자리: `27357395` → `01027357395`
   - 9자리: `108302565` → `0108302565`
   - 10자리: `1026417075` → `01026417075`

---

## ⚠️ 주의사항

1. **텍스트 기반 PDF만** 지원 (이미지 PDF는 불가)
2. 엑셀 파일에 **필수 컬럼**: 구매자명, 전화번호, 주소, 주문번호
3. 매칭은 **4가지 정보 모두 확인** (이름+전화+주소+주문번호)

---

## 🆘 문제 해결

### Python이 없다고 나옴
- Python 설치: https://www.python.org/downloads/
- 설치 시 "Add Python to PATH" 체크!

### pip 명령어가 안됨
```bash
python -m pip install -r requirements.txt
```

### 패키지 설치 오류
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 📞 문의
- GitHub: https://github.com/sprjihoon/pdf01
- Issues: https://github.com/sprjihoon/pdf01/issues

---

**최종 업데이트**: 2025-10-10

