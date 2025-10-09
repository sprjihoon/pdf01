"""
파일 입출력 및 버전 관리 유틸리티
- 엑셀 로드
- 파일명 버전 관리 (ordered_YYYYMMDD_vN.pdf)
- PDF/CSV 저장
"""

import os
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
from pypdf import PdfWriter


def load_excel(path):
    """
    엑셀 파일 로드
    
    Args:
        path: 엑셀 파일 경로
        
    Returns:
        pandas.DataFrame
        
    Raises:
        ValueError: 필수 컬럼이 없는 경우
    """
    # 파일 확장자 확인
    ext = Path(path).suffix.lower()
    if ext not in ['.xls', '.xlsx']:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")
    
    # 엑셀 읽기
    if ext == '.xls':
        df = pd.read_excel(path, engine='xlrd')
    else:
        df = pd.read_excel(path, engine='openpyxl')
    
    # 필수 컬럼 확인 (대소문자 무시)
    required_cols = ['구매자명', '전화번호', '주소', '주문번호']
    df_cols_lower = {col.strip().lower(): col for col in df.columns}
    
    # 컬럼 매핑
    col_mapping = {}
    
    # 다양한 컬럼명 패턴 정의
    column_patterns = {
        '구매자명': [
            'name', 'buyer', 'customer', '이름', '성명',
            '구매자', '수령자명', '수령자', '받는사람', '받는분',
            '고객명', '주문자명', '주문자'
        ],
        '전화번호': [
            'phone', 'tel', 'mobile', '연락처', '핸드폰',
            '전화', '휴대폰', '휴대전화', '휴대폰번호', '전화번호',
            '연락처', 'contact', '수령자전화', '수령자휴대폰',
            '수령자휴대폰번호', '수령자연락처'
        ],
        '주소': [
            'address', 'addr', 'location', '주소',
            '배송지', '배송주소', '수령지', '수령지주소',
            '도착지', '배달주소', '받는주소'
        ],
        '주문번호': [
            'order', 'ordernumber', 'ordernum', 'orderno', 'order_no',
            '주문번호', '주문', '오더', '오더번호', '주문num',
            '계약번호', '거래번호', '구매번호'
        ]
    }
    
    for req_col in required_cols:
        found = False
        
        # 직접 일치 체크
        for df_col_lower, df_col in df_cols_lower.items():
            if req_col.lower() in df_col_lower or df_col_lower in req_col.lower():
                col_mapping[df_col] = req_col
                found = True
                break
        
        if not found:
            # 패턴 매칭
            for df_col_lower, df_col in df_cols_lower.items():
                for pattern in column_patterns[req_col]:
                    # 공백 제거 후 비교
                    df_col_clean = df_col_lower.replace(' ', '').replace('_', '')
                    pattern_clean = pattern.replace(' ', '').replace('_', '')
                    
                    if pattern_clean in df_col_clean or df_col_clean in pattern_clean:
                        col_mapping[df_col] = req_col
                        found = True
                        break
                if found:
                    break
        
        if not found:
            # 사용 가능한 컬럼 목록 표시
            available_cols = ', '.join([f'"{col}"' for col in df.columns[:10]])
            raise ValueError(
                f"필수 컬럼 '{req_col}'을(를) 찾을 수 없습니다.\n"
                f"엑셀 파일의 컬럼명을 확인하세요.\n\n"
                f"현재 엑셀 파일의 컬럼: {available_cols}...\n\n"
                f"지원하는 컬럼명 예시:\n"
                f"  - 이름: {', '.join(column_patterns['구매자명'][:5])}, ...\n"
                f"  - 전화: {', '.join(column_patterns['전화번호'][:5])}, ...\n"
                f"  - 주소: {', '.join(column_patterns['주소'][:5])}, ..."
            )
    
    # 컬럼 이름 통일
    df = df.rename(columns=col_mapping)
    
    # 필수 컬럼만 선택
    df = df[required_cols].copy()
    
    # 빈 행 제거
    df = df.dropna(how='all')
    
    return df


def get_output_filenames(base_dir, base_name='ordered'):
    """
    출력 파일명 생성 (버전 관리)
    
    Args:
        base_dir: 저장 디렉토리
        base_name: 기본 파일명 (기본값: 'ordered')
        
    Returns:
        tuple: (pdf_path, csv_path)
        
    예시:
        ordered_20250109.pdf, ordered_20250109_match_report.csv
        파일이 이미 존재하면:
        ordered_20250109_v2.pdf, ordered_20250109_v2_match_report.csv
    """
    # 오늘 날짜
    today = datetime.now().strftime('%Y%m%d')
    
    # 기본 파일명
    base = f"{base_name}_{today}"
    
    # 기존 파일 확인 및 버전 찾기
    existing_files = list(Path(base_dir).glob(f"{base}*.pdf"))
    
    if not existing_files:
        # 첫 번째 파일
        pdf_name = f"{base}.pdf"
        csv_name = f"{base}_match_report.csv"
    else:
        # 버전 번호 추출
        version_pattern = re.compile(rf"{re.escape(base)}(?:_v(\d+))?\.pdf")
        max_version = 0
        
        for file_path in existing_files:
            match = version_pattern.match(file_path.name)
            if match:
                version = match.group(1)
                if version is None:
                    max_version = max(max_version, 1)
                else:
                    max_version = max(max_version, int(version))
        
        # 다음 버전
        next_version = max_version + 1
        pdf_name = f"{base}_v{next_version}.pdf"
        csv_name = f"{base}_v{next_version}_match_report.csv"
    
    pdf_path = os.path.join(base_dir, pdf_name)
    csv_path = os.path.join(base_dir, csv_name)
    
    return pdf_path, csv_path


def save_pdf(pdf_writer, out_pdf_path):
    """
    PDF 파일 저장
    
    Args:
        pdf_writer: PdfWriter 객체
        out_pdf_path: 출력 PDF 경로
    """
    # 디렉토리 생성
    os.makedirs(os.path.dirname(out_pdf_path), exist_ok=True)
    
    # PDF 저장
    with open(out_pdf_path, 'wb') as f:
        pdf_writer.write(f)


def save_report(rows, out_csv_path):
    """
    매칭 리포트 CSV 저장
    
    Args:
        rows: 리포트 행 데이터 리스트
              각 행은 dict: {
                  '엑셀행번호': int,
                  '매칭페이지': int or 'UNMATCHED',
                  '점수': float,
                  '매칭키': str (예: 'name+phone+addr+order'),
                  '구매자명': str,
                  '전화번호': str,
                  '주소': str,
                  '주문번호': str
              }
        out_csv_path: 출력 CSV 경로
    """
    # 디렉토리 생성
    os.makedirs(os.path.dirname(out_csv_path), exist_ok=True)
    
    # DataFrame 생성
    df = pd.DataFrame(rows)
    
    # 컬럼 순서 정렬
    columns = ['엑셀행번호', '매칭페이지', '점수', '매칭키', '구매자명', '전화번호', '주소', '주문번호']
    df = df[columns]
    
    # CSV 저장 (UTF-8 with BOM for Excel compatibility)
    df.to_csv(out_csv_path, index=False, encoding='utf-8-sig')


def is_text_based_pdf(pdf_path, sample_pages=5):
    """
    PDF가 텍스트 기반인지 확인 (이미지 스캔 여부 체크)
    
    Args:
        pdf_path: PDF 파일 경로
        sample_pages: 확인할 샘플 페이지 수
        
    Returns:
        tuple: (is_text_based: bool, message: str)
    """
    import pdfplumber
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            check_pages = min(sample_pages, total_pages)
            
            text_found = 0
            for i in range(check_pages):
                text = pdf.pages[i].extract_text() or ""
                # 의미있는 텍스트가 있는지 확인 (최소 50자)
                if len(text.strip()) > 50:
                    text_found += 1
            
            # 샘플 페이지의 80% 이상에서 텍스트를 찾아야 함
            threshold = check_pages * 0.8
            
            if text_found >= threshold:
                return True, "텍스트 기반 PDF입니다."
            else:
                return False, (
                    f"이미지 스캔 PDF로 판단됩니다. "
                    f"({check_pages}페이지 중 {text_found}페이지에서만 텍스트 발견)\n"
                    f"텍스트 추출이 가능한 PDF를 사용하세요."
                )
    except Exception as e:
        return False, f"PDF 확인 중 오류: {str(e)}"

