"""
PDF 텍스트 추출 및 매칭 로직
- 텍스트 추출 및 정규화
- 이름 + 전화번호 + 주소 3요소 매칭
- PDF 페이지 재정렬
"""

import re
from dataclasses import dataclass
from typing import List, Dict, Tuple, Set
import pdfplumber
from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz


@dataclass
class PageInfo:
    """PDF 페이지 정보"""
    index: int  # 페이지 번호 (0-based)
    raw_text: str  # 원본 텍스트
    norm_name_candidates: List[str]  # 정규화된 이름 후보들
    norm_phone_list: List[str]  # 정규화된 전화번호 리스트
    norm_addr_candidates: List[str]  # 정규화된 주소 후보들
    norm_order_candidates: List[str]  # 정규화된 주문번호 후보들


def remove_special_chars(text):
    """제로폭 문자 및 비정상 공백 제거"""
    if not text:
        return ""
    
    # 제로폭 문자 제거
    text = re.sub(r'[\u200b-\u200f\ufeff]', '', text)
    
    # 비정상 공백 제거
    text = re.sub(r'[\u3000\xa0]', ' ', text)
    
    return text


def normalize_name(text):
    """
    이름 정규화
    - 공백, 특수문자 제거
    - 영문은 대문자화
    - 한글은 그대로
    """
    if not text or pd.isna(text):
        return ""
    
    text = str(text).strip()
    text = remove_special_chars(text)
    
    # 공백 및 특수문자 제거 (한글, 영문, 숫자만 유지)
    text = re.sub(r'[^가-힣A-Za-z0-9]', '', text)
    
    # 영문은 대문자화
    text = text.upper()
    
    return text


def extract_name_candidates(text):
    """
    이름에서 여러 후보 추출
    - 소괄호 안 이름: 임재숙(양성희) -> [임재숙, 양성희]
    - 대괄호 안 이름: 이순자[이화순] -> [이순자, 이화순]
    - 공백으로 구분된 이름: 윤익주 김민정 -> [윤익주, 김민정]
    - 숫자 제거 버전: 본순박0 -> [본순박0, 본순박]
    - 전체 정규화된 이름도 포함
    """
    if not text or pd.isna(text):
        return []
    
    text = str(text).strip()
    candidates = []
    
    # 1. 전체 정규화 (기본)
    normalized_full = normalize_name(text)
    if normalized_full:
        candidates.append(normalized_full)
        # 숫자 제거 버전도 추가 (예: 본순박0 -> 본순박)
        no_digit = re.sub(r'[0-9]', '', normalized_full)
        if no_digit and no_digit != normalized_full and no_digit not in candidates:
            candidates.append(no_digit)
    
    # 2. 소괄호 () 안의 이름 추출
    # 예: 임재숙(양성희) -> 양성희
    paren_matches = re.findall(r'\(([^)]+)\)', text)
    for match in paren_matches:
        normalized = normalize_name(match)
        if normalized and normalized not in candidates:
            candidates.append(normalized)
            # 숫자 제거 버전
            no_digit = re.sub(r'[0-9]', '', normalized)
            if no_digit and no_digit != normalized and no_digit not in candidates:
                candidates.append(no_digit)
    
    # 3. 대괄호 [] 안의 이름 추출
    # 예: 이순자[이화순] -> 이화순
    bracket_matches = re.findall(r'\[([^\]]+)\]', text)
    for match in bracket_matches:
        normalized = normalize_name(match)
        if normalized and normalized not in candidates:
            candidates.append(normalized)
            # 숫자 제거 버전
            no_digit = re.sub(r'[0-9]', '', normalized)
            if no_digit and no_digit != normalized and no_digit not in candidates:
                candidates.append(no_digit)
    
    # 4. 괄호 밖의 이름 추출
    # 예: 임재숙(양성희) -> 임재숙, 이순자[이화순] -> 이순자
    text_without_brackets = re.sub(r'\([^)]+\)', '', text)
    text_without_brackets = re.sub(r'\[[^\]]+\]', '', text_without_brackets)
    normalized = normalize_name(text_without_brackets)
    if normalized and normalized not in candidates:
        candidates.append(normalized)
        # 숫자 제거 버전
        no_digit = re.sub(r'[0-9]', '', normalized)
        if no_digit and no_digit != normalized and no_digit not in candidates:
            candidates.append(no_digit)
    
    # 5. 공백으로 구분된 이름들
    # 예: 윤익주 김민정 -> 윤익주, 김민정
    parts = text.split()
    if len(parts) > 1:
        for part in parts:
            # 모든 괄호 제거
            part_clean = re.sub(r'\([^)]+\)', '', part)
            part_clean = re.sub(r'\[[^\]]+\]', '', part_clean)
            normalized = normalize_name(part_clean)
            if normalized and normalized not in candidates:
                candidates.append(normalized)
                # 숫자 제거 버전
                no_digit = re.sub(r'[0-9]', '', normalized)
                if no_digit and no_digit != normalized and no_digit not in candidates:
                    candidates.append(no_digit)
    
    return candidates


def normalize_phone(text):
    """
    전화번호 정규화
    - 숫자만 추출
    - 010으로 시작하는 번호만 인정
    - 10으로 시작하는 10자리는 앞에 0 추가 (엑셀에서 0이 사라진 경우 대응)
    - 8자리 숫자는 앞에 010 추가 시도
    """
    if not text or pd.isna(text):
        return ""
    
    text = str(text).strip()
    text = remove_special_chars(text)
    
    # 숫자만 추출
    numbers = re.sub(r'[^0-9]', '', text)
    
    # 010으로 시작하고 10-11자리인 번호
    if numbers.startswith('010') and len(numbers) in [10, 11]:
        return numbers
    
    # 10으로 시작하고 10자리인 번호 (엑셀에서 앞의 0이 사라진 경우)
    # 예: 1026417075 -> 01026417075
    if numbers.startswith('10') and len(numbers) == 10:
        return '0' + numbers
    
    # 9자리 숫자이고 10으로 시작하면 앞에 0 추가 (0이 하나 빠진 경우)
    # 예: 108302565 -> 0108302565
    if numbers.startswith('10') and len(numbers) == 9:
        return '0' + numbers
    
    # 8자리 숫자인 경우 앞에 010 추가 시도 (010에서 01이 빠진 경우)
    # 예: 27357395 -> 01027357395
    if len(numbers) == 8:
        return '010' + numbers
    
    # 7자리 숫자인 경우 앞에 010 추가 시도 (010에서 010이 빠진 경우)
    # 예: 2584757 -> 01012584757 (추가 확인 필요)
    if len(numbers) == 7:
        return '010' + numbers
    
    return ""


def normalize_addr(text):
    """
    주소 정규화
    - 괄호, 쉼표, 하이픈, 공백 제거
    - 숫자, 한글, 영문만 유지
    """
    if not text or pd.isna(text):
        return ""
    
    text = str(text).strip()
    text = remove_special_chars(text)
    
    # 괄호 내용 추출 우선
    bracket_match = re.search(r'\(([^)]{5,})\)', text)
    if bracket_match:
        text = bracket_match.group(1)
    
    # 특수문자, 공백 제거 (한글, 영문, 숫자만 유지)
    text = re.sub(r'[^가-힣A-Za-z0-9]', '', text)
    
    # 영문은 대문자화
    text = text.upper()
    
    return text


def normalize_order_number(text):
    """
    주문번호 정규화
    - 영문자와 숫자만 추출
    - 대문자 변환
    - 하이픈, 공백 등 특수문자 제거
    """
    if not text or pd.isna(text):
        return ""
    
    text = str(text).strip()
    text = remove_special_chars(text)
    
    # 영문자와 숫자만 추출
    result = re.sub(r'[^a-zA-Z0-9]', '', text)
    
    # 대문자 변환
    result = result.upper()
    
    return result


def extract_names_from_text(text):
    """텍스트에서 이름 후보 추출"""
    candidates = []
    
    # 한글 이름 패턴 (2-10자로 확장 - 긴 이름도 인식)
    # 예: 이경애본순박영아 (8자)
    korean_names = re.findall(r'[가-힣]{2,10}', text)
    candidates.extend(korean_names)
    
    # 한글 외자 이름 패턴 (독립적으로 나타나는 1자)
    # 공백, 줄바꿈, 쉼표 등으로 구분된 1자만 추출
    # 예: "셈" (독립적), "서울"의 "서"는 제외
    single_char_names = re.findall(r'(?:^|\s|,|\.|\(|\)|:)([가-힣])(?:\s|,|\.|\(|\)|:|$)', text)
    candidates.extend(single_char_names)
    
    # 영문 이름 패턴 1: 일반적인 형식 (John Smith)
    english_names_1 = re.findall(r'[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*', text)
    candidates.extend(english_names_1)
    
    # 영문 이름 패턴 2: 전체 대문자 형식 (LI JINGSHI, BAI FENGJIU 등)
    # 2개 또는 3개의 대문자 단어로 이루어진 패턴 (각각 독립적으로)
    english_names_2_word = re.findall(r'\b[A-Z]{2,}\s+[A-Z]{2,}\b', text)
    candidates.extend(english_names_2_word)
    english_names_3_word = re.findall(r'\b[A-Z]{2,}\s+[A-Z]{2,}\s+[A-Z]{2,}\b', text)
    candidates.extend(english_names_3_word)
    
    return candidates


def extract_phones_from_text(text):
    """텍스트에서 전화번호 추출"""
    phones = []
    
    # 1. 010으로 시작하는 전화번호 패턴 (기본)
    patterns = [
        r'010[-\s]?\d{3,4}[-\s]?\d{4}',  # 010-1234-5678 또는 010 1234 5678
        r'010\d{7,8}',  # 01012345678
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text)
        phones.extend(matches)
    
    # 2. 10으로 시작하는 9~10자리 (앞의 0이 빠진 경우)
    # 예: 1026417075, 108302565
    pattern_10 = r'(?:^|[^\d])(10\d{7,8})(?:[^\d]|$)'
    matches_10 = re.findall(pattern_10, text)
    phones.extend(matches_10)
    
    # 3. "전화", "연락처", "TEL", "PHONE" 키워드 근처의 7~10자리 숫자
    # 예: 전화번호: 27357395
    phone_keywords = ['전화', '연락', 'TEL', 'PHONE', 'HP', 'MOBILE', '핸드폰', '휴대폰', 'CALL']
    for keyword in phone_keywords:
        # 키워드 뒤 50자 이내의 7~10자리 숫자
        pattern = rf'{keyword}[^\d]{{0,50}}?(\d{{7,10}})'
        matches = re.findall(pattern, text, re.IGNORECASE)
        phones.extend(matches)
    
    # 4. 한 줄에 숫자만 있는 8~10자리 (전화번호 가능성이 높음)
    # 예: 줄바꿈 후 "27357395" 줄바꿈
    lines = text.split('\n')
    for line in lines:
        line_stripped = line.strip()
        # 숫자만 있고 8~10자리인 경우
        if line_stripped.isdigit() and len(line_stripped) in [8, 9, 10]:
            phones.append(line_stripped)
    
    # 중복 제거
    phones = list(dict.fromkeys(phones))
    
    return phones


def extract_addresses_from_text(text):
    """텍스트에서 주소 후보 추출"""
    candidates = []
    
    # 괄호 안의 주소
    bracket_addrs = re.findall(r'\(([^)]{10,})\)', text)
    candidates.extend(bracket_addrs)
    
    # 한국 주소 키워드 포함 문장
    addr_keywords = ['시', '구', '동', '로', '길', '번지', '호', '아파트', '빌딩']
    lines = text.split('\n')
    
    for line in lines:
        if any(keyword in line for keyword in addr_keywords):
            if len(line.strip()) >= 10:  # 너무 짧은 주소는 제외
                candidates.append(line.strip())
    
    return candidates


def extract_order_numbers_from_text(text):
    """텍스트에서 주문번호 후보 추출"""
    candidates = []
    
    # 주문번호 패턴: A-1234567 형식
    pattern1 = re.findall(r'[A-Z]-\d{6,}', text)
    candidates.extend(pattern1)
    
    # 주문번호 패턴: ORD-2024-001 형식
    pattern2 = re.findall(r'[A-Z]{2,}-\d{4,}-\d{3,}', text)
    candidates.extend(pattern2)
    
    # 주문번호 패턴: 20241009001 형식 (날짜+번호)
    pattern3 = re.findall(r'20\d{6,}', text)
    candidates.extend(pattern3)
    
    return candidates


def extract_pages(pdf_path):
    """
    PDF에서 모든 페이지의 텍스트 추출 및 정규화
    
    Args:
        pdf_path: PDF 파일 경로
        
    Returns:
        List[PageInfo]: 페이지 정보 리스트
    """
    pages = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            # 텍스트 추출
            raw_text = page.extract_text() or ""
            raw_text = remove_special_chars(raw_text)
            
            # 이름 후보 추출 및 정규화
            name_candidates = extract_names_from_text(raw_text)
            norm_names = [normalize_name(name) for name in name_candidates]
            norm_names = [n for n in norm_names if n]  # 빈 문자열 제거
            
            # 전화번호 추출 및 정규화
            phone_candidates = extract_phones_from_text(raw_text)
            norm_phones = [normalize_phone(phone) for phone in phone_candidates]
            norm_phones = [p for p in norm_phones if p]  # 빈 문자열 제거
            
            # 주소 후보 추출 및 정규화
            addr_candidates = extract_addresses_from_text(raw_text)
            norm_addrs = [normalize_addr(addr) for addr in addr_candidates]
            norm_addrs = [a for a in norm_addrs if a]  # 빈 문자열 제거
            
            # 주문번호 후보 추출 및 정규화
            order_candidates = extract_order_numbers_from_text(raw_text)
            norm_orders = [normalize_order_number(order) for order in order_candidates]
            norm_orders = [o for o in norm_orders if o]  # 빈 문자열 제거
            
            page_info = PageInfo(
                index=i,
                raw_text=raw_text,
                norm_name_candidates=norm_names,
                norm_phone_list=norm_phones,
                norm_addr_candidates=norm_addrs,
                norm_order_candidates=norm_orders
            )
            
            pages.append(page_info)
    
    return pages


def calc_match_score(excel_name, excel_phone, excel_addr, excel_order, page_info, use_fuzzy=False, threshold=90):
    """
    엑셀 행과 PDF 페이지의 매칭 점수 계산 - 주문번호 기준 매칭
    
    Args:
        excel_name: 정규화된 엑셀 이름 (사용 안함)
        excel_phone: 정규화된 엑셀 전화번호 (사용 안함)
        excel_addr: 정규화된 엑셀 주소 (사용 안함)
        excel_order: 정규화된 엑셀 주문번호
        page_info: PageInfo 객체
        use_fuzzy: 유사도 매칭 사용 여부
        threshold: 유사도 임계값
        
    Returns:
        tuple: (score, reason)
            - score: 매칭 점수 (0-100)
            - reason: 매칭 근거 ('order_exact', 'order_fuzzy', 등)
    """
    # 주문번호가 비어있으면 매칭 불가
    if not excel_order or not page_info.norm_order_candidates:
        return 0, 'no_order_data'
    
    # 정확 일치 확인 (주문번호만)
    order_match = excel_order in page_info.norm_order_candidates
    
    # 주문번호가 정확히 일치하면 100점 (조기 종료)
    if order_match:
        return 100, 'order_exact'
    
    # 유사도 매칭 사용하지 않으면 0점
    if not use_fuzzy:
        return 0, 'no_match'
    
    # 주문번호 유사도 매칭
    best_order_sim = max((fuzz.ratio(excel_order, o) for o in page_info.norm_order_candidates), default=0)
    
    # 임계값 이상이면 점수 반환
    if best_order_sim >= threshold:
        return best_order_sim, f'order_fuzzy({best_order_sim:.0f}%)'
    
    return 0, 'no_match'


def match_rows_to_pages(df, pages, use_fuzzy=False, threshold=90):
    """
    엑셀 행과 PDF 페이지를 매칭 - 주문번호 기준
    
    Args:
        df: 엑셀 DataFrame (주문번호 컬럼 포함)
        pages: List[PageInfo]
        use_fuzzy: 유사도 매칭 사용 여부
        threshold: 유사도 임계값
        
    Returns:
        tuple: (assignments, leftover_pages, match_details)
            - assignments: {excel_row_idx: page_idx, ...}
            - leftover_pages: [page_idx, ...] (매칭되지 않은 페이지)
            - match_details: {excel_row_idx: {'page_idx': int, 'score': float, 'reason': str}, ...}
    """
    assignments = {}
    match_details = {}
    used_pages = set()
    
    # 각 엑셀 행에 대해 매칭
    for row_idx, row in df.iterrows():
        # 주문번호만 정규화
        excel_order = normalize_order_number(row['주문번호'])
        
        # 주문번호가 비어있으면 매칭 불가
        if not excel_order:
            match_details[row_idx] = {
                'page_idx': -1,
                'score': 0,
                'reason': 'empty_order_number'
            }
            continue
        
        # 모든 페이지와 비교
        best_page = -1
        best_score = 0
        best_reason = 'no_match'
        
        for page_info in pages:
            # 이미 사용된 페이지는 건너뛰기
            if page_info.index in used_pages:
                continue
            
            # 주문번호 기준으로 매칭 점수 계산
            score, reason = calc_match_score(
                "", "", "", excel_order,  # 이름, 전화번호, 주소는 빈 값으로 전달
                page_info, use_fuzzy, threshold
            )
            
            if score > best_score:
                best_score = score
                best_page = page_info.index
                best_reason = reason
                
                # 완벽한 매칭을 찾으면 즉시 중단 (조기 종료)
                if score == 100:
                    break
        
        # 매칭 결과 저장
        if best_score > 0:
            assignments[row_idx] = best_page
            used_pages.add(best_page)
            match_details[row_idx] = {
                'page_idx': best_page,
                'score': best_score,
                'reason': best_reason
            }
        else:
            match_details[row_idx] = {
                'page_idx': -1,
                'score': 0,
                'reason': 'no_match'
            }
    
    # 남은 페이지
    all_pages = set(p.index for p in pages)
    leftover_pages = sorted(all_pages - used_pages)
    
    return assignments, leftover_pages, match_details


def reorder_pdf(pdf_path, ordered_indices, out_pdf_path):
    """
    PDF 페이지 재정렬
    
    Args:
        pdf_path: 원본 PDF 경로
        ordered_indices: 정렬된 페이지 인덱스 리스트
        out_pdf_path: 출력 PDF 경로
    """
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    
    # 지정된 순서대로 페이지 추가
    for idx in ordered_indices:
        writer.add_page(reader.pages[idx])
    
    # 저장
    with open(out_pdf_path, 'wb') as f:
        writer.write(f)


# pandas import 추가
import pandas as pd

