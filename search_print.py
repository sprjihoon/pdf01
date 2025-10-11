"""
주문번호 검색 및 인쇄 기능
- PDF 파일/폴더에서 주문번호 검색
- 해당 페이지만 추출하여 별도 PDF로 저장
- 기본 프로그램으로 열어서 인쇄
- 멀티프로세싱으로 빠른 검색
"""

import os
import re
from pathlib import Path
from typing import List, Optional, Tuple
import pdfplumber
from pypdf import PdfReader, PdfWriter
from matcher import normalize_order_number, extract_order_numbers_from_text


def extract_text_pages_fast(pdf_path: str) -> List[str]:
    """
    빠른 텍스트 추출기: 우선 PyMuPDF(fitz) 사용, 실패 시 pdfplumber로 폴백
    Returns: 각 페이지의 텍스트 리스트
    """
    # 1) PyMuPDF(fitz) 시도
    try:
        import fitz  # PyMuPDF
        texts: List[str] = []
        with fitz.open(pdf_path) as doc:
            for page in doc:
                # 기본 텍스트 추출 (레이아웃 무시, 속도 우선)
                txt = page.get_text("text") or ""
                texts.append(txt)
        # PyMuPDF로 성공
        return texts
    except Exception:
        pass

    # 2) pdfplumber 폴백
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            texts.append(txt)
    return texts
from multiprocessing import cpu_count
import multiprocessing


def search_order_in_pdf(pdf_path: str, order_number: str) -> Optional[List[int]]:
    """
    PDF에서 주문번호를 찾아 페이지 번호 리스트 반환
    
    Args:
        pdf_path: PDF 파일 경로
        order_number: 검색할 주문번호
        
    Returns:
        페이지 번호 리스트 (1-based) 또는 None
    """
    normalized_order = normalize_order_number(order_number)
    if not normalized_order:
        return None
    
    found_pages = []
    
    try:
        # 빠른 텍스트 추출 (PyMuPDF → pdfplumber 폴백)
        page_texts = extract_text_pages_fast(pdf_path)

        for page_num, text in enumerate(page_texts, start=1):
            # 페이지에서 주문번호 후보들 추출
            order_candidates = extract_order_numbers_from_text(text)

            # 후보 정규화 집합으로 가속
            normalized_set = set()
            for cand in order_candidates:
                nc = normalize_order_number(cand)
                if nc:
                    normalized_set.add(nc)

            # 정확 일치
            if normalized_order in normalized_set:
                found_pages.append(page_num)
                continue

            # 부분 일치 (최소 6자리)
            if len(normalized_order) >= 6:
                match_partial = any(
                    (len(nc) >= 6 and (normalized_order in nc or nc in normalized_order))
                    for nc in normalized_set
                )
                if match_partial:
                    found_pages.append(page_num)

        return found_pages if found_pages else None

    except Exception as e:
        print(f"PDF 검색 중 오류 ({pdf_path}): {e}")
        return None


def _search_single_file(args):
    """
    단일 파일 검색 (멀티프로세싱용 헬퍼 함수)
    
    Args:
        args: (pdf_path, order_number) 튜플
        
    Returns:
        (pdf_path, pages, modified_time) 또는 None
    """
    pdf_path, order_number = args
    try:
        pages = search_order_in_pdf(pdf_path, order_number)
        if pages:
            modified_time = os.path.getmtime(pdf_path)
            return (pdf_path, pages, modified_time)
    except Exception as e:
        print(f"Error searching {pdf_path}: {e}")
    return None


def search_order_in_folder_multiprocess(folder_path: str, order_number: str,
                                        progress_callback=None, stop_flag=None) -> List[Tuple[str, List[int], float]]:
    """
    폴더 내 모든 PDF에서 주문번호 검색 (멀티프로세싱 버전)
    
    Args:
        folder_path: 검색할 폴더 경로
        order_number: 검색할 주문번호
        progress_callback: 진행 상황 콜백 함수
        stop_flag: 중지 플래그 (callable)
        
    Returns:
        [(pdf_path, [page_numbers], modified_time), ...] 리스트 (수정시간 최신순 정렬)
    """
    results = []
    
    # PDF 파일 찾기
    if progress_callback:
        progress_callback("📂 PDF 파일 목록 생성 중...")
    
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        if stop_flag and stop_flag():
            return results
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    total_files = len(pdf_files)
    if progress_callback:
        progress_callback(f"📋 총 {total_files}개 PDF 파일 발견")
    
    if total_files == 0:
        return results
    
    # 파일 크기 기준으로 정렬 (작은 파일 먼저)
    pdf_files_with_size = []
    for pdf_path in pdf_files:
        try:
            size = os.path.getsize(pdf_path)
            pdf_files_with_size.append((pdf_path, size))
        except:
            pdf_files_with_size.append((pdf_path, 0))
    
    pdf_files_with_size.sort(key=lambda x: x[1])
    pdf_files_sorted = [p[0] for p in pdf_files_with_size]
    
    # CPU 코어 수 결정 (최대 4개로 제한해 안정성 확보)
    num_processes = min(cpu_count(), 4, total_files)
    
    if progress_callback:
        progress_callback(f"⚡ {num_processes}개 프로세스로 병렬 검색 시작...")
    
    # 검색 인자 준비
    search_args = [(pdf_path, order_number) for pdf_path in pdf_files_sorted]
    
    # 멀티프로세싱으로 검색
    found_count = 0
    processed_count = 0
    
    try:
        # 청크 크기 설정 (진행 상황 업데이트를 위해)
        chunk_size = max(1, total_files // (num_processes * 10))

        # Windows 안전한 spawn 컨텍스트 사용 및 작업 수 제한으로 누수 방지
        ctx = multiprocessing.get_context("spawn")
        with ctx.Pool(processes=num_processes, maxtasksperchild=20) as pool:
            for result in pool.imap_unordered(_search_single_file, search_args, chunksize=chunk_size):
                # 중지 플래그 확인
                if stop_flag and stop_flag():
                    pool.terminate()
                    if progress_callback:
                        progress_callback(f"⏸️ 검색 중지됨 ({processed_count}/{total_files})")
                    break
                
                processed_count += 1
                
                if result:
                    pdf_path, pages, modified_time = result
                    results.append(result)
                    found_count += 1
                    
                    filename = os.path.basename(pdf_path)
                    if progress_callback:
                        progress_callback(f"✅ [{processed_count}/{total_files}] {filename} - 페이지: {', '.join(map(str, pages))}")
                else:
                    # 10개마다 진행 상황 업데이트
                    if processed_count % 10 == 0 and progress_callback:
                        progress_callback(f"🔍 검색 중... [{processed_count}/{total_files}] (발견: {found_count}개)")
    
    except Exception as e:
        if progress_callback:
            progress_callback(f"⚠️ 검색 중 오류: {e}")
    
    if progress_callback:
        progress_callback(f"🎯 검색 완료: {found_count}개 파일에서 발견 (총 {processed_count}개 검색)")
    
    # 수정 시간 기준 내림차순 정렬 (최신 파일이 먼저)
    results.sort(key=lambda x: x[2], reverse=True)
    
    return results


def search_order_in_folder(folder_path: str, order_number: str, 
                           progress_callback=None, stop_flag=None, use_multiprocess=True) -> List[Tuple[str, List[int], float]]:
    """
    폴더 내 모든 PDF에서 주문번호 검색
    
    Args:
        folder_path: 검색할 폴더 경로
        order_number: 검색할 주문번호
        progress_callback: 진행 상황 콜백 함수
        stop_flag: 중지 플래그 (callable)
        use_multiprocess: 멀티프로세싱 사용 여부 (기본 True)
        
    Returns:
        [(pdf_path, [page_numbers], modified_time), ...] 리스트 (수정시간 최신순 정렬)
    """
    # 멀티프로세싱 사용 여부 결정
    if use_multiprocess:
        # 파일 수가 적어도 병렬 이점이 있는 경우가 많음 → 2개 이상이면 사용
        try:
            pdf_count = sum(1 for root, dirs, files in os.walk(folder_path) 
                          for file in files if file.lower().endswith('.pdf'))
            if pdf_count >= 2:
                return search_order_in_folder_multiprocess(folder_path, order_number, progress_callback, stop_flag)
        except:
            pass
    
    # 단일 프로세스 버전 (원본 코드)
    results = []
    
    # PDF 파일 찾기
    if progress_callback:
        progress_callback("📂 PDF 파일 목록 생성 중...")
    
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        if stop_flag and stop_flag():
            return results
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    total_files = len(pdf_files)
    if progress_callback:
        progress_callback(f"📋 총 {total_files}개 PDF 파일 발견")
        progress_callback("🔧 단일 프로세스로 검색 중...")
    
    # 파일 크기 기준으로 정렬 (작은 파일 먼저 - 빠른 검색)
    pdf_files_with_size = []
    for pdf_path in pdf_files:
        try:
            size = os.path.getsize(pdf_path)
            pdf_files_with_size.append((pdf_path, size))
        except:
            pdf_files_with_size.append((pdf_path, 0))
    
    pdf_files_with_size.sort(key=lambda x: x[1])  # 작은 파일부터
    
    # 각 PDF에서 검색
    for idx, (pdf_path, size) in enumerate(pdf_files_with_size, 1):
        # 중지 플래그 확인
        if stop_flag and stop_flag():
            if progress_callback:
                progress_callback(f"⏸️ 검색 중지됨 ({idx}/{total_files})")
            break
        
        filename = os.path.basename(pdf_path)
        size_mb = size / (1024 * 1024)
        
        if progress_callback:
            progress_callback(f"🔍 [{idx}/{total_files}] {filename} ({size_mb:.1f}MB)")
        
        pages = search_order_in_pdf(pdf_path, order_number)
        if pages:
            # 파일 수정 시간 가져오기
            modified_time = os.path.getmtime(pdf_path)
            results.append((pdf_path, pages, modified_time))
            
            if progress_callback:
                progress_callback(f"✅ 발견! {filename} - 페이지: {', '.join(map(str, pages))}")
    
    if progress_callback:
        progress_callback(f"🎯 검색 완료: {len(results)}개 파일에서 발견")
    
    # 수정 시간 기준 내림차순 정렬 (최신 파일이 먼저)
    results.sort(key=lambda x: x[2], reverse=True)
    
    return results


def extract_pages_to_pdf(input_pdf: str, page_numbers: List[int], output_pdf: str):
    """
    특정 페이지들만 추출하여 새 PDF로 저장
    
    Args:
        input_pdf: 원본 PDF 경로
        page_numbers: 추출할 페이지 번호 리스트 (1-based)
        output_pdf: 출력 PDF 경로
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    # 페이지 추가 (1-based → 0-based)
    for page_num in page_numbers:
        if 1 <= page_num <= len(reader.pages):
            writer.add_page(reader.pages[page_num - 1])
    
    # 저장
    with open(output_pdf, 'wb') as f:
        writer.write(f)


def open_pdf_for_print(pdf_path: str):
    """
    PDF를 기본 프로그램으로 열기
    
    Args:
        pdf_path: PDF 파일 경로
    """
    if os.path.exists(pdf_path):
        os.startfile(pdf_path)

