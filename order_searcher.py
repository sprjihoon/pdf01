"""
주문번호 검색 및 파일 관리
- PDF 파일에서 주문번호 검색
- 중복 주문번호 중 최신 파일 선택
- 날짜 기준 우선순위: 문서날짜 > 파일명날짜 > 수정시각
"""

import os
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Set
from dataclasses import dataclass
import pdfplumber
from config_manager import config


@dataclass
class OrderMatch:
    """주문번호 매칭 결과"""
    order_number: str  # 주문번호
    file_path: str  # 파일 경로
    page_numbers: List[int]  # 주문번호가 발견된 페이지 번호 리스트
    doc_date: Optional[datetime]  # 문서 내 날짜
    filename_date: Optional[datetime]  # 파일명에서 추출한 날짜
    modified_time: datetime  # 파일 수정 시간
    priority_score: int  # 우선순위 점수 (높을수록 최신)


@dataclass
class SearchResult:
    """검색 결과"""
    order_number: str
    best_match: OrderMatch  # 최신 파일
    all_matches: List[OrderMatch]  # 모든 매칭 파일들
    decided_by: str  # 최신 파일 결정 기준


class OrderSearcher:
    """주문번호 검색 및 관리 클래스"""
    
    def __init__(self):
        self.order_pattern = re.compile(config.get_order_pattern())
        
        # 날짜 추출 패턴들
        self.date_patterns = [
            # 문서 내 날짜 패턴
            (r'날짜[:\s]*(\d{4})[.-](\d{1,2})[.-](\d{1,2})', 'doc_date'),
            (r'작성일[:\s]*(\d{4})[.-](\d{1,2})[.-](\d{1,2})', 'doc_date'),
            (r'발행일[:\s]*(\d{4})[.-](\d{1,2})[.-](\d{1,2})', 'doc_date'),
            (r'(\d{4})[.-](\d{1,2})[.-](\d{1,2})', 'doc_date'),
            
            # 파일명 날짜 패턴
            (r'(\d{4})[_-](\d{1,2})[_-](\d{1,2})', 'filename_date'),
            (r'(\d{8})', 'filename_date_compact'),  # 20241209 형식
        ]
    
    def search_order_in_folder(self, folder_path: str, order_number: str) -> Optional[SearchResult]:
        """
        폴더에서 특정 주문번호 검색
        
        Args:
            folder_path: 검색할 폴더 경로
            order_number: 찾을 주문번호
            
        Returns:
            SearchResult 또는 None (미발견시)
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")
        
        pdf_files = self._find_pdf_files(folder_path)
        matches = []
        
        for pdf_file in pdf_files:
            try:
                match = self._search_order_in_file(pdf_file, order_number)
                if match:
                    matches.append(match)
            except Exception as e:
                print(f"파일 검색 중 오류 ({pdf_file}): {e}")
                continue
        
        if not matches:
            return None
        
        # 최신 파일 선택
        best_match, decided_by = self._select_latest_file(matches)
        
        return SearchResult(
            order_number=order_number,
            best_match=best_match,
            all_matches=matches,
            decided_by=decided_by
        )
    
    def search_all_orders_in_folder(self, folder_path: str) -> Dict[str, SearchResult]:
        """
        폴더에서 모든 주문번호 검색
        
        Args:
            folder_path: 검색할 폴더 경로
            
        Returns:
            {주문번호: SearchResult} 딕셔너리
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")
        
        pdf_files = self._find_pdf_files(folder_path)
        all_matches = {}  # {order_number: [OrderMatch, ...]}
        
        for pdf_file in pdf_files:
            try:
                file_matches = self._find_all_orders_in_file(pdf_file)
                for match in file_matches:
                    order_num = match.order_number
                    if order_num not in all_matches:
                        all_matches[order_num] = []
                    all_matches[order_num].append(match)
                    
            except Exception as e:
                print(f"파일 검색 중 오류 ({pdf_file}): {e}")
                continue
        
        # 각 주문번호별로 최신 파일 선택
        results = {}
        for order_number, matches in all_matches.items():
            best_match, decided_by = self._select_latest_file(matches)
            results[order_number] = SearchResult(
                order_number=order_number,
                best_match=best_match,
                all_matches=matches,
                decided_by=decided_by
            )
        
        return results
    
    def _find_pdf_files(self, folder_path: str) -> List[str]:
        """폴더에서 PDF 파일 목록 가져오기"""
        pdf_files = []
        folder = Path(folder_path)
        
        if config.get("search_settings.recursive_search", True):
            # 재귀 검색
            pdf_files = list(folder.rglob("*.pdf"))
        else:
            # 현재 폴더만
            pdf_files = list(folder.glob("*.pdf"))
        
        return [str(f) for f in pdf_files]
    
    def _search_order_in_file(self, file_path: str, order_number: str) -> Optional[OrderMatch]:
        """특정 파일에서 주문번호 검색"""
        try:
            with pdfplumber.open(file_path) as pdf:
                matching_pages = []
                doc_date = None
                
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text() or ""
                    
                    # 주문번호 검색
                    if order_number in text:
                        matching_pages.append(page_num)
                    
                    # 문서 날짜 추출 (첫 번째 페이지에서만)
                    if page_num == 1 and not doc_date:
                        doc_date = self._extract_doc_date(text)
                
                if not matching_pages:
                    return None
                
                # 파일명에서 날짜 추출
                filename_date = self._extract_filename_date(file_path)
                
                # 파일 수정 시간
                modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                
                # 우선순위 점수 계산
                priority_score = self._calculate_priority_score(doc_date, filename_date, modified_time)
                
                return OrderMatch(
                    order_number=order_number,
                    file_path=file_path,
                    page_numbers=matching_pages,
                    doc_date=doc_date,
                    filename_date=filename_date,
                    modified_time=modified_time,
                    priority_score=priority_score
                )
                
        except Exception as e:
            print(f"PDF 파일 읽기 오류 ({file_path}): {e}")
            return None
    
    def _find_all_orders_in_file(self, file_path: str) -> List[OrderMatch]:
        """파일에서 모든 주문번호 찾기"""
        try:
            with pdfplumber.open(file_path) as pdf:
                all_orders = set()  # 중복 제거
                page_mapping = {}  # {order_number: [page_numbers]}
                doc_date = None
                
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text() or ""
                    
                    # 주문번호 패턴 검색
                    orders_in_page = self.order_pattern.findall(text)
                    
                    for order in orders_in_page:
                        all_orders.add(order)
                        if order not in page_mapping:
                            page_mapping[order] = []
                        if page_num not in page_mapping[order]:
                            page_mapping[order].append(page_num)
                    
                    # 문서 날짜 추출 (첫 번째 페이지에서만)
                    if page_num == 1 and not doc_date:
                        doc_date = self._extract_doc_date(text)
                
                if not all_orders:
                    return []
                
                # 파일명에서 날짜 추출
                filename_date = self._extract_filename_date(file_path)
                
                # 파일 수정 시간
                modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                
                # 우선순위 점수 계산
                priority_score = self._calculate_priority_score(doc_date, filename_date, modified_time)
                
                # 각 주문번호별로 OrderMatch 생성
                matches = []
                for order_num in all_orders:
                    matches.append(OrderMatch(
                        order_number=order_num,
                        file_path=file_path,
                        page_numbers=page_mapping[order_num],
                        doc_date=doc_date,
                        filename_date=filename_date,
                        modified_time=modified_time,
                        priority_score=priority_score
                    ))
                
                return matches
                
        except Exception as e:
            print(f"PDF 파일 읽기 오류 ({file_path}): {e}")
            return []
    
    def _extract_doc_date(self, text: str) -> Optional[datetime]:
        """문서 텍스트에서 날짜 추출"""
        for pattern, date_type in self.date_patterns:
            if date_type == 'doc_date':
                matches = re.findall(pattern, text)
                if matches:
                    try:
                        if len(matches[0]) == 3:
                            year, month, day = matches[0]
                            return datetime(int(year), int(month), int(day))
                    except (ValueError, IndexError):
                        continue
        return None
    
    def _extract_filename_date(self, file_path: str) -> Optional[datetime]:
        """파일명에서 날짜 추출"""
        filename = os.path.basename(file_path)
        
        for pattern, date_type in self.date_patterns:
            if date_type in ['filename_date', 'filename_date_compact']:
                matches = re.findall(pattern, filename)
                if matches:
                    try:
                        if date_type == 'filename_date_compact':
                            # 20241209 형식
                            date_str = matches[0]
                            if len(date_str) == 8:
                                year = int(date_str[:4])
                                month = int(date_str[4:6])
                                day = int(date_str[6:8])
                                return datetime(year, month, day)
                        else:
                            # YYYY-MM-DD 형식
                            if len(matches[0]) == 3:
                                year, month, day = matches[0]
                                return datetime(int(year), int(month), int(day))
                    except (ValueError, IndexError):
                        continue
        return None
    
    def _calculate_priority_score(self, doc_date: Optional[datetime], 
                                filename_date: Optional[datetime],
                                modified_time: datetime) -> int:
        """우선순위 점수 계산 (높을수록 최신)"""
        score = 0
        
        # 문서 날짜 (가장 높은 우선순위)
        if doc_date:
            score += int(doc_date.timestamp()) * 1000000
        
        # 파일명 날짜 (중간 우선순위)
        elif filename_date:
            score += int(filename_date.timestamp()) * 1000
        
        # 파일 수정 시간 (가장 낮은 우선순위)
        score += int(modified_time.timestamp())
        
        return score
    
    def _select_latest_file(self, matches: List[OrderMatch]) -> Tuple[OrderMatch, str]:
        """
        최신 파일 선택
        
        Returns:
            (최신 OrderMatch, 결정 기준)
        """
        if not matches:
            raise ValueError("매칭 결과가 없습니다")
        
        if len(matches) == 1:
            match = matches[0]
            if match.doc_date:
                return match, "doc_date"
            elif match.filename_date:
                return match, "filename_date"
            else:
                return match, "modified_time"
        
        # 우선순위 점수로 정렬
        sorted_matches = sorted(matches, key=lambda x: x.priority_score, reverse=True)
        best_match = sorted_matches[0]
        
        # 결정 기준 판단
        if best_match.doc_date:
            return best_match, "doc_date"
        elif best_match.filename_date:
            return best_match, "filename_date"
        else:
            return best_match, "modified_time"
    
    def get_page_ranges_str(self, page_numbers: List[int]) -> str:
        """페이지 번호 리스트를 범위 문자열로 변환"""
        if not page_numbers:
            return ""
        
        page_numbers = sorted(set(page_numbers))  # 중복 제거 및 정렬
        
        if len(page_numbers) == 1:
            return str(page_numbers[0])
        
        # 연속된 페이지는 범위로 표시
        ranges = []
        start = page_numbers[0]
        end = start
        
        for i in range(1, len(page_numbers)):
            if page_numbers[i] == end + 1:
                end = page_numbers[i]
            else:
                if start == end:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{end}")
                start = end = page_numbers[i]
        
        # 마지막 범위 추가
        if start == end:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{end}")
        
        return ",".join(ranges)
