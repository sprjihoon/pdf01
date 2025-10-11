"""
주문번호 검색 및 인쇄 로깅 시스템
- 검색/인쇄 작업 이력 기록
- CSV 형태로 저장하여 엑셀에서 분석 가능
"""

import os
import csv
from datetime import datetime
from typing import Optional, List, Dict, Any
from pathlib import Path
from dataclasses import dataclass, asdict
from order_searcher import SearchResult, OrderMatch


@dataclass
class SearchLogEntry:
    """검색 로그 엔트리"""
    timestamp: datetime
    order_no: str
    search_folder: str
    found: bool
    used_file: str = ""
    doc_date: Optional[datetime] = None
    filename_date: Optional[datetime] = None
    modified_time: Optional[datetime] = None
    page_ranges: str = ""
    decided_by: str = ""
    total_matches: int = 0
    search_duration_ms: int = 0


@dataclass
class PrintLogEntry:
    """인쇄 로그 엔트리"""
    timestamp: datetime
    order_no: str
    file_path: str
    page_ranges: str
    printer_name: str
    copies: int
    duplex: bool
    success: bool
    error_message: str = ""
    print_duration_ms: int = 0


class OrderLogger:
    """주문번호 검색/인쇄 로깅 클래스"""
    
    def __init__(self, log_dir: str = "logs"):
        """
        Args:
            log_dir: 로그 저장 디렉토리
        """
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # 로그 파일 경로들
        self.search_log_path = self.log_dir / "search_log.csv"
        self.print_log_path = self.log_dir / "print_log.csv"
        
        # CSV 헤더 초기화
        self._init_log_files()
    
    def _init_log_files(self):
        """로그 파일 초기화 (헤더 생성)"""
        # 검색 로그 초기화
        if not self.search_log_path.exists():
            self._write_search_log_header()
        
        # 인쇄 로그 초기화
        if not self.print_log_path.exists():
            self._write_print_log_header()
    
    def _write_search_log_header(self):
        """검색 로그 헤더 작성"""
        headers = [
            'timestamp', 'order_no', 'search_folder', 'found',
            'used_file', 'doc_date', 'filename_date', 'modified_time',
            'page_ranges', 'decided_by', 'total_matches', 'search_duration_ms'
        ]
        
        with open(self.search_log_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
    
    def _write_print_log_header(self):
        """인쇄 로그 헤더 작성"""
        headers = [
            'timestamp', 'order_no', 'file_path', 'page_ranges',
            'printer_name', 'copies', 'duplex', 'success',
            'error_message', 'print_duration_ms'
        ]
        
        with open(self.print_log_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
    
    def log_search(self, entry: SearchLogEntry):
        """검색 결과 로깅"""
        try:
            with open(self.search_log_path, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                row = [
                    entry.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                    entry.order_no,
                    entry.search_folder,
                    'Y' if entry.found else 'N',
                    entry.used_file,
                    entry.doc_date.strftime('%Y-%m-%d') if entry.doc_date else '',
                    entry.filename_date.strftime('%Y-%m-%d') if entry.filename_date else '',
                    entry.modified_time.strftime('%Y-%m-%d %H:%M:%S') if entry.modified_time else '',
                    entry.page_ranges,
                    entry.decided_by,
                    entry.total_matches,
                    entry.search_duration_ms
                ]
                writer.writerow(row)
        except Exception as e:
            print(f"검색 로그 저장 실패: {e}")
    
    def log_print(self, entry: PrintLogEntry):
        """인쇄 결과 로깅"""
        try:
            with open(self.print_log_path, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                row = [
                    entry.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                    entry.order_no,
                    entry.file_path,
                    entry.page_ranges,
                    entry.printer_name,
                    entry.copies,
                    'Y' if entry.duplex else 'N',
                    'Y' if entry.success else 'N',
                    entry.error_message,
                    entry.print_duration_ms
                ]
                writer.writerow(row)
        except Exception as e:
            print(f"인쇄 로그 저장 실패: {e}")
    
    def log_search_result(self, order_number: str, search_folder: str,
                         search_result: Optional[SearchResult],
                         search_duration_ms: int):
        """검색 결과를 로그에 기록"""
        if search_result:
            # 검색 성공
            best_match = search_result.best_match
            entry = SearchLogEntry(
                timestamp=datetime.now(),
                order_no=order_number,
                search_folder=search_folder,
                found=True,
                used_file=best_match.file_path,
                doc_date=best_match.doc_date,
                filename_date=best_match.filename_date,
                modified_time=best_match.modified_time,
                page_ranges=self._format_page_ranges(best_match.page_numbers),
                decided_by=search_result.decided_by,
                total_matches=len(search_result.all_matches),
                search_duration_ms=search_duration_ms
            )
        else:
            # 검색 실패
            entry = SearchLogEntry(
                timestamp=datetime.now(),
                order_no=order_number,
                search_folder=search_folder,
                found=False,
                search_duration_ms=search_duration_ms
            )
        
        self.log_search(entry)
    
    def log_print_result(self, order_number: str, file_path: str, 
                        page_ranges: str, printer_name: str,
                        copies: int, duplex: bool, success: bool,
                        error_message: str = "", print_duration_ms: int = 0):
        """인쇄 결과를 로그에 기록"""
        entry = PrintLogEntry(
            timestamp=datetime.now(),
            order_no=order_number,
            file_path=file_path,
            page_ranges=page_ranges,
            printer_name=printer_name,
            copies=copies,
            duplex=duplex,
            success=success,
            error_message=error_message,
            print_duration_ms=print_duration_ms
        )
        
        self.log_print(entry)
    
    def _format_page_ranges(self, page_numbers: List[int]) -> str:
        """페이지 번호 리스트를 범위 문자열로 변환"""
        if not page_numbers:
            return ""
        
        page_numbers = sorted(set(page_numbers))
        
        if len(page_numbers) == 1:
            return str(page_numbers[0])
        
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
    
    def get_search_statistics(self, days: int = 30) -> Dict[str, Any]:
        """검색 통계 정보 가져오기"""
        try:
            if not self.search_log_path.exists():
                return {}
            
            # 최근 N일 데이터 필터링
            cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
            
            total_searches = 0
            successful_searches = 0
            unique_orders = set()
            total_duration = 0
            
            with open(self.search_log_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        log_date = datetime.strptime(row['timestamp'], '%Y-%m-%d %H:%M:%S')
                        if log_date >= cutoff_date:
                            total_searches += 1
                            if row['found'] == 'Y':
                                successful_searches += 1
                            unique_orders.add(row['order_no'])
                            total_duration += int(row.get('search_duration_ms', 0))
                    except (ValueError, KeyError):
                        continue
            
            return {
                'period_days': days,
                'total_searches': total_searches,
                'successful_searches': successful_searches,
                'success_rate': (successful_searches / total_searches * 100) if total_searches > 0 else 0,
                'unique_orders': len(unique_orders),
                'avg_duration_ms': total_duration // total_searches if total_searches > 0 else 0
            }
            
        except Exception as e:
            print(f"검색 통계 생성 실패: {e}")
            return {}
    
    def get_print_statistics(self, days: int = 30) -> Dict[str, Any]:
        """인쇄 통계 정보 가져오기"""
        try:
            if not self.print_log_path.exists():
                return {}
            
            cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
            
            total_prints = 0
            successful_prints = 0
            total_pages = 0
            total_copies = 0
            
            with open(self.print_log_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        log_date = datetime.strptime(row['timestamp'], '%Y-%m-%d %H:%M:%S')
                        if log_date >= cutoff_date:
                            total_prints += 1
                            if row['success'] == 'Y':
                                successful_prints += 1
                            
                            # 페이지 수 계산 (간단한 추정)
                            page_ranges = row.get('page_ranges', '')
                            if page_ranges:
                                page_count = len(page_ranges.split(','))
                                total_pages += page_count
                            
                            copies = int(row.get('copies', 1))
                            total_copies += copies
                            
                    except (ValueError, KeyError):
                        continue
            
            return {
                'period_days': days,
                'total_prints': total_prints,
                'successful_prints': successful_prints,
                'success_rate': (successful_prints / total_prints * 100) if total_prints > 0 else 0,
                'total_pages': total_pages,
                'total_copies': total_copies
            }
            
        except Exception as e:
            print(f"인쇄 통계 생성 실패: {e}")
            return {}
    
    def export_logs(self, export_dir: str, start_date: Optional[datetime] = None,
                   end_date: Optional[datetime] = None):
        """로그를 지정된 디렉토리로 내보내기"""
        try:
            export_path = Path(export_dir)
            export_path.mkdir(exist_ok=True)
            
            # 날짜 필터링이 있는 경우 필터링된 로그 생성
            if start_date or end_date:
                self._export_filtered_logs(export_path, start_date, end_date)
            else:
                # 전체 로그 복사
                if self.search_log_path.exists():
                    import shutil
                    shutil.copy2(self.search_log_path, export_path / "search_log.csv")
                
                if self.print_log_path.exists():
                    import shutil
                    shutil.copy2(self.print_log_path, export_path / "print_log.csv")
            
            return True
            
        except Exception as e:
            print(f"로그 내보내기 실패: {e}")
            return False
    
    def _export_filtered_logs(self, export_path: Path, 
                             start_date: Optional[datetime],
                             end_date: Optional[datetime]):
        """날짜 필터링된 로그 내보내기"""
        # 검색 로그 필터링
        if self.search_log_path.exists():
            with open(self.search_log_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            filtered_rows = [rows[0]]  # 헤더 포함
            for row in rows[1:]:
                try:
                    log_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                    if self._is_date_in_range(log_date, start_date, end_date):
                        filtered_rows.append(row)
                except (ValueError, IndexError):
                    continue
            
            # 필터링된 검색 로그 저장
            with open(export_path / "search_log.csv", 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerows(filtered_rows)
        
        # 인쇄 로그 필터링
        if self.print_log_path.exists():
            with open(self.print_log_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            filtered_rows = [rows[0]]  # 헤더 포함
            for row in rows[1:]:
                try:
                    log_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                    if self._is_date_in_range(log_date, start_date, end_date):
                        filtered_rows.append(row)
                except (ValueError, IndexError):
                    continue
            
            # 필터링된 인쇄 로그 저장
            with open(export_path / "print_log.csv", 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerows(filtered_rows)
    
    def _is_date_in_range(self, date: datetime, 
                         start_date: Optional[datetime],
                         end_date: Optional[datetime]) -> bool:
        """날짜가 지정된 범위 내에 있는지 확인"""
        if start_date and date < start_date:
            return False
        if end_date and date > end_date:
            return False
        return True


# 전역 로거 인스턴스
logger = OrderLogger()
