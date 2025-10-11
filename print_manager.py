"""
PDF 인쇄 관리
- SumatraPDF CLI를 사용한 페이지별 인쇄
- 프린터 설정, 매수, 양면 옵션 지원
"""

import os
import subprocess
import sys
import tempfile
from typing import List, Optional, Dict, Any
from pathlib import Path
import win32print
import win32api
from pypdf import PdfReader, PdfWriter
from config_manager import config


class PrintManager:
    """PDF 인쇄 관리 클래스"""
    
    def __init__(self):
        self.sumatra_path = self._find_sumatra_path()
    
    def _find_sumatra_path(self) -> Optional[str]:
        """SumatraPDF 실행 파일을 찾기"""
        # 설정에서 경로 확인
        configured_path = config.get_sumatra_path()
        if configured_path and os.path.exists(configured_path):
            return configured_path
        
        # 일반적인 설치 경로들
        common_paths = [
            "SumatraPDF.exe",  # PATH에 있는 경우
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            r"C:\Users\{}\AppData\Local\SumatraPDF\SumatraPDF.exe".format(os.getenv('USERNAME', '')),
            # 현재 디렉토리
            os.path.join(os.getcwd(), "SumatraPDF.exe"),
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                config.set_sumatra_path(path)
                return path
        
        return None
    
    def get_available_printers(self) -> List[str]:
        """사용 가능한 프린터 목록 가져오기"""
        try:
            printers = []
            printer_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            for printer in printer_info:
                printers.append(printer[2])  # 프린터 이름
            return printers
        except Exception as e:
            print(f"프린터 목록 가져오기 실패: {e}")
            return []
    
    def get_default_printer(self) -> Optional[str]:
        """기본 프린터 이름 가져오기"""
        try:
            return win32print.GetDefaultPrinter()
        except Exception as e:
            print(f"기본 프린터 가져오기 실패: {e}")
            return None
    
    def print_pages(self, pdf_path: str, page_ranges: str, 
                   printer_name: Optional[str] = None,
                   copies: int = 1,
                   duplex: bool = False,
                   silent: bool = True) -> bool:
        """
        PDF의 특정 페이지들을 인쇄 - 임시 PDF 생성 방식
        
        Args:
            pdf_path: PDF 파일 경로
            page_ranges: 페이지 범위 (예: "1,3-5,7")
            printer_name: 프린터 이름 (None이면 기본 프린터)
            copies: 인쇄 매수
            duplex: 양면 인쇄 여부
            silent: 조용한 인쇄 (대화상자 없이)
            
        Returns:
            성공 여부
        """
        if not self.sumatra_path:
            raise RuntimeError("SumatraPDF를 찾을 수 없습니다. 경로를 설정해주세요.")
        
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
        
        # 프린터 이름 결정
        if not printer_name:
            printer_name = config.get_printer_name()
            if not printer_name:
                printer_name = self.get_default_printer()
        
        if not printer_name:
            raise RuntimeError("사용할 프린터를 찾을 수 없습니다.")
        
        # 페이지 범위가 있으면 임시 PDF 생성 후 인쇄
        if page_ranges and page_ranges != "all":
            temp_pdf_path = self._create_temp_pdf(pdf_path, page_ranges)
            if temp_pdf_path:
                try:
                    # 임시 PDF 전체 인쇄 (페이지 범위 지정 없이)
                    result = self._print_whole_pdf(temp_pdf_path, printer_name, copies, duplex, silent)
                    return result
                finally:
                    # 임시 파일 삭제
                    try:
                        os.unlink(temp_pdf_path)
                    except:
                        pass
            else:
                print("Temp PDF creation failed - trying original method")
        
        # 기본 방식: SumatraPDF 페이지 범위 지정
        return self._print_with_page_range(pdf_path, page_ranges, printer_name, copies, duplex, silent)
    
    def _create_temp_pdf(self, pdf_path: str, page_ranges: str) -> Optional[str]:
        """특정 페이지들만 포함한 임시 PDF 생성"""
        try:
            # 페이지 번호 파싱
            page_numbers = self._parse_page_ranges(page_ranges)
            if not page_numbers:
                return None
            
            print(f"Creating temp PDF with pages: {page_numbers}")
            
            # 원본 PDF 읽기
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            
            # 페이지 번호 유효성 검사 (1-based → 0-based)
            valid_pages = []
            for page_num in page_numbers:
                if 1 <= page_num <= total_pages:
                    valid_pages.append(page_num - 1)  # 0-based 인덱스
                    
            if not valid_pages:
                print(f"No valid pages: {page_numbers}")
                return None
            
            # 새 PDF 생성
            writer = PdfWriter()
            for page_idx in valid_pages:
                writer.add_page(reader.pages[page_idx])
            
            # 임시 파일로 저장
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf', prefix='print_')
            os.close(temp_fd)
            
            with open(temp_path, 'wb') as temp_file:
                writer.write(temp_file)
            
            print(f"Temp PDF created: {temp_path}")
            return temp_path
            
        except Exception as e:
            print(f"Temp PDF creation failed: {e}")
            return None
    
    def _parse_page_ranges(self, page_ranges: str) -> List[int]:
        """페이지 범위 문자열을 페이지 번호 리스트로 변환"""
        page_numbers = []
        
        try:
            for part in page_ranges.split(','):
                part = part.strip()
                if '-' in part:
                    # 범위: "3-7"
                    start, end = map(int, part.split('-', 1))
                    page_numbers.extend(range(start, end + 1))
                else:
                    # 단일 페이지: "5"
                    page_numbers.append(int(part))
                    
            return sorted(set(page_numbers))  # 중복 제거 및 정렬
            
        except Exception as e:
            print(f"Page range parsing failed: {e}")
            return []
    
    def _print_whole_pdf(self, pdf_path: str, printer_name: str, copies: int, duplex: bool, silent: bool) -> bool:
        """전체 PDF 인쇄 (페이지 범위 지정 없이)"""
        cmd = [self.sumatra_path]
        
        if silent:
            cmd.append("-silent")
        
        # 프린터 지정
        cmd.extend(["-print-to", printer_name])
        
        # 인쇄 설정 (페이지 범위 제외)
        print_settings = []
        
        # 매수
        if copies > 1:
            print_settings.append(f"copies:{copies}")
        
        # 양면 인쇄
        if duplex:
            print_settings.append("duplex")
        
        if print_settings:
            cmd.extend(["-print-settings", ",".join(print_settings)])
        
        # PDF 파일 경로
        cmd.append(pdf_path)
        
        return self._execute_sumatra_command(cmd)
    
    def _print_with_page_range(self, pdf_path: str, page_ranges: str, printer_name: str, copies: int, duplex: bool, silent: bool) -> bool:
        """기존 방식: 페이지 범위 지정 인쇄"""
        cmd = [self.sumatra_path]
        
        if silent:
            cmd.append("-silent")
        
        # 프린터 지정
        cmd.extend(["-print-to", printer_name])
        
        # 인쇄 설정
        print_settings = []
        
        # 페이지 범위
        if page_ranges:
            print_settings.append(f"pages:{page_ranges}")
        
        # 매수
        if copies > 1:
            print_settings.append(f"copies:{copies}")
        
        # 양면 인쇄
        if duplex:
            print_settings.append("duplex")
        
        if print_settings:
            cmd.extend(["-print-settings", ",".join(print_settings)])
        
        # PDF 파일 경로
        cmd.append(pdf_path)
        
        return self._execute_sumatra_command(cmd)
    
    def _execute_sumatra_command(self, cmd: List[str]) -> bool:
        """SumatraPDF 명령 실행"""
        
        try:
            # 디버깅: 실행할 명령어 출력
            cmd_str = " ".join([f'"{arg}"' if " " in arg else arg for arg in cmd])
            print(f"SumatraPDF command:")
            print(f"   {cmd_str}")
            
            # SumatraPDF 실행
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=60)
            
            print(f"SumatraPDF result code: {result.returncode}")
            if result.stdout:
                print(f"   출력: {result.stdout}")
            if result.stderr:
                print(f"   오류: {result.stderr}")
            
            if result.returncode == 0:
                return True
            else:
                print(f"인쇄 실패 (코드: {result.returncode})")
                return False
                
        except subprocess.TimeoutExpired:
            print("인쇄 작업 시간 초과")
            return False
        except Exception as e:
            print(f"인쇄 중 오류: {e}")
            return False
    
    def print_dialog(self, pdf_path: str) -> bool:
        """
        인쇄 대화상자를 통한 인쇄
        
        Args:
            pdf_path: PDF 파일 경로
            
        Returns:
            성공 여부
        """
        if not self.sumatra_path:
            raise RuntimeError("SumatraPDF를 찾을 수 없습니다.")
        
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
        
        try:
            # 인쇄 대화상자 열기
            cmd = [self.sumatra_path, "-print-dialog", pdf_path]
            
            result = subprocess.run(cmd, timeout=300)  # 5분 타임아웃
            return result.returncode == 0
            
        except subprocess.TimeoutExpired:
            print("인쇄 대화상자 시간 초과")
            return False
        except Exception as e:
            print(f"인쇄 대화상자 오류: {e}")
            return False
    
    def test_printer(self, printer_name: str) -> bool:
        """프린터 연결 테스트"""
        try:
            # 프린터가 목록에 있는지 확인
            available_printers = self.get_available_printers()
            if printer_name not in available_printers:
                return False
            
            # 프린터 핸들 열어보기
            handle = win32print.OpenPrinter(printer_name)
            win32print.ClosePrinter(handle)
            return True
            
        except Exception as e:
            print(f"프린터 테스트 실패 ({printer_name}): {e}")
            return False
    
    def get_printer_info(self, printer_name: str) -> Optional[Dict[str, Any]]:
        """프린터 정보 가져오기"""
        try:
            handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(handle, 2)
            win32print.ClosePrinter(handle)
            
            return {
                'name': printer_info.get('pPrinterName', ''),
                'driver': printer_info.get('pDriverName', ''),
                'port': printer_info.get('pPortName', ''),
                'location': printer_info.get('pLocation', ''),
                'comment': printer_info.get('pComment', ''),
                'status': printer_info.get('Status', 0),
            }
            
        except Exception as e:
            print(f"프린터 정보 가져오기 실패 ({printer_name}): {e}")
            return None
    
    def is_sumatra_available(self) -> bool:
        """SumatraPDF 사용 가능 여부 확인"""
        return self.sumatra_path is not None and os.path.exists(self.sumatra_path)
    
    def get_sumatra_version(self) -> Optional[str]:
        """SumatraPDF 버전 정보 가져오기"""
        if not self.sumatra_path:
            return None
        
        try:
            result = subprocess.run([self.sumatra_path, "-version"], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=10)
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                return None
                
        except Exception:
            return None
    
    def set_sumatra_path(self, path: str) -> bool:
        """SumatraPDF 경로 설정"""
        if os.path.exists(path):
            self.sumatra_path = path
            config.set_sumatra_path(path)
            return True
        else:
            return False
    
    def validate_page_ranges(self, page_ranges: str, total_pages: int) -> tuple[bool, str]:
        """
        페이지 범위 유효성 검사
        
        Args:
            page_ranges: 페이지 범위 문자열 (예: "1,3-5,7")
            total_pages: 총 페이지 수
            
        Returns:
            (유효여부, 오류메시지)
        """
        if not page_ranges.strip():
            return False, "페이지 범위가 비어있습니다."
        
        try:
            # 페이지 범위 파싱
            ranges = page_ranges.split(',')
            all_pages = set()
            
            for range_str in ranges:
                range_str = range_str.strip()
                
                if '-' in range_str:
                    # 범위 (예: 3-5)
                    start, end = map(int, range_str.split('-', 1))
                    if start > end:
                        return False, f"잘못된 범위: {range_str} (시작이 끝보다 큽니다)"
                    
                    for page in range(start, end + 1):
                        if page < 1 or page > total_pages:
                            return False, f"페이지 {page}는 범위를 벗어났습니다 (1-{total_pages})"
                        all_pages.add(page)
                else:
                    # 단일 페이지
                    page = int(range_str)
                    if page < 1 or page > total_pages:
                        return False, f"페이지 {page}는 범위를 벗어났습니다 (1-{total_pages})"
                    all_pages.add(page)
            
            if not all_pages:
                return False, "유효한 페이지가 없습니다."
            
            return True, ""
            
        except ValueError as e:
            return False, f"페이지 범위 형식이 잘못되었습니다: {e}"
        except Exception as e:
            return False, f"페이지 범위 검증 오류: {e}"


class PrintJob:
    """인쇄 작업 정보"""
    
    def __init__(self, pdf_path: str, page_ranges: str, 
                 printer_name: str = "", copies: int = 1, duplex: bool = False):
        self.pdf_path = pdf_path
        self.page_ranges = page_ranges
        self.printer_name = printer_name
        self.copies = copies
        self.duplex = duplex
        self.created_at = None
        self.completed_at = None
        self.status = "pending"  # pending, printing, completed, failed
        self.error_message = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """딕셔너리로 변환 (로깅용)"""
        return {
            'pdf_path': self.pdf_path,
            'page_ranges': self.page_ranges,
            'printer_name': self.printer_name,
            'copies': self.copies,
            'duplex': self.duplex,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'status': self.status,
            'error_message': self.error_message
        }
