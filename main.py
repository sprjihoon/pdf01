"""
PDF Excel Matcher - 메인 GUI 애플리케이션
엑셀 구매자 순서대로 PDF 페이지 자동 정렬
"""

import sys
import os
import traceback
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit,
    QProgressBar, QGroupBox, QMessageBox, QCheckBox, QSpinBox,
    QTabWidget, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QFont

from io_utils import load_excel, get_output_filenames, save_pdf, save_report, is_text_based_pdf
from matcher import extract_pages, match_rows_to_pages, reorder_pdf
from order_searcher import OrderSearcher
from print_manager import PrintManager
from order_logger import logger
from config_manager import config
import time


class ProcessingWorker(QThread):
    """백그라운드 작업 스레드 - PDF 정렬용"""
    progress = Signal(str)  # 진행 상황 메시지
    finished = Signal(dict)  # 완료 (결과 데이터)
    error = Signal(str)  # 오류
    
    def __init__(self, excel_path, pdf_path, output_dir, use_fuzzy, threshold):
        super().__init__()
        self.excel_path = excel_path
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        self.use_fuzzy = use_fuzzy
        self.threshold = threshold
    
    def run(self):
        try:
            # 1. PDF 텍스트 기반 확인
            self.progress.emit("📄 PDF 파일 확인 중...")
            is_text, msg = is_text_based_pdf(self.pdf_path)
            if not is_text:
                self.error.emit(msg)
                return
            self.progress.emit(f"✓ {msg}")
            
            # 2. 엑셀 로드
            self.progress.emit("📊 엑셀 파일 로드 중...")
            df = load_excel(self.excel_path)
            self.progress.emit(f"✓ 엑셀 로드 완료 ({len(df)}행)")
            
            # 3. PDF 텍스트 추출
            self.progress.emit("📖 PDF 텍스트 추출 중...")
            pages = extract_pages(self.pdf_path)
            self.progress.emit(f"✓ PDF 텍스트 추출 완료 ({len(pages)}페이지)")
            
            # 4. 매칭
            self.progress.emit("🔍 엑셀-PDF 매칭 중...")
            fuzzy_text = f" (유사도 매칭: ON, 임계값: {self.threshold})" if self.use_fuzzy else " (정확 일치만)"
            self.progress.emit(f"   매칭 옵션{fuzzy_text}")
            
            assignments, leftover_pages, match_details = match_rows_to_pages(
                df, pages, self.use_fuzzy, self.threshold
            )
            
            matched_count = len(assignments)
            unmatched_count = len(df) - matched_count
            leftover_count = len(leftover_pages)
            
            self.progress.emit(f"✓ 매칭 완료: {matched_count}건 성공, {unmatched_count}건 실패, {leftover_count}페이지 남음")
            self.progress.emit("")
            self.progress.emit("📋 매칭 상세 (처음 10건):")
            for i, row_idx in enumerate(range(min(10, len(df)))):
                row = df.iloc[row_idx]
                detail = match_details.get(row_idx, {'page_idx': -1, 'score': 0, 'reason': 'no_match'})
                if detail['page_idx'] >= 0:
                    order_num = row.get('주문번호', 'N/A')
                    self.progress.emit(
                        f"   엑셀 {row_idx+1}행 (주문번호: {order_num}) → "
                        f"PDF {detail['page_idx']+1}페이지 (점수: {detail['score']:.1f})"
                    )
                else:
                    order_num = row.get('주문번호', 'N/A')
                    self.progress.emit(f"   엑셀 {row_idx+1}행 (주문번호: {order_num}) → 매칭 실패")
            if len(df) > 10:
                self.progress.emit(f"   ... 외 {len(df)-10}건")
            
            # 5. 페이지 순서 결정
            self.progress.emit("")
            self.progress.emit("📑 페이지 순서 결정 중...")
            ordered_indices = []
            
            # 매칭된 페이지 (엑셀 순서대로)
            for row_idx in range(len(df)):
                if row_idx in assignments:
                    page_idx = assignments[row_idx]
                    ordered_indices.append(page_idx)
            
            # 미매칭 페이지 (원본 순서대로 뒤에 추가)
            ordered_indices.extend(leftover_pages)
            
            self.progress.emit(f"✓ 총 {len(ordered_indices)}페이지 정렬 준비 완료")
            self.progress.emit("")
            self.progress.emit("📄 결과 PDF 페이지 순서 (처음 10개):")
            for i, page_idx in enumerate(ordered_indices[:10]):
                if i < len(assignments):
                    self.progress.emit(f"   결과 {i+1}페이지 ← 원본 {page_idx+1}페이지 (엑셀 {i+1}행과 매칭됨)")
                else:
                    self.progress.emit(f"   결과 {i+1}페이지 ← 원본 {page_idx+1}페이지 (미매칭)")
            if len(ordered_indices) > 10:
                self.progress.emit(f"   ... 외 {len(ordered_indices)-10}페이지")
            
            # 6. 출력 파일명 결정
            self.progress.emit("📁 출력 파일명 생성 중...")
            pdf_out_path, csv_out_path = get_output_filenames(self.output_dir)
            self.progress.emit(f"✓ PDF: {os.path.basename(pdf_out_path)}")
            self.progress.emit(f"✓ CSV: {os.path.basename(csv_out_path)}")
            
            # 7. PDF 재정렬
            self.progress.emit("💾 PDF 저장 중...")
            reorder_pdf(self.pdf_path, ordered_indices, pdf_out_path)
            self.progress.emit(f"✓ PDF 저장 완료")
            
            # 8. 리포트 생성
            self.progress.emit("📝 리포트 생성 중...")
            report_rows = []
            
            for row_idx in range(len(df)):
                row = df.iloc[row_idx]
                detail = match_details.get(row_idx, {'page_idx': -1, 'score': 0, 'reason': 'no_match'})
                
                report_row = {
                    '엑셀행번호': row_idx + 2,  # 엑셀 행번호 (헤더 제외, 1-based)
                    '매칭페이지': detail['page_idx'] + 1 if detail['page_idx'] >= 0 else 'UNMATCHED',
                    '점수': round(detail['score'], 1),
                    '매칭키': detail['reason'],
                    '구매자명': row['구매자명'],
                    '전화번호': row['전화번호'],
                    '주소': row['주소'],
                    '주문번호': row['주문번호']
                }
                report_rows.append(report_row)
            
            save_report(report_rows, csv_out_path)
            self.progress.emit(f"✓ 리포트 저장 완료")
            
            # 9. 완료
            self.progress.emit("")
            self.progress.emit("=" * 60)
            self.progress.emit("✅ 모든 작업이 완료되었습니다!")
            self.progress.emit("=" * 60)
            self.progress.emit(f"📊 매칭 통계:")
            self.progress.emit(f"   - 성공: {matched_count}건")
            self.progress.emit(f"   - 실패: {unmatched_count}건")
            self.progress.emit(f"   - 남은 페이지: {leftover_count}개")
            self.progress.emit(f"📁 출력 파일:")
            self.progress.emit(f"   - PDF: {pdf_out_path}")
            self.progress.emit(f"   - CSV: {csv_out_path}")
            
            # 결과 전달
            self.finished.emit({
                'pdf_path': pdf_out_path,
                'csv_path': csv_out_path,
                'matched_count': matched_count,
                'unmatched_count': unmatched_count,
                'leftover_count': leftover_count,
                'total_excel_rows': len(df),
                'total_pdf_pages': len(pages)
            })
            
        except Exception as e:
            error_msg = f"❌ 오류 발생:\n{str(e)}\n\n상세 정보:\n{traceback.format_exc()}"
            self.error.emit(error_msg)


class OrderSearchWorker(QThread):
    """주문번호 검색 백그라운드 작업 스레드"""
    progress = Signal(str)  # 진행 상황 메시지
    finished = Signal(object)  # 완료 (SearchResult 또는 None)
    error = Signal(str)  # 오류
    
    def __init__(self, order_number, folder_path):
        super().__init__()
        self.order_number = order_number
        self.folder_path = folder_path
        self.order_searcher = OrderSearcher()
    
    def run(self):
        try:
            self.progress.emit(f"🔍 검색 시작: {self.order_number}")
            self.progress.emit(f"📁 대상 폴더: {self.folder_path}")
            
            # PDF 파일 목록 가져오기
            self.progress.emit("📄 PDF 파일 목록 확인 중...")
            pdf_files = self.order_searcher._find_pdf_files(self.folder_path)
            self.progress.emit(f"✓ {len(pdf_files)}개 PDF 파일 발견")
            
            if not pdf_files:
                self.progress.emit("❌ PDF 파일이 없습니다")
                self.finished.emit(None)
                return
            
            # 파일별로 검색 (진행률 표시)
            matches = []
            total_files = len(pdf_files)
            
            for i, pdf_file in enumerate(pdf_files):
                try:
                    filename = os.path.basename(pdf_file)
                    self.progress.emit(f"🔍 검색 중... ({i+1}/{total_files}) {filename}")
                    
                    match = self.order_searcher._search_order_in_file(pdf_file, self.order_number)
                    if match:
                        matches.append(match)
                        self.progress.emit(f"✅ 발견: {filename}")
                    
                    # 주기적으로 진행 상황 업데이트
                    if (i + 1) % 10 == 0 or i == total_files - 1:
                        progress = int((i + 1) / total_files * 100)
                        self.progress.emit(f"📊 진행률: {progress}% ({i+1}/{total_files})")
                        
                except Exception as e:
                    self.progress.emit(f"⚠️ 파일 오류 ({filename}): {str(e)}")
                    continue
            
            # 검색 완료
            if matches:
                self.progress.emit(f"✅ 검색 완료: {len(matches)}개 파일에서 발견")
                
                # 최신 파일 선택
                best_match, decided_by = self.order_searcher._select_latest_file(matches)
                
                from order_searcher import SearchResult
                search_result = SearchResult(
                    order_number=self.order_number,
                    best_match=best_match,
                    all_matches=matches,
                    decided_by=decided_by
                )
                
                self.progress.emit(f"🎯 최신 파일: {os.path.basename(best_match.file_path)} ({decided_by} 기준)")
                self.finished.emit(search_result)
            else:
                self.progress.emit(f"❌ '{self.order_number}'를 찾을 수 없습니다")
                self.finished.emit(None)
                
        except Exception as e:
            error_msg = f"❌ 검색 중 오류 발생:\n{str(e)}\n\n상세 정보:\n{traceback.format_exc()}"
            self.error.emit(error_msg)


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Excel Matcher - 구매자 순서로 PDF 자동 정렬")
        
        # 전체화면으로 설정
        self.showMaximized()  # 최대화 모드로 시작
        self.setMinimumSize(1200, 800)  # 최소 크기도 증가
        
        # 설정 관리자 (Windows 레지스트리, macOS/Linux는 적절한 위치에 저장)
        self.settings = QSettings("PDFExcelMatcher", "PathSettings")
        
        # 메인 위젯 및 레이아웃
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(10)  # 0 → 10으로 여백 추가하여 겹침 방지
        main_layout.setContentsMargins(15, 15, 15, 15)  # 여백 확대
        
        # 상단 통합 경로 영역 생성
        self.create_base_path_section(main_layout)
        
        # 탭 위젯 생성 (여백 추가)
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget {
                margin-top: 10px;
            }
            QTabWidget::pane {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 10px;
            }
        """)
        main_layout.addWidget(self.tab_widget)
        
        # PDF 정렬 탭 생성
        self.create_pdf_sort_tab()
        
        # 주문번호 검색 탭 생성
        self.create_order_search_tab()
        
        # 초기 경로 설정 확인
        self.check_initial_path()
    
    def create_base_path_section(self, parent_layout):
        """상단 통합 경로 선택 영역 생성"""
        # 구분선이 있는 프레임
        path_frame = QFrame()
        path_frame.setFrameStyle(QFrame.StyledPanel)
        path_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                margin-bottom: 15px;
                padding: 5px;
            }
        """)
        
        path_layout = QVBoxLayout(path_frame)
        path_layout.setSpacing(12)
        path_layout.setContentsMargins(15, 15, 15, 15)  # 내부 여백 확대
        
        # 제목
        title_layout = QHBoxLayout()
        title_label = QLabel("📂 통합 작업 경로")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_layout.addWidget(title_label)
        
        # 날짜별 폴더 옵션
        self.date_subfolder_check = QCheckBox("날짜별 하위폴더 사용")
        self.date_subfolder_check.setToolTip("활성화시 선택한 폴더 아래에 YYYY-MM-DD 폴더를 자동 생성합니다")
        self.date_subfolder_check.stateChanged.connect(self.on_date_subfolder_changed)
        title_layout.addStretch()
        title_layout.addWidget(self.date_subfolder_check)
        
        path_layout.addLayout(title_layout)
        
        # 경로 선택 영역 (한 줄로 깔끔하게)
        path_select_layout = QHBoxLayout()
        path_select_layout.addWidget(QLabel("현재 경로:"))
        
        self.current_path_label = QLabel("경로가 설정되지 않았습니다")
        self.current_path_label.setStyleSheet("""
            QLabel {
                background-color: #ffffff;
                border: 2px solid #2196F3;
                padding: 12px 15px;
                border-radius: 6px;
                font-family: 'Arial', 'Malgun Gothic', sans-serif;
                font-size: 12pt;
                color: #000000;
                font-weight: bold;
            }
        """)
        self.current_path_label.setMinimumHeight(45)
        self.current_path_label.setWordWrap(False)
        self.current_path_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.current_path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        path_select_layout.addWidget(self.current_path_label, 1)
        
        # 경로 선택 버튼 (바로 옆에 붙이기)
        self.select_path_btn = QPushButton("📂 경로 선택")
        self.select_path_btn.setMinimumSize(100, 45)
        self.select_path_btn.clicked.connect(self.select_base_path)
        path_select_layout.addWidget(self.select_path_btn, 0)
        
        path_layout.addLayout(path_select_layout)
        
        # 상태 메시지
        self.path_status_label = QLabel("")
        self.path_status_label.setStyleSheet("color: #666; font-size: 9pt; margin-top: 5px;")
        path_layout.addWidget(self.path_status_label)
        
        parent_layout.addWidget(path_frame)
        
        # 최근 경로 기능 제거함
    
    def check_initial_path(self):
        """초기 경로 설정 확인 및 로드"""
        base_path = config.get_base_path()
        
        if base_path:
            # 기존 경로가 있으면 유효성 검사
            is_valid, message = config.validate_base_path(base_path)
            if is_valid:
                self.update_path_display(base_path)
                self.path_status_label.setText(f"✅ {message}")
                self.path_status_label.setStyleSheet("color: #28a745; font-size: 9pt;")
            else:
                self.path_status_label.setText(f"⚠️ {message}")
                self.path_status_label.setStyleSheet("color: #ffc107; font-size: 9pt;")
                self.show_path_selection_dialog()
        else:
            # 경로가 설정되지 않았으면 선택 요청
            self.show_path_selection_dialog()
        
        # 날짜별 폴더 옵션 로드
        self.date_subfolder_check.setChecked(config.get_use_date_subfolder())
    
    def show_path_selection_dialog(self):
        """경로 선택 대화상자 표시"""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("작업 폴더 설정")
        msg.setText("PDF 매칭 프로그램에 오신 것을 환영합니다!")
        msg.setInformativeText(
            "모든 PDF 파일 검색과 결과 저장을 위한 기본 폴더를 선택해주세요.\n"
            "이 폴더는 다음 용도로 사용됩니다:\n\n"
            "• PDF 파일 검색\n"
            "• 매칭 결과 저장 (ordered_*.pdf)\n"
            "• 리포트 저장 (*.csv)\n"
            "• 로그 파일 저장\n\n"
            "한번 설정하면 자동으로 기억됩니다."
        )
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()
        
        self.select_base_path()
    
    def select_base_path(self):
        """기본 경로 선택"""
        current_path = config.get_base_path()
        start_dir = current_path if current_path and os.path.exists(current_path) else os.path.expanduser("~")
        
        selected_path = QFileDialog.getExistingDirectory(
            self, "PDF 작업 폴더 선택", start_dir
        )
        
        if selected_path:
            # 경로 유효성 검사
            is_valid, message = config.validate_base_path(selected_path)
            
            if is_valid:
                config.set_base_path(selected_path)
                self.update_path_display(selected_path)
                self.update_recent_paths_combo()
                self.path_status_label.setText(f"✅ {message}")
                self.path_status_label.setStyleSheet("color: #28a745; font-size: 9pt;")
                self.log(f"✓ 작업 폴더 설정: {selected_path}")
                self.search_log(f"✓ 작업 폴더 설정: {selected_path}")
            else:
                self.path_status_label.setText(f"❌ {message}")
                self.path_status_label.setStyleSheet("color: #dc3545; font-size: 9pt;")
                
                # 폴더 생성 제안
                if "존재하지 않습니다" in message:
                    reply = QMessageBox.question(
                        self, "폴더 생성", 
                        f"선택한 폴더가 존재하지 않습니다.\n{selected_path}\n\n폴더를 생성하시겠습니까?",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    
                    if reply == QMessageBox.Yes:
                        success, create_message = config.create_base_path(selected_path)
                        if success:
                            config.set_base_path(selected_path)
                            self.update_path_display(selected_path)
                            self.update_recent_paths_combo()
                            self.path_status_label.setText(f"✅ {create_message}")
                            self.path_status_label.setStyleSheet("color: #28a745; font-size: 9pt;")
                        else:
                            QMessageBox.critical(self, "오류", create_message)
                else:
                    QMessageBox.critical(self, "경로 오류", message)
    
    def update_path_display(self, path: str):
        """경로 표시 업데이트 - 개선된 버전"""
        if path:
            # 경로 축약 로직 개선
            display_path = path
            
            # 경로 축약을 최대한 피함 (200자까지 전체 표시)
            if len(display_path) > 200:  # 200자까지 허용 (최대한 확장)
                parts = display_path.split('\\')
                if len(parts) > 5:
                    # 첫 부분과 마지막 부분만 표시
                    display_path = f"{parts[0]}\\{parts[1]}\\...\\{parts[-2]}\\{parts[-1]}"
                elif len(parts) > 3:
                    # 드라이브:\...\마지막2개폴더
                    display_path = f"{parts[0]}\\...\\{parts[-2]}\\{parts[-1]}"
                else:
                    # 거의 전체 표시
                    display_path = "..." + display_path[-147:]
            
            self.current_path_label.setText(display_path)
            self.current_path_label.setToolTip(f"전체 경로: {path}")
            
            # 작업 경로 표시 (날짜 폴더 포함 시)
            working_path = config.get_working_path()
            if working_path != path:
                working_display = working_path
                if len(working_display) > 50:
                    working_parts = working_display.split('\\')
                    if len(working_parts) > 2:
                        working_display = f"...\\{working_parts[-2]}\\{working_parts[-1]}"
                
                self.path_status_label.setText(f"📁 실제 작업 경로: {working_display}")
                self.path_status_label.setStyleSheet("color: #17a2b8; font-size: 9pt;")
                self.path_status_label.setToolTip(f"전체 작업 경로: {working_path}")
            else:
                self.path_status_label.setToolTip("")
        else:
            self.current_path_label.setText("경로가 설정되지 않았습니다")
            self.current_path_label.setToolTip("")
    
    # 최근 경로 기능 제거됨
    
    def on_date_subfolder_changed(self, state):
        """날짜별 하위폴더 옵션 변경"""
        use_date = state == Qt.Checked
        config.set_use_date_subfolder(use_date)
        
        # 현재 경로 다시 표시 (작업 경로 변경 반영)
        current_path = config.get_base_path()
        if current_path:
            self.update_path_display(current_path)
            
            if use_date:
                self.search_log("✓ 날짜별 하위폴더 사용 활성화")
            else:
                self.search_log("✓ 날짜별 하위폴더 사용 비활성화")
        
    def create_pdf_sort_tab(self):
        """PDF 정렬 탭 생성"""
        tab_widget = QWidget()
        layout = QVBoxLayout(tab_widget)
        layout.setSpacing(15)
        
        # 제목
        title = QLabel("📄 PDF Excel Matcher")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        subtitle = QLabel("엑셀 구매자 순서대로 PDF 페이지를 자동 정렬합니다")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; font-size: 11pt;")
        layout.addWidget(subtitle)
        
        # 파일 선택 그룹
        file_group = QGroupBox("📁 파일 선택")
        file_layout = QVBoxLayout()
        
        # 엑셀
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("엑셀 파일:"))
        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText("주문번호 컬럼이 있는 엑셀 파일...")
        excel_layout.addWidget(self.excel_edit)
        excel_btn = QPushButton("찾아보기")
        excel_btn.clicked.connect(self.browse_excel)
        excel_layout.addWidget(excel_btn)
        
        file_layout.addLayout(excel_layout)
        
        # PDF
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(QLabel("PDF 파일:"))
        self.pdf_edit = QLineEdit()
        self.pdf_edit.setPlaceholderText("정렬할 PDF 파일 (텍스트 기반)...")
        pdf_layout.addWidget(self.pdf_edit)
        pdf_btn = QPushButton("찾아보기")
        pdf_btn.clicked.connect(self.browse_pdf)
        pdf_layout.addWidget(pdf_btn)
        
        file_layout.addLayout(pdf_layout)
        
        # 출력 폴더는 작업 폴더로 자동 설정되므로 제거
        output_info_layout = QHBoxLayout()
        output_info_layout.addWidget(QLabel("결과 저장:"))
        output_info_label = QLabel("작업 폴더에 자동 저장 (ordered_YYYYMMDD.pdf, match_report.csv)")
        output_info_label.setStyleSheet("color: #666; font-style: italic; padding: 8px; border: 1px solid #ddd; border-radius: 4px; background-color: #f9f9f9;")
        output_info_layout.addWidget(output_info_label)
        file_layout.addLayout(output_info_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 옵션 그룹
        option_group = QGroupBox("⚙️ 매칭 옵션")
        option_layout = QVBoxLayout()
        
        # 유사도 매칭 체크박스
        fuzzy_layout = QHBoxLayout()
        self.fuzzy_check = QCheckBox("유사도 매칭 사용 (보조)")
        self.fuzzy_check.setChecked(True)  # 기본값을 True로 변경
        self.fuzzy_check.stateChanged.connect(self.on_fuzzy_changed)
        fuzzy_layout.addWidget(self.fuzzy_check)
        
        fuzzy_layout.addWidget(QLabel("임계값:"))
        self.threshold_spin = QSpinBox()
        self.threshold_spin.setMinimum(50)
        self.threshold_spin.setMaximum(100)
        self.threshold_spin.setValue(90)
        self.threshold_spin.setSuffix("%")
        self.threshold_spin.setEnabled(True)  # 기본값이 True이므로 활성화
        fuzzy_layout.addWidget(self.threshold_spin)
        
        fuzzy_layout.addStretch()
        option_layout.addLayout(fuzzy_layout)
        
        # 옵션 설명
        help_label = QLabel(
            "💡 매칭 기준: 주문번호 (고유번호로 1:1 매칭)\n"
            "   • 기본: 정확 일치만 인정 (권장)\n"
            "   • 유사도 매칭: 주문번호에 오타가 있을 때 보조적으로 사용 (임계값 조정 가능)"
        )
        help_label.setStyleSheet("color: #555; font-size: 9pt; padding: 10px; background: #f5f5f5; border-radius: 5px;")
        option_layout.addWidget(help_label)
        
        option_group.setLayout(option_layout)
        layout.addWidget(option_group)
        
        # 실행 버튼 (크기 축소)
        self.run_btn = QPushButton("▶️ PDF 정렬 실행")
        self.run_btn.setMinimumHeight(40)
        self.run_btn.setMaximumHeight(40)
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.run_btn.clicked.connect(self.run_processing)
        layout.addWidget(self.run_btn)
        
        # 진행 상황 그룹
        progress_group = QGroupBox("📊 진행 상황")
        progress_layout = QVBoxLayout()
        
        # 프로그레스 바
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(0)  # Indeterminate
        self.progress_bar.setTextVisible(True)
        self.progress_bar.hide()
        progress_layout.addWidget(self.progress_bar)
        
        # 로그
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(250)
        self.log_text.setStyleSheet("font-family: 'Consolas', monospace; font-size: 9pt;")
        progress_layout.addWidget(self.log_text)
        
        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)
        
        # 초기 로그
        self.log("프로그램이 준비되었습니다.")
        self.log("파일을 선택하고 [PDF 정렬 실행] 버튼을 눌러주세요.")
        
        # 작업 스레드
        self.worker = None
        
        # 탭에 추가
        self.tab_widget.addTab(tab_widget, "📄 PDF 정렬")
        
        # 저장된 경로 로드
        self.load_saved_paths()
    
    def create_order_search_tab(self):
        """주문번호 검색 탭 생성"""
        tab_widget = QWidget()
        layout = QVBoxLayout(tab_widget)
        layout.setSpacing(15)
        
        # 제목
        title = QLabel("🔍 주문번호 검색 & 인쇄")
        title_font = QFont()
        title_font.setPointSize(14)  # 18 → 14로 축소
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        subtitle = QLabel("PDF 파일에서 주문번호를 검색하고 해당 페이지를 인쇄합니다")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; font-size: 11pt;")
        layout.addWidget(subtitle)
        
        # 검색 정보 표시 (폴더 선택 불필요 - 통합 경로 사용)
        folder_group = QGroupBox("📋 검색 설정")
        folder_layout = QVBoxLayout()
        
        folder_info_layout = QHBoxLayout()
        folder_info_layout.addWidget(QLabel("검색 대상:"))
        search_info_label = QLabel("상단에 설정된 작업 폴더에서 PDF 파일 검색\n💡 원본 PDF가 있는 폴더로 설정하세요! (예: 카카오톡 받은 파일)")
        search_info_label.setStyleSheet("color: #666; font-style: italic; padding: 10px; border: 1px solid #17a2b8; border-radius: 4px; background-color: #e3f2fd;")
        folder_info_layout.addWidget(search_info_label)
        folder_layout.addLayout(folder_info_layout)
        
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        
        # 검색 그룹
        search_group = QGroupBox("🔍 주문번호 검색")
        search_layout = QVBoxLayout()
        
        # 주문번호 입력 및 검색
        order_input_layout = QHBoxLayout()
        order_input_layout.addWidget(QLabel("주문번호:"))
        
        self.order_number_edit = QLineEdit()
        self.order_number_edit.setPlaceholderText("예: 800017 (뒷자리만 입력)")
        self.order_number_edit.returnPressed.connect(self.search_order)  # Enter 키 지원
        order_input_layout.addWidget(self.order_number_edit)
        
        # 검색 버튼
        self.search_btn = QPushButton("🔍 검색")
        self.search_btn.setMinimumHeight(35)
        self.search_btn.setMinimumWidth(80)
        self.search_btn.clicked.connect(self.search_order)
        order_input_layout.addWidget(self.search_btn)
        
        # 검색 중지 버튼 (검색 버튼 옆에)
        self.stop_search_btn = QPushButton("⏹️ 중지")
        self.stop_search_btn.setMinimumHeight(35)
        self.stop_search_btn.setMinimumWidth(80)
        self.stop_search_btn.setEnabled(False)
        self.stop_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                border-radius: 8px;
                font-size: 11pt;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.stop_search_btn.clicked.connect(self.stop_search)
        order_input_layout.addWidget(self.stop_search_btn)
        search_layout.addLayout(order_input_layout)
        
        # 검색 결과 테이블
        self.search_result_table = QTableWidget()
        self.search_result_table.setColumnCount(6)
        self.search_result_table.setHorizontalHeaderLabels([
            "파일명", "문서날짜", "파일명날짜", "수정시간", "페이지", "선택기준"
        ])
        self.search_result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.search_result_table.setMaximumHeight(150)
        search_layout.addWidget(self.search_result_table)
        
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)
        
        # 인쇄 설정 그룹
        print_group = QGroupBox("🖨️ 인쇄 설정")
        print_layout = QVBoxLayout()
        
        # 프린터 선택
        printer_layout = QHBoxLayout()
        printer_layout.addWidget(QLabel("프린터:"))
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(200)
        printer_layout.addWidget(self.printer_combo)
        
        refresh_printer_btn = QPushButton("새로고침")
        refresh_printer_btn.clicked.connect(self.refresh_printers)
        printer_layout.addWidget(refresh_printer_btn)
        printer_layout.addStretch()
        print_layout.addLayout(printer_layout)
        
        # PDF 뷰어 선택 옵션 추가
        viewer_layout = QHBoxLayout()
        viewer_layout.addWidget(QLabel("PDF 뷰어:"))
        self.viewer_combo = QComboBox()
        self.viewer_combo.addItem("🚀 SumatraPDF (빠름)", "sumatra")
        self.viewer_combo.addItem("🖥️ 기본 뷰어 (폰트 안정)", "default")
        self.viewer_combo.addItem("📱 Edge PDF (권장)", "edge")
        self.viewer_combo.setCurrentIndex(2)  # Edge를 기본으로 설정
        self.viewer_combo.setMinimumWidth(200)
        viewer_layout.addWidget(self.viewer_combo)
        viewer_layout.addStretch()
        print_layout.addLayout(viewer_layout)
        
        # 인쇄 옵션
        options_layout = QHBoxLayout()
        
        options_layout.addWidget(QLabel("매수:"))
        self.copies_spin = QSpinBox()
        self.copies_spin.setMinimum(1)
        self.copies_spin.setMaximum(99)
        self.copies_spin.setValue(1)
        options_layout.addWidget(self.copies_spin)
        
        self.duplex_check = QCheckBox("양면 인쇄")
        options_layout.addWidget(self.duplex_check)
        
        options_layout.addStretch()
        print_layout.addLayout(options_layout)
        
        print_group.setLayout(print_layout)
        layout.addWidget(print_group)
        
        # 인쇄 버튼들
        print_buttons_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton("👀 미리보기")
        self.preview_btn.setMinimumHeight(35)
        self.preview_btn.setMaximumHeight(35)
        self.preview_btn.setEnabled(False)
        self.preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.preview_btn.clicked.connect(self.preview_order)
        print_buttons_layout.addWidget(self.preview_btn)
        
        self.print_btn = QPushButton("⚡ 빠른 인쇄")
        self.print_btn.setMinimumHeight(35)
        self.print_btn.setMaximumHeight(35)
        self.print_btn.setEnabled(False)
        self.print_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.print_btn.clicked.connect(self.print_order_direct)
        print_buttons_layout.addWidget(self.print_btn)
        
        layout.addLayout(print_buttons_layout)
        
        # 로그 영역
        log_group = QGroupBox("📝 작업 로그")
        log_layout = QVBoxLayout()
        
        self.search_log_text = QTextEdit()
        self.search_log_text.setReadOnly(True)
        self.search_log_text.setMaximumHeight(200)
        self.search_log_text.setStyleSheet("font-family: 'Consolas', monospace; font-size: 9pt;")
        log_layout.addWidget(self.search_log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # 탭에 추가
        self.tab_widget.addTab(tab_widget, "🔍 주문번호 검색")
        
        # 초기화
        self.order_searcher = OrderSearcher()
        self.print_manager = PrintManager()
        self.search_result = None
        self.search_worker = None  # 검색 작업 스레드
        
        # 진행률 표시
        self.search_progress = QProgressBar()
        self.search_progress.hide()
        log_layout.insertWidget(0, self.search_progress)
        
        # 저장된 뷰어 설정 로드
        self.load_viewer_settings()
        
        # 초기 로그
        self.search_log("주문번호 검색 기능이 준비되었습니다.")
        self.search_log("💡 폰트 문제시 PDF 뷰어를 'Edge PDF'로 변경하세요")
        self.refresh_printers()
    
    def load_saved_paths(self):
        """저장된 경로들을 로드하여 UI에 설정"""
        # 저장된 경로 가져오기
        excel_path = self.settings.value("excel_path", "")
        pdf_path = self.settings.value("pdf_path", "")
        output_path = self.settings.value("output_path", "")
        
        # UI에 설정 (파일이 실제로 존재하는 경우만)
        if excel_path and os.path.exists(excel_path):
            self.excel_edit.setText(excel_path)
            self.log(f"💾 저장된 엑셀 경로: {os.path.basename(excel_path)}")
        
        if pdf_path and os.path.exists(pdf_path):
            self.pdf_edit.setText(pdf_path)
            self.log(f"💾 저장된 PDF 경로: {os.path.basename(pdf_path)}")
        
        if output_path and os.path.exists(output_path):
            self.output_edit.setText(output_path)
            self.log(f"💾 저장된 출력 폴더: {output_path}")
    
    def save_path(self, key, path):
        """경로를 설정에 저장"""
        self.settings.setValue(key, path)
        self.settings.sync()  # 즉시 저장
    
    def on_fuzzy_changed(self, state):
        """유사도 매칭 체크박스 변경 이벤트"""
        self.threshold_spin.setEnabled(state == Qt.Checked)
    
    def browse_excel(self):
        """엑셀 파일 찾아보기 (작업 폴더 기준)"""
        # 작업 폴더에서 시작
        working_path = config.get_working_path()
        
        if not working_path:
            QMessageBox.warning(self, "경로 오류", "먼저 작업 폴더를 설정해주세요.")
            return
        
        # 현재 파일이 있으면 해당 디렉토리, 없으면 작업 폴더
        current_file = self.excel_edit.text()
        if current_file and os.path.exists(current_file):
            start_dir = os.path.dirname(current_file)
        else:
            start_dir = working_path
            
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", start_dir,
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.excel_edit.setText(file_path)
            self.save_path("excel_path", file_path)  # 경로 저장
            self.log(f"✓ 엑셀 선택: {os.path.basename(file_path)}")
    
    def browse_pdf(self):
        """PDF 파일 찾아보기 (작업 폴더 기준)"""
        # 작업 폴더에서 시작
        working_path = config.get_working_path()
        
        if not working_path:
            QMessageBox.warning(self, "경로 오류", "먼저 작업 폴더를 설정해주세요.")
            return
        
        # 현재 파일이 있으면 해당 디렉토리, 없으면 작업 폴더
        current_file = self.pdf_edit.text()
        if current_file and os.path.exists(current_file):
            start_dir = os.path.dirname(current_file)
        else:
            start_dir = working_path
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "PDF 파일 선택", start_dir,
            "PDF Files (*.pdf);;All Files (*)"
        )
        if file_path:
            self.pdf_edit.setText(file_path)
            self.save_path("pdf_path", file_path)  # 경로 저장
            self.log(f"✓ PDF 선택: {os.path.basename(file_path)}")
    
    # browse_output 메서드는 더 이상 사용하지 않음 (통합 경로 사용)
    
    def log(self, message):
        """로그 메시지 추가"""
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
    
    def run_processing(self):
        """처리 실행"""
        # 작업 폴더 확인
        working_path = config.get_working_path()
        if not working_path:
            QMessageBox.warning(self, "설정 오류", "작업 폴더가 설정되지 않았습니다.")
            return
        
        # 작업 폴더 유효성 재확인
        is_valid, message = config.validate_base_path(working_path)
        if not is_valid:
            QMessageBox.warning(self, "경로 오류", f"작업 폴더에 문제가 있습니다:\n{message}")
            return
        
        # 입력 검증
        excel_path = self.excel_edit.text().strip()
        pdf_path = self.pdf_edit.text().strip()
        output_dir = working_path  # 작업 폴더를 출력 디렉토리로 사용
        
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "입력 오류", "유효한 엑셀 파일을 선택하세요.")
            return
        
        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "입력 오류", "유효한 PDF 파일을 선택하세요.")
            return
        
        # 옵션
        use_fuzzy = self.fuzzy_check.isChecked()
        threshold = self.threshold_spin.value()
        
        # UI 비활성화
        self.run_btn.setEnabled(False)
        self.progress_bar.show()
        
        self.log("")
        self.log("=" * 60)
        self.log("🚀 PDF 정렬 작업을 시작합니다...")
        self.log("=" * 60)
        self.log(f"📊 엑셀: {os.path.basename(excel_path)}")
        self.log(f"📄 PDF: {os.path.basename(pdf_path)}")
        self.log(f"📁 출력: {output_dir}")
        
        # 작업 스레드 시작
        self.worker = ProcessingWorker(excel_path, pdf_path, output_dir, use_fuzzy, threshold)
        self.worker.progress.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
    
    def on_finished(self, result):
        """작업 완료"""
        self.progress_bar.hide()
        self.run_btn.setEnabled(True)
        
        # 완료 메시지
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("작업 완료")
        msg.setText("✅ PDF 정렬이 완료되었습니다!")
        msg.setInformativeText(
            f"엑셀 {result['total_excel_rows']}행 중 {result['matched_count']}건 매칭 성공\n"
            f"PDF {result['total_pdf_pages']}페이지 처리 완료"
        )
        msg.setDetailedText(
            f"결과 파일:\n"
            f"• PDF: {result['pdf_path']}\n"
            f"• 리포트: {result['csv_path']}\n\n"
            f"통계:\n"
            f"• 매칭 성공: {result['matched_count']}건\n"
            f"• 매칭 실패: {result['unmatched_count']}건\n"
            f"• 남은 페이지: {result['leftover_count']}개"
        )
        msg.exec()
    
    def on_error(self, error_msg):
        """작업 오류"""
        self.progress_bar.hide()
        self.run_btn.setEnabled(True)
        
        self.log("")
        self.log(error_msg)
        
        QMessageBox.critical(self, "오류", error_msg)
    
    # ===== 주문번호 검색 탭 관련 메서드들 =====
    
    # 검색 폴더 관련 메서드들은 더 이상 필요 없음 (통합 경로 사용)
    
    def search_order(self):
        """주문번호 검색 - 백그라운드 처리"""
        order_number = self.order_number_edit.text().strip()
        
        if not order_number:
            QMessageBox.warning(self, "입력 오류", "주문번호를 입력하세요.")
            return
        
        # 작업 폴더 확인
        working_path = config.get_working_path()
        if not working_path:
            QMessageBox.warning(self, "설정 오류", "작업 폴더가 설정되지 않았습니다.")
            return
        
        # 작업 폴더 유효성 확인
        is_valid, message = config.validate_base_path(working_path)
        if not is_valid:
            QMessageBox.warning(self, "경로 오류", f"작업 폴더에 문제가 있습니다:\n{message}")
            return
        
        # 이미 검색 중이면 중단
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.terminate()
            self.search_worker.wait()
        
        # UI 상태 변경: 검색 중 모드
        self.search_btn.setEnabled(False)
        self.stop_search_btn.setEnabled(True)
        self.print_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.search_progress.show()
        self.search_progress.setMinimum(0)
        self.search_progress.setMaximum(0)  # Indeterminate
        
        # 검색 시작 시간 기록
        self.search_start_time = time.time()
        
        # 백그라운드 검색 시작
        self.search_worker = OrderSearchWorker(order_number, working_path)
        self.search_worker.progress.connect(self.search_log)
        self.search_worker.finished.connect(self.on_search_finished)
        self.search_worker.error.connect(self.on_search_error)
        self.search_worker.start()
    
    def on_search_finished(self, search_result):
        """검색 완료 처리"""
        # UI 상태 복원: 검색 완료 모드
        self.search_btn.setEnabled(True)
        self.stop_search_btn.setEnabled(False)
        self.search_progress.hide()
        
        # 검색 시간 계산
        search_duration = int((time.time() - self.search_start_time) * 1000)
        
        if search_result:
            self.search_result = search_result
            self.display_search_result(search_result)
            self.print_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)
            self.search_log(f"⏱️ 검색 완료 ({search_duration}ms)")
            
            # 로그 기록
            logger.log_search_result(
                self.order_number_edit.text().strip(), 
                config.get_working_path(), 
                search_result, 
                search_duration
            )
        else:
            self.search_result = None
            self.clear_search_result()
            self.print_btn.setEnabled(False)
            self.preview_btn.setEnabled(False)
            self.search_log(f"⏱️ 검색 완료 ({search_duration}ms) - 결과 없음")
            
            # 비슷한 주문번호 추천
            self.suggest_similar_orders(self.order_number_edit.text().strip())
            
            # 로그 기록
            logger.log_search_result(
                self.order_number_edit.text().strip(), 
                config.get_working_path(), 
                None, 
                search_duration
            )
    
    def on_search_error(self, error_msg):
        """검색 오류 처리"""
        # UI 상태 복원: 검색 오류 모드
        self.search_btn.setEnabled(True)
        self.stop_search_btn.setEnabled(False)
        self.search_progress.hide()
        
        self.search_log(error_msg)
        QMessageBox.critical(self, "검색 오류", error_msg)
    
    def stop_search(self):
        """검색 중지"""
        if self.search_worker and self.search_worker.isRunning():
            self.search_log("⏹️ 검색 중지 요청...")
            
            # 스레드 종료
            self.search_worker.terminate()
            self.search_worker.wait(3000)  # 3초 대기
            
            # 강제 종료가 안되면 더 강력하게
            if self.search_worker.isRunning():
                self.search_worker.quit()
                self.search_worker.wait()
            
            # UI 상태 복원: 중지 완료 모드
            self.search_btn.setEnabled(True)
            self.stop_search_btn.setEnabled(False)
            self.search_progress.hide()
            
            # 검색 결과 클리어
            self.search_result = None
            self.clear_search_result()
            self.print_btn.setEnabled(False)
            self.preview_btn.setEnabled(False)
            
            self.search_log("✅ 검색이 중지되었습니다")
            
            # 검색 시간 계산 (중지된 경우)
            if hasattr(self, 'search_start_time'):
                elapsed = int((time.time() - self.search_start_time) * 1000)
                self.search_log(f"⏱️ 중지된 검색 시간: {elapsed}ms")
        else:
            self.search_log("⚠️ 진행 중인 검색이 없습니다")
    
    def display_search_result(self, search_result):
        """검색 결과를 테이블에 표시"""
        self.search_result_table.setRowCount(1)
        
        best_match = search_result.best_match
        
        # 파일명
        filename = os.path.basename(best_match.file_path)
        self.search_result_table.setItem(0, 0, QTableWidgetItem(filename))
        
        # 문서날짜
        doc_date = best_match.doc_date.strftime('%Y-%m-%d') if best_match.doc_date else ""
        self.search_result_table.setItem(0, 1, QTableWidgetItem(doc_date))
        
        # 파일명날짜
        filename_date = best_match.filename_date.strftime('%Y-%m-%d') if best_match.filename_date else ""
        self.search_result_table.setItem(0, 2, QTableWidgetItem(filename_date))
        
        # 수정시간
        modified_time = best_match.modified_time.strftime('%Y-%m-%d %H:%M')
        self.search_result_table.setItem(0, 3, QTableWidgetItem(modified_time))
        
        # 페이지 범위
        page_ranges = self.order_searcher.get_page_ranges_str(best_match.page_numbers)
        self.search_result_table.setItem(0, 4, QTableWidgetItem(page_ranges))
        
        # 선택 기준
        decided_by_text = {
            'doc_date': '문서날짜',
            'filename_date': '파일명날짜', 
            'modified_time': '수정시간'
        }.get(search_result.decided_by, search_result.decided_by)
        self.search_result_table.setItem(0, 5, QTableWidgetItem(decided_by_text))
        
        # 추가 정보 로그
        if len(search_result.all_matches) > 1:
            self.search_log(f"📊 총 {len(search_result.all_matches)}개 파일 중 최신 파일 선택 ({decided_by_text} 기준)")
    
    def clear_search_result(self):
        """검색 결과 테이블 클리어"""
        self.search_result_table.setRowCount(0)
    
    def preview_order(self):
        """주문번호 미리보기 - 다양한 뷰어 지원"""
        if not self.search_result:
            QMessageBox.warning(self, "미리보기 오류", "먼저 주문번호를 검색하세요.")
            return
        
        best_match = self.search_result.best_match
        page_ranges = self.order_searcher.get_page_ranges_str(best_match.page_numbers)
        
        # 선택된 뷰어 확인
        selected_viewer = self.viewer_combo.currentData()
        
        self.search_log(f"👀 미리보기 실행: {os.path.basename(best_match.file_path)} 페이지 {page_ranges}")
        self.search_log(f"📱 사용 뷰어: {self.viewer_combo.currentText()}")
        
        try:
            if selected_viewer == "sumatra":
                # SumatraPDF 사용
                if self.print_manager.is_sumatra_available():
                    import subprocess
                    subprocess.Popen([self.print_manager.sumatra_path, best_match.file_path])
                    self.search_log(f"✅ SumatraPDF로 미리보기 열림")
                else:
                    self.search_log("⚠️ SumatraPDF를 찾을 수 없어 기본 뷰어로 실행")
                    os.startfile(best_match.file_path)
                    
            elif selected_viewer == "edge":
                # Microsoft Edge로 실행 (폰트 지원 우수)
                import subprocess
                try:
                    subprocess.Popen([
                        "msedge.exe", 
                        best_match.file_path,
                        "--new-window"
                    ])
                    self.search_log(f"✅ Edge로 미리보기 열림 (폰트 안정)")
                except:
                    # Edge 실행 실패시 기본 뷰어로
                    self.search_log("⚠️ Edge 실행 실패, 기본 뷰어로 실행")
                    os.startfile(best_match.file_path)
                    
            else:  # default
                # 기본 PDF 뷰어 사용
                os.startfile(best_match.file_path)
                self.search_log(f"✅ 기본 뷰어로 미리보기 열림")
            
            # 안내 메시지 (폰트 문제 해결 팁 포함)
            QMessageBox.information(self, "미리보기 열림", 
                f"PDF 미리보기가 열렸습니다.\n\n"
                f"📄 파일: {os.path.basename(best_match.file_path)}\n"
                f"📄 해당 페이지: {page_ranges}\n"
                f"📱 뷰어: {self.viewer_combo.currentText()}\n\n"
                f"💡 폰트가 깨져 보이면:\n"
                f"   1. PDF 뷰어를 'Edge PDF'로 변경\n"
                f"   2. 또는 '기본 뷰어'로 변경 후 재시도\n\n"
                f"확인 후 뷰어에서 직접 인쇄하거나\n"
                f"'빠른 인쇄' 버튼을 사용하세요.")
                
        except Exception as e:
            error_msg = f"미리보기 실행 중 오류: {str(e)}"
            self.search_log(f"❌ {error_msg}")
            QMessageBox.critical(self, "미리보기 오류", error_msg)
    
    def print_order_direct(self):
        """주문번호 직접 인쇄 (설정 기반)"""
        if not self.search_result:
            QMessageBox.warning(self, "인쇄 오류", "먼저 주문번호를 검색하세요.")
            return
        
        # 인쇄 설정 확인
        printer_name = self.printer_combo.currentText()
        if not printer_name:
            QMessageBox.warning(self, "인쇄 오류", "프린터를 선택하세요.")
            return
        
        # 선택된 뷰어에 따른 인쇄 방식 결정
        selected_viewer = self.viewer_combo.currentData()
        
        if selected_viewer == "sumatra" and not self.print_manager.is_sumatra_available():
            QMessageBox.warning(self, "인쇄 오류", 
                f"SumatraPDF를 찾을 수 없습니다.\n\n"
                f"다음 중 하나를 시도해보세요:\n"
                f"1. PDF 뷰어를 'Edge PDF' 또는 '기본 뷰어'로 변경\n"
                f"2. SumatraPDF 설치 후 재시도\n\n"
                f"현재 SumatraPDF 경로: {self.print_manager.sumatra_path or '없음'}")
            return
        
        # 인쇄 확인 대화상자
        best_match = self.search_result.best_match
        page_ranges = self.order_searcher.get_page_ranges_str(best_match.page_numbers)
        copies = self.copies_spin.value()
        duplex = self.duplex_check.isChecked()
        
        # 디버깅 정보
        self.search_log(f"🔍 디버그: 페이지 번호 리스트 = {best_match.page_numbers}")
        self.search_log(f"🔍 디버그: 페이지 범위 문자열 = '{page_ranges}'")
        
        reply = QMessageBox.question(
            self, "인쇄 확인", 
            f"다음 내용으로 인쇄하시겠습니까?\n\n"
            f"📄 파일: {os.path.basename(best_match.file_path)}\n"
            f"📄 페이지: {page_ranges}\n"
            f"🖨️ 프린터: {printer_name}\n"
            f"📰 매수: {copies}매\n"
            f"📋 양면: {'예' if duplex else '아니오'}\n\n"
            f"🔍 디버그: 실제 페이지 번호 = {best_match.page_numbers}",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # 선택된 뷰어에 따라 다른 방식으로 인쇄
        if selected_viewer == "sumatra":
            self.print_order_execute(best_match.file_path, page_ranges, printer_name, copies, duplex)
        else:
            # Edge나 기본 뷰어는 인쇄 대화상자 방식 사용
            self.search_log(f"📱 {self.viewer_combo.currentText()}에서 인쇄 대화상자 실행")
            try:
                if selected_viewer == "edge":
                    # Edge로 파일 열고 인쇄 가이드
                    import subprocess
                    subprocess.Popen([
                        "msedge.exe", 
                        best_match.file_path,
                        "--new-window"
                    ])
                    
                    QMessageBox.information(self, "인쇄 안내", 
                        f"Edge에서 PDF가 열렸습니다.\n\n"
                        f"📄 해당 페이지: {page_ranges}\n\n"
                        f"인쇄 방법:\n"
                        f"1. Edge에서 Ctrl+P 누르기\n"
                        f"2. 페이지 범위에 '{page_ranges}' 입력\n"
                        f"3. 매수: {copies}매 설정\n"
                        f"4. 인쇄 실행")
                        
                else:
                    # 기본 뷰어로 열고 안내
                    os.startfile(best_match.file_path)
                    
                    QMessageBox.information(self, "인쇄 안내",
                        f"기본 PDF 뷰어로 파일이 열렸습니다.\n\n"
                        f"📄 해당 페이지: {page_ranges}\n\n"
                        f"인쇄 방법:\n"
                        f"1. PDF 뷰어에서 인쇄 (Ctrl+P)\n"
                        f"2. 페이지 범위에 '{page_ranges}' 입력\n"
                        f"3. 매수: {copies}매 설정\n"
                        f"4. 인쇄 실행")
                
            except Exception as e:
                self.search_log(f"❌ 뷰어 실행 오류: {str(e)}")
                QMessageBox.critical(self, "오류", f"뷰어 실행 중 오류: {str(e)}")
    
    def print_order(self):
        """주문번호 인쇄 (미리보기 후 인쇄)"""
        if not self.search_result:
            QMessageBox.warning(self, "인쇄 오류", "먼저 주문번호를 검색하세요.")
            return
        
        best_match = self.search_result.best_match
        
        self.search_log(f"🖨️ 인쇄 대화상자 실행: {os.path.basename(best_match.file_path)}")
        
        try:
            # SumatraPDF 인쇄 대화상자로 실행
            if self.print_manager.is_sumatra_available():
                success = self.print_manager.print_dialog(best_match.file_path)
                if success:
                    self.search_log(f"✅ 인쇄 대화상자 열림")
                    
                    # 로그 기록 (대화상자 형태)
                    logger.log_print_result(
                        order_number=self.search_result.order_number,
                        file_path=best_match.file_path,
                        page_ranges="dialog",
                        printer_name="user_selected",
                        copies=0,
                        duplex=False,
                        success=True
                    )
                else:
                    self.search_log(f"❌ 인쇄 대화상자 실행 실패")
            else:
                QMessageBox.warning(self, "인쇄 오류", 
                    "SumatraPDF를 찾을 수 없습니다.\n"
                    "미리보기 기능을 사용하여 기본 뷰어에서 인쇄하세요.")
                    
        except Exception as e:
            error_msg = f"인쇄 대화상자 실행 중 오류: {str(e)}"
            self.search_log(f"❌ {error_msg}")
            QMessageBox.critical(self, "인쇄 오류", error_msg)
    
    def print_order_execute(self, file_path, page_ranges, printer_name, copies, duplex):
        """실제 인쇄 실행"""
        self.print_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        
        self.search_log(f"🖨️ 인쇄 실행 중...")
        
        try:
            start_time = time.time()
            
            # 인쇄 실행
            success = self.print_manager.print_pages(
                pdf_path=file_path,
                page_ranges=page_ranges,
                printer_name=printer_name,
                copies=copies,
                duplex=duplex
            )
            
            print_duration = int((time.time() - start_time) * 1000)
            
            if success:
                self.search_log(f"✅ 인쇄 완료: {printer_name}에서 {copies}매 출력 ({print_duration}ms)")
                QMessageBox.information(self, "인쇄 완료", 
                    f"인쇄가 완료되었습니다!\n\n"
                    f"⏱️ 소요 시간: {print_duration}ms")
            else:
                self.search_log(f"❌ 인쇄 실패 ({print_duration}ms)")
                QMessageBox.critical(self, "인쇄 실패", 
                    "인쇄 중 오류가 발생했습니다.\n"
                    "프린터 상태를 확인해주세요.")
            
            # 로그 기록
            logger.log_print_result(
                order_number=self.search_result.order_number,
                file_path=file_path,
                page_ranges=page_ranges,
                printer_name=printer_name,
                copies=copies,
                duplex=duplex,
                success=success,
                print_duration_ms=print_duration
            )
            
        except Exception as e:
            error_msg = str(e)
            self.search_log(f"❌ 인쇄 중 오류: {error_msg}")
            QMessageBox.critical(self, "인쇄 오류", f"인쇄 중 오류가 발생했습니다:\n{error_msg}")
            
            # 오류 로그 기록
            logger.log_print_result(
                order_number=self.search_result.order_number,
                file_path=file_path,
                page_ranges=page_ranges,
                printer_name=printer_name,
                copies=copies,
                duplex=duplex,
                success=False,
                error_message=error_msg
            )
            
        finally:
            self.print_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)
    
    def refresh_printers(self):
        """프린터 목록 새로고침"""
        try:
            self.printer_combo.clear()
            
            # 사용 가능한 프린터 목록 가져오기
            printers = self.print_manager.get_available_printers()
            
            if printers:
                self.printer_combo.addItems(printers)
                
                # 기본 프린터 선택
                default_printer = self.print_manager.get_default_printer()
                if default_printer and default_printer in printers:
                    self.printer_combo.setCurrentText(default_printer)
                
                self.search_log(f"✓ {len(printers)}개 프린터 발견")
            else:
                self.search_log("⚠️ 사용 가능한 프린터가 없습니다")
                
        except Exception as e:
            self.search_log(f"❌ 프린터 목록 가져오기 실패: {str(e)}")
    
    def search_log(self, message: str):
        """검색 로그 메시지 추가"""
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.search_log_text.append(f"[{timestamp}] {message}")
        
        # 스크롤을 맨 아래로
        scrollbar = self.search_log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def load_viewer_settings(self):
        """저장된 뷰어 설정 로드"""
        try:
            default_viewer = config.get("print_settings.default_viewer", "edge")
            
            # 콤보박스에서 해당 뷰어 선택
            for i in range(self.viewer_combo.count()):
                if self.viewer_combo.itemData(i) == default_viewer:
                    self.viewer_combo.setCurrentIndex(i)
                    break
                    
            # 뷰어 변경 이벤트 연결
            self.viewer_combo.currentTextChanged.connect(self.on_viewer_changed)
            
        except Exception as e:
            self.search_log(f"뷰어 설정 로드 오류: {str(e)}")
    
    def on_viewer_changed(self):
        """PDF 뷰어 변경시 설정 저장"""
        try:
            selected_viewer = self.viewer_combo.currentData()
            config.set("print_settings.default_viewer", selected_viewer)
            
            viewer_name = self.viewer_combo.currentText()
            self.search_log(f"📱 PDF 뷰어 변경: {viewer_name}")
            
            # 뷰어별 안내 메시지
            if selected_viewer == "edge":
                self.search_log("✅ Edge PDF: 한글 폰트 지원 우수, 권장")
            elif selected_viewer == "sumatra":
                self.search_log("⚡ SumatraPDF: 빠른 실행, 폰트 문제 있을 수 있음")
            else:
                self.search_log("🖥️ 기본 뷰어: 시스템 설정 PDF 프로그램 사용")
                
        except Exception as e:
            self.search_log(f"뷰어 설정 저장 오류: {str(e)}")
    
    def suggest_similar_orders(self, search_order):
        """비슷한 주문번호 추천"""
        try:
            from matcher import normalize_order_number, extract_order_numbers_from_text
            normalized_search = normalize_order_number(search_order)
            if not normalized_search or len(normalized_search) < 6:
                return
            
            # 같은 시리즈 찾기 (뒷자리 기준)
            if len(normalized_search) >= 8:
                series_prefix = normalized_search[-8:]  # 뒷자리 8자리로 시리즈 판단
            else:
                series_prefix = normalized_search
            working_path = config.get_working_path()
            
            if not working_path:
                return
                
            # 빠른 스캔으로 비슷한 주문번호 찾기
            similar_orders = set()
            
            import pdfplumber
            from pathlib import Path
            
            pdf_files = list(Path(working_path).glob("*.pdf"))
            
            for pdf_file in pdf_files[:3]:  # 최대 3개 파일만 확인
                try:
                    with pdfplumber.open(str(pdf_file)) as pdf:
                        for i, page in enumerate(pdf.pages[:20]):  # 각 파일당 최대 20페이지
                            text = page.extract_text() or ''
                            
                            extracted = extract_order_numbers_from_text(text)
                            for order_raw in extracted:
                                normalized = normalize_order_number(order_raw)
                                # 같은 시리즈이고 유효한 길이인 경우
                                if (len(normalized) >= 8 and 
                                    normalized.endswith(series_prefix) and 
                                    normalized != normalized_search):
                                    similar_orders.add(normalized)
                            
                            if len(similar_orders) >= 5:  # 5개 찾으면 중단
                                break
                        
                        if len(similar_orders) >= 5:
                            break
                            
                except Exception:
                    continue
            
            if similar_orders:
                similar_list = sorted(list(similar_orders))[:5]
                self.search_log(f"💡 비슷한 주문번호 추천:")
                for order in similar_list:
                    # 10자리 주문번호를 사용자 친화적 형태로 표시
                    if len(order) == 10:
                        # 10자리면 앞에 적절한 접두사 추가
                        display_form = f"예상 형태: ****{order}"
                    else:
                        display_form = order
                    self.search_log(f"   • {display_form}")
                    
        except Exception as e:
            self.search_log(f"추천 검색 중 오류: {str(e)}")


def main():
    app = QApplication(sys.argv)
    
    # 스타일
    app.setStyle('Fusion')
    
    # 윈도우 생성 및 표시
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
