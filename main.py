"""
PDF 프로그램 - 주문번호 기반 PDF 자동 정렬 및 검색/인쇄
"""

import sys
import os
import traceback
from pathlib import Path
import tempfile
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit,
    QProgressBar, QGroupBox, QMessageBox, QCheckBox, QSpinBox,
    QTabWidget, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QFont

from io_utils import load_excel, get_output_filenames, is_text_based_pdf
from matcher import extract_pages, match_rows_to_pages, reorder_pdf
from search_print import search_order_in_pdf, search_order_in_folder, extract_pages_to_pdf, open_pdf_for_print
from pdf_numbering import add_page_numbers_by_order


class ProcessingWorker(QThread):
    """PDF 정렬 작업 스레드"""
    progress = Signal(str)
    finished = Signal(object)
    error = Signal(str)
    
    def __init__(self, excel_path, pdf_path, output_dir, use_fuzzy, threshold, only_matched=False):
        super().__init__()
        self.excel_path = excel_path
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        self.use_fuzzy = use_fuzzy
        self.threshold = threshold
        self.only_matched = only_matched
    
    def run(self):
        try:
            self.progress.emit("📄 PDF 파일 확인 중...")
            is_text, msg = is_text_based_pdf(self.pdf_path)
            if not is_text:
                self.error.emit(msg)
                return
            self.progress.emit(f"✓ {msg}")
            
            self.progress.emit("📊 엑셀 파일 로드 중...")
            df = load_excel(self.excel_path)
            self.progress.emit(f"✓ 엑셀 로드 완료 ({len(df)}행)")
            
            self.progress.emit("📖 PDF 텍스트 추출 중...")
            pages = extract_pages(self.pdf_path)
            self.progress.emit(f"✓ PDF 텍스트 추출 완료 ({len(pages)}페이지)")
            
            self.progress.emit("🔍 주문번호 매칭 중...")
            assignments, leftover_pages, match_details = match_rows_to_pages(
                df, pages, self.use_fuzzy, self.threshold
            )
            
            matched_count = len(assignments)
            unmatched_count = len(df) - matched_count
            self.progress.emit(f"✓ 매칭 완료: {matched_count}건 성공, {unmatched_count}건 실패")
            
            self.progress.emit("📑 페이지 순서 결정 중...")
            ordered_indices = []
            page_to_order = {}  # {결과_PDF_페이지_인덱스: 주문_번호}
            
            # 엑셀에서 고유한 주문번호에 순차적 번호 부여
            order_number_to_display_num = {}  # {주문번호: 표시할_번호}
            display_num = 1
            
            for row_idx in range(len(df)):
                row = df.iloc[row_idx]
                order_number = str(row['주문번호']).strip()
                
                # 이미 본 주문번호가 아니면 새 번호 부여
                if order_number not in order_number_to_display_num:
                    order_number_to_display_num[order_number] = display_num
                    display_num += 1
            
            result_page_idx = 0
            
            # 매칭된 페이지 (엑셀 순서대로)
            for row_idx in range(len(df)):
                if row_idx in assignments:
                    page_idx = assignments[row_idx]
                    ordered_indices.append(page_idx)
            
                    # 해당 엑셀 행의 주문번호로 표시 번호 결정
                    row = df.iloc[row_idx]
                    order_number = str(row['주문번호']).strip()
                    page_to_order[result_page_idx] = order_number_to_display_num[order_number]
                    result_page_idx += 1
            
            # 미매칭 페이지 처리 (옵션에 따라)
            if not self.only_matched:
                # 미매칭 페이지도 번호 없이 마지막에 추가
                for page_idx in leftover_pages:
                    ordered_indices.append(page_idx)
                    # page_to_order에 추가하지 않음 = 넘버링 없음
                    result_page_idx += 1
            else:
                self.progress.emit(f"ℹ️ 미매칭 페이지 {len(leftover_pages)}개 제외됨")
            
            self.progress.emit("💾 PDF 저장 중...")
            pdf_out_path, csv_out_path = get_output_filenames(self.output_dir)
            
            # 임시 파일로 먼저 정렬
            temp_pdf = pdf_out_path + ".temp"
            reorder_pdf(self.pdf_path, ordered_indices, temp_pdf)
            
            # 페이지 번호 추가 (주문번호 기준)
            self.progress.emit("🔢 페이지 번호 추가 중 (주문번호 기준)...")
            add_page_numbers_by_order(temp_pdf, pdf_out_path, page_to_order, font_size=5)
            
            # 임시 파일 삭제
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            
            self.progress.emit("📝 리포트 생성 중...")
            report_rows = []
            for row_idx in range(len(df)):
                row = df.iloc[row_idx]
                detail = match_details.get(row_idx, {'page_idx': -1, 'score': 0, 'reason': 'no_match'})
                report_row = {
                    '엑셀행번호': row_idx + 2,
                    '매칭페이지': detail['page_idx'] + 1 if detail['page_idx'] >= 0 else 'UNMATCHED',
                    '점수': round(detail['score'], 1),
                    '매칭키': detail['reason'],
                    '주문번호': row['주문번호']
                }
                report_rows.append(report_row)
            
            from io_utils import save_report
            save_report(report_rows, csv_out_path)
            
            self.progress.emit("")
            self.progress.emit("✅ 완료!")
            self.progress.emit(f"📂 PDF: {os.path.basename(pdf_out_path)}")
            self.progress.emit(f"📊 리포트: {os.path.basename(csv_out_path)}")
            
            result = {
                'pdf_path': pdf_out_path,
                'csv_path': csv_out_path,
                'matched': matched_count,
                'unmatched': unmatched_count
            }
            self.finished.emit(result)
            
        except Exception as e:
            self.error.emit(f"❌ 오류:\n{str(e)}\n\n{traceback.format_exc()}")


class SearchWorker(QThread):
    """주문번호 검색 작업 스레드"""
    progress = Signal(str)
    finished = Signal(object)
    error = Signal(str)
    
    def __init__(self, search_path, order_number, is_folder, save_folder=None, use_multiprocess=True, find_all=False):
        super().__init__()
        self.search_path = search_path
        self.order_number = order_number
        self.is_folder = is_folder
        self.save_folder = save_folder
        self.use_multiprocess = use_multiprocess
        self.find_all = find_all
        self._stop_flag = False
    
    def stop(self):
        """검색 중지"""
        self._stop_flag = True
        self.progress.emit("⏸️ 검색 중지 요청됨...")
    
    def run(self):
        try:
            if self.is_folder:
                self.progress.emit(f"📁 폴더 검색 시작...")
                
                # 진행 상황 콜백과 중지 플래그 전달
                def progress_cb(msg):
                    self.progress.emit(msg)
                
                def stop_check():
                    return self._stop_flag
                
                results = search_order_in_folder(
                    self.search_path, 
                    self.order_number,
                    progress_callback=progress_cb,
                    stop_flag=stop_check,
                    use_multiprocess=self.use_multiprocess,
                    find_all=self.find_all
                )
                
                if self._stop_flag:
                    self.progress.emit("⏹️ 검색이 중지되었습니다.")
                    self.finished.emit(None)
                    return
                
                if not results:
                    self.progress.emit(f"❌ 주문번호 '{self.order_number}'를 찾을 수 없습니다.")
                    self.finished.emit(None)
                    return
                
                # 결과가 여러 개인 경우 (이미 최신순 정렬됨)
                if len(results) > 1:
                    self.progress.emit(f"📋 {len(results)}개 파일에서 발견됨")
                    for i, (path, _, _) in enumerate(results[:3], 1):  # 상위 3개만 표시
                        self.progress.emit(f"   {i}. {os.path.basename(path)}")
                    if len(results) > 3:
                        self.progress.emit(f"   ... 외 {len(results)-3}개")
                
                # 가장 최신 파일 선택 (첫 번째)
                pdf_path, pages, modified_time = results[0]
                mod_date = datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')
                self.progress.emit(f"✅ 최신 파일 선택: {os.path.basename(pdf_path)}")
                self.progress.emit(f"   수정 시간: {mod_date}")
                
            else:
                self.progress.emit(f"📄 PDF 검색 중...")
                pages = search_order_in_pdf(self.search_path, self.order_number)
                
                if not pages:
                    self.progress.emit(f"❌ 주문번호 '{self.order_number}'를 찾을 수 없습니다.")
                    self.finished.emit(None)
                    return
                
                pdf_path = self.search_path
                self.progress.emit(f"✅ 발견!")
            
            self.progress.emit(f"📄 페이지: {', '.join(map(str, pages))}")
            
            # 저장 위치 결정
            if self.save_folder and os.path.exists(self.save_folder):
                save_dir = self.save_folder
            else:
                save_dir = tempfile.gettempdir()
            
            saved_files = []
            
            # 모든 페이지를 하나의 PDF로 저장 (단일 파일)
            self.progress.emit("💾 PDF 생성 중...")
            pdf_name = f"order_{self.order_number}.pdf"
            output_pdf = os.path.join(save_dir, pdf_name)
            extract_pages_to_pdf(pdf_path, pages, output_pdf)
            saved_files.append(output_pdf)
            self.progress.emit(f"✅ 저장 완료: {pdf_name}")
            
            result = {
                'pdf_path': pdf_path,
                'pages': pages,
                'saved_files': saved_files,
                'save_folder': save_dir
            }
            self.finished.emit(result)
                
        except Exception as e:
            self.error.emit(f"❌ 오류:\n{str(e)}\n\n{traceback.format_exc()}")


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 프로그램 - 정렬 및 검색/인쇄")
        self.setGeometry(100, 100, 1000, 800)
        
        self.settings = QSettings("PDFMatcher", "Settings")
        
        # 메인 위젯
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # 제목
        title = QLabel("📊 PDF 프로그램")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # 구분선
        line = QLabel()
        line.setFrameStyle(QLabel.HLine | QLabel.Sunken)
        main_layout.addWidget(line)
        
        # 탭 위젯
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 탭 생성
        self.create_sort_tab()
        self.create_search_tab()
        
        # 작업 스레드
        self.sort_worker = None
        self.search_worker = None
    
    def create_sort_tab(self):
        """PDF 정렬 탭"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 파일 선택
        file_group = QGroupBox("📁 파일 선택")
        file_layout = QVBoxLayout()
        
        # 엑셀
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("엑셀:"))
        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText("주문번호 컬럼이 포함된 엑셀 파일")
        self.excel_edit.setText(self.settings.value("sort/excel_path", ""))
        excel_layout.addWidget(self.excel_edit)
        excel_btn = QPushButton("선택")
        excel_btn.clicked.connect(self.select_excel)
        excel_layout.addWidget(excel_btn)
        file_layout.addLayout(excel_layout)
        
        # PDF
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(QLabel("PDF:"))
        self.pdf_edit = QLineEdit()
        self.pdf_edit.setPlaceholderText("정렬할 PDF 파일")
        self.pdf_edit.setText(self.settings.value("sort/pdf_path", ""))
        pdf_layout.addWidget(self.pdf_edit)
        pdf_btn = QPushButton("선택")
        pdf_btn.clicked.connect(self.select_pdf)
        pdf_layout.addWidget(pdf_btn)
        file_layout.addLayout(pdf_layout)
        
        # 출력 폴더
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("출력:"))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("결과 파일을 저장할 폴더")
        self.output_edit.setText(self.settings.value("sort/output_path", ""))
        output_layout.addWidget(self.output_edit)
        output_btn = QPushButton("선택")
        output_btn.clicked.connect(self.select_output)
        output_layout.addWidget(output_btn)
        file_layout.addLayout(output_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 옵션
        option_group = QGroupBox("⚙️ 옵션")
        option_layout = QHBoxLayout()
        
        self.fuzzy_check = QCheckBox("유사도 매칭")
        self.fuzzy_check.setToolTip("정확히 일치하지 않아도 유사한 경우 매칭 (권장하지 않음)")
        self.fuzzy_check.setChecked(self.settings.value("sort/use_fuzzy", False, type=bool))
        option_layout.addWidget(self.fuzzy_check)
        
        self.only_matched_check = QCheckBox("매칭된 페이지만 포함")
        self.only_matched_check.setToolTip("체크: 넘버링된 페이지만 PDF 생성\n체크 해제: 미매칭 페이지도 함께 포함 (번호 없음)")
        self.only_matched_check.setChecked(self.settings.value("sort/only_matched", False, type=bool))
        option_layout.addWidget(self.only_matched_check)
        
        option_layout.addStretch()
        option_group.setLayout(option_layout)
        layout.addWidget(option_group)
        
        # 실행 버튼
        self.sort_btn = QPushButton("🚀 PDF 정렬 시작")
        self.sort_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.sort_btn.clicked.connect(self.start_sort)
        layout.addWidget(self.sort_btn)
        
        # 진행 상황
        progress_group = QGroupBox("📊 진행 상황")
        progress_layout = QVBoxLayout()
        self.sort_log = QTextEdit()
        self.sort_log.setReadOnly(True)
        self.sort_log.setMinimumHeight(200)
        progress_layout.addWidget(self.sort_log)
        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)
        
        # 결과 버튼
        result_layout = QHBoxLayout()
        self.open_result_pdf_btn = QPushButton("📄 결과 PDF 열기")
        self.open_result_pdf_btn.setEnabled(False)
        self.open_result_pdf_btn.clicked.connect(self.open_result_pdf)
        result_layout.addWidget(self.open_result_pdf_btn)
        
        self.open_result_csv_btn = QPushButton("📊 리포트 열기")
        self.open_result_csv_btn.setEnabled(False)
        self.open_result_csv_btn.clicked.connect(self.open_result_csv)
        result_layout.addWidget(self.open_result_csv_btn)
        
        result_layout.addStretch()
        layout.addLayout(result_layout)
        
        self.tab_widget.addTab(tab, "📄 PDF 정렬")
        
        # 결과 경로 저장용
        self.result_pdf_path = None
        self.result_csv_path = None
    
    def create_search_tab(self):
        """주문번호 검색/인쇄 탭"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 검색 대상 선택
        target_group = QGroupBox("🔍 검색 대상")
        target_layout = QVBoxLayout()
        
        # 라디오 버튼
        self.search_type_group = QButtonGroup()
        self.radio_file = QRadioButton("PDF 파일")
        self.radio_folder = QRadioButton("폴더")
        
        # 마지막 선택 복원
        is_folder = self.settings.value("search/is_folder", False, type=bool)
        if is_folder:
            self.radio_folder.setChecked(True)
        else:
            self.radio_file.setChecked(True)
        
        self.search_type_group.addButton(self.radio_file)
        self.search_type_group.addButton(self.radio_folder)
        
        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.radio_file)
        radio_layout.addWidget(self.radio_folder)
        radio_layout.addStretch()
        target_layout.addLayout(radio_layout)
        
        # 경로 선택
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("경로:"))
        self.search_path_edit = QLineEdit()
        self.search_path_edit.setPlaceholderText("검색할 PDF 파일 또는 폴더 선택")
        self.search_path_edit.setText(self.settings.value("search/path", ""))
        path_layout.addWidget(self.search_path_edit)
        
        self.search_path_btn = QPushButton("선택")
        self.search_path_btn.clicked.connect(self.select_search_path)
        path_layout.addWidget(self.search_path_btn)
        target_layout.addLayout(path_layout)
        
        target_group.setLayout(target_layout)
        layout.addWidget(target_group)
        
        # 주문번호 및 저장 위치
        order_group = QGroupBox("📝 주문번호 및 저장 위치")
        order_layout = QVBoxLayout()
        
        # 주문번호
        order_num_layout = QHBoxLayout()
        order_num_layout.addWidget(QLabel("주문번호:"))
        self.order_number_edit = QLineEdit()
        self.order_number_edit.setPlaceholderText("검색할 주문번호 입력")
        order_num_layout.addWidget(self.order_number_edit)
        order_layout.addLayout(order_num_layout)
        
        # 저장 위치
        save_layout = QHBoxLayout()
        save_layout.addWidget(QLabel("저장 위치:"))
        self.save_path_edit = QLineEdit()
        self.save_path_edit.setPlaceholderText("추출한 PDF를 저장할 폴더 선택 (선택사항)")
        self.save_path_edit.setText(self.settings.value("search/save_path", ""))
        save_layout.addWidget(self.save_path_edit)
        save_btn = QPushButton("선택")
        save_btn.clicked.connect(self.select_save_folder)
        save_layout.addWidget(save_btn)
        order_layout.addLayout(save_layout)
        
        order_group.setLayout(order_layout)
        layout.addWidget(order_group)
        
        # 검색 버튼 영역
        search_btn_layout = QHBoxLayout()
        
        self.search_btn = QPushButton("🔍 검색 시작")
        self.search_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #0b7dda; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.search_btn.clicked.connect(self.start_search)
        search_btn_layout.addWidget(self.search_btn)
        
        self.stop_search_btn = QPushButton("⏹️ 검색 중지")
        self.stop_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #da190b; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.stop_search_btn.setEnabled(False)
        self.stop_search_btn.clicked.connect(self.stop_search)
        search_btn_layout.addWidget(self.stop_search_btn)
        
        # 멀티프로세싱 토글
        self.use_mp_check = QCheckBox("멀티프로세싱 사용")
        self.use_mp_check.setToolTip("PDF가 10개 이상일 때 병렬 검색 사용 (빠르지만 시스템에 따라 불안정할 수 있음)")
        self.use_mp_check.setChecked(self.settings.value("search/use_multiprocess", True, type=bool))
        search_btn_layout.addWidget(self.use_mp_check)
        
        # 전체 검색 토글
        self.find_all_check = QCheckBox("전체 검색 후 최신 파일 찾기")
        self.find_all_check.setToolTip("체크: 모든 파일 검색 후 최신 파일 선택\n체크 해제: 첫 번째 매칭 파일에서 검색 중단 (빠름)")
        self.find_all_check.setChecked(self.settings.value("search/find_all", False, type=bool))
        search_btn_layout.addWidget(self.find_all_check)

        layout.addLayout(search_btn_layout)
        
        # 검색 결과
        result_group = QGroupBox("📊 검색 결과")
        result_layout = QVBoxLayout()
        self.search_log = QTextEdit()
        self.search_log.setReadOnly(True)
        self.search_log.setMinimumHeight(200)
        result_layout.addWidget(self.search_log)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 인쇄/열기 버튼
        action_layout = QHBoxLayout()
        
        self.print_btn = QPushButton("🖨️ PDF 열기 (인쇄)")
        self.print_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #e68900; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.print_btn.setEnabled(False)
        self.print_btn.clicked.connect(self.open_for_print)
        action_layout.addWidget(self.print_btn)
        
        self.open_folder_btn_search = QPushButton("📁 저장 폴더 열기")
        self.open_folder_btn_search.setEnabled(False)
        self.open_folder_btn_search.clicked.connect(self.open_save_folder)
        action_layout.addWidget(self.open_folder_btn_search)
        
        layout.addLayout(action_layout)
        
        self.tab_widget.addTab(tab, "🔍 검색/인쇄")
        
        # 검색 결과 저장용
        self.temp_pdf_paths = []  # 여러 파일 경로 저장
        self.current_save_folder = None
    
    # === PDF 정렬 탭 메서드 ===
    
    def select_excel(self):
        last_dir = os.path.dirname(self.excel_edit.text()) or ""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", last_dir, "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_edit.setText(file_path)
            self.settings.setValue("sort/excel_path", file_path)
    
    def select_pdf(self):
        last_dir = os.path.dirname(self.pdf_edit.text()) or ""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "PDF 파일 선택", last_dir, "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_edit.setText(file_path)
            self.settings.setValue("sort/pdf_path", file_path)
    
    def select_output(self):
        last_dir = self.output_edit.text() or ""
        folder_path = QFileDialog.getExistingDirectory(self, "출력 폴더 선택", last_dir)
        if folder_path:
            self.output_edit.setText(folder_path)
            self.settings.setValue("sort/output_path", folder_path)
    
    def start_sort(self):
        excel_path = self.excel_edit.text().strip()
        pdf_path = self.pdf_edit.text().strip()
        output_dir = self.output_edit.text().strip()
        
        if not excel_path or not pdf_path or not output_dir:
            QMessageBox.warning(self, "입력 오류", "모든 파일/폴더를 선택하세요.")
            return
        
        if not os.path.exists(excel_path) or not os.path.exists(pdf_path) or not os.path.exists(output_dir):
            QMessageBox.warning(self, "파일 오류", "선택한 파일/폴더가 존재하지 않습니다.")
            return
        
        self.sort_log.clear()
        self.sort_btn.setEnabled(False)
        self.open_result_pdf_btn.setEnabled(False)
        self.open_result_csv_btn.setEnabled(False)
        
        use_fuzzy = self.fuzzy_check.isChecked()
        only_matched = self.only_matched_check.isChecked()
        
        # 설정 저장
        self.settings.setValue("sort/use_fuzzy", use_fuzzy)
        self.settings.setValue("sort/only_matched", only_matched)
        
        self.sort_worker = ProcessingWorker(excel_path, pdf_path, output_dir, use_fuzzy, 98, only_matched)
        self.sort_worker.progress.connect(self.update_sort_log)
        self.sort_worker.finished.connect(self.sort_finished)
        self.sort_worker.error.connect(self.sort_error)
        self.sort_worker.start()
    
    def update_sort_log(self, message):
        self.sort_log.append(message)
        self.sort_log.verticalScrollBar().setValue(
            self.sort_log.verticalScrollBar().maximum()
        )
    
    def sort_finished(self, result):
        self.sort_btn.setEnabled(True)
        if result:
            self.result_pdf_path = result['pdf_path']
            self.result_csv_path = result['csv_path']
            self.open_result_pdf_btn.setEnabled(True)
            self.open_result_csv_btn.setEnabled(True)
            
            QMessageBox.information(
                self, "완료",
                f"PDF 정렬 완료!\n\n"
                f"매칭 성공: {result['matched']}건\n"
                f"매칭 실패: {result['unmatched']}건"
            )
    
    def sort_error(self, error_msg):
        self.sort_btn.setEnabled(True)
        QMessageBox.critical(self, "오류", error_msg)
    
    def open_result_pdf(self):
        if self.result_pdf_path and os.path.exists(self.result_pdf_path):
            os.startfile(self.result_pdf_path)
    
    def open_result_csv(self):
        if self.result_csv_path and os.path.exists(self.result_csv_path):
            os.startfile(self.result_csv_path)
    
    # === 검색/인쇄 탭 메서드 ===
    
    def select_save_folder(self):
        """저장 폴더 선택"""
        last_dir = self.save_path_edit.text() or ""
        folder_path = QFileDialog.getExistingDirectory(self, "저장 폴더 선택", last_dir)
        if folder_path:
            self.save_path_edit.setText(folder_path)
            self.settings.setValue("search/save_path", folder_path)
    
    def select_search_path(self):
        last_dir = os.path.dirname(self.search_path_edit.text()) if self.search_path_edit.text() else ""
        
        if self.radio_file.isChecked():
            file_path, _ = QFileDialog.getOpenFileName(
                self, "PDF 파일 선택", last_dir, "PDF Files (*.pdf)"
            )
            if file_path:
                self.search_path_edit.setText(file_path)
                self.settings.setValue("search/path", file_path)
                self.settings.setValue("search/is_folder", False)
        else:
            if not last_dir:
                last_dir = self.search_path_edit.text() or ""
            folder_path = QFileDialog.getExistingDirectory(self, "폴더 선택", last_dir)
            if folder_path:
                self.search_path_edit.setText(folder_path)
                self.settings.setValue("search/path", folder_path)
                self.settings.setValue("search/is_folder", True)
    
    def start_search(self):
        search_path = self.search_path_edit.text().strip()
        order_number = self.order_number_edit.text().strip()
        save_folder = self.save_path_edit.text().strip()
        
        if not search_path or not order_number:
            QMessageBox.warning(self, "입력 오류", "경로와 주문번호를 모두 입력하세요.")
            return
        
        if not os.path.exists(search_path):
            QMessageBox.warning(self, "경로 오류", "선택한 경로가 존재하지 않습니다.")
            return
        
        # 저장 폴더 확인 (선택사항)
        if save_folder and not os.path.exists(save_folder):
            reply = QMessageBox.question(
                self, "폴더 생성", 
                f"저장 폴더가 존재하지 않습니다.\n생성하시겠습니까?\n\n{save_folder}",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                try:
                    os.makedirs(save_folder, exist_ok=True)
                except Exception as e:
                    QMessageBox.warning(self, "오류", f"폴더 생성 실패: {e}")
                    return
            else:
                save_folder = ""
        
        self.search_log.clear()
        self.search_btn.setEnabled(False)
        self.stop_search_btn.setEnabled(True)
        self.print_btn.setEnabled(False)
        self.open_folder_btn_search.setEnabled(False)
        
        is_folder = self.radio_folder.isChecked()
        find_all = self.find_all_check.isChecked()
        
        # 설정 저장
        self.settings.setValue("search/use_multiprocess", self.use_mp_check.isChecked())
        self.settings.setValue("search/find_all", find_all)

        self.search_worker = SearchWorker(
            search_path, order_number, is_folder, 
            save_folder if save_folder else None,
            use_multiprocess=self.use_mp_check.isChecked(),
            find_all=find_all
        )
        self.search_worker.progress.connect(self.update_search_log)
        self.search_worker.finished.connect(self.search_finished)
        self.search_worker.error.connect(self.search_error)
        self.search_worker.start()
    
    def stop_search(self):
        """검색 중지"""
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.stop()
        self.stop_search_btn.setEnabled(False)
    
    def update_search_log(self, message):
        self.search_log.append(message)
        self.search_log.verticalScrollBar().setValue(
            self.search_log.verticalScrollBar().maximum()
        )
    
    def search_finished(self, result):
        self.search_btn.setEnabled(True)
        self.stop_search_btn.setEnabled(False)
        if result:
            self.temp_pdf_paths = result['saved_files']
            self.current_save_folder = result['save_folder']
            self.print_btn.setEnabled(True)
            self.open_folder_btn_search.setEnabled(True)
            
            files_info = f"{len(result['saved_files'])}개 PDF 파일" if len(result['saved_files']) > 1 else "PDF 파일"
            
            QMessageBox.information(
                self, "검색 완료",
                f"주문번호를 찾았습니다!\n\n"
                f"원본: {os.path.basename(result['pdf_path'])}\n"
                f"페이지: {', '.join(map(str, result['pages']))}\n"
                f"저장: {files_info}\n"
                f"위치: {result['save_folder']}\n\n"
                f"'PDF 열기' 버튼을 클릭하여 확인하세요."
            )
    
    def search_error(self, error_msg):
        self.search_btn.setEnabled(True)
        self.stop_search_btn.setEnabled(False)
        QMessageBox.critical(self, "오류", error_msg)
    
    def open_for_print(self):
        """저장된 PDF 파일 열기"""
        if not self.temp_pdf_paths:
            return
            
        # 첫 번째 파일 열기
        if os.path.exists(self.temp_pdf_paths[0]):
            open_pdf_for_print(self.temp_pdf_paths[0])
            
            # 여러 파일이 있으면 알림
            if len(self.temp_pdf_paths) > 1:
                reply = QMessageBox.question(
                    self, "여러 파일",
                    f"총 {len(self.temp_pdf_paths)}개의 PDF 파일이 있습니다.\n"
                    f"나머지 파일도 열까요?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    for pdf_path in self.temp_pdf_paths[1:]:
                        if os.path.exists(pdf_path):
                            open_pdf_for_print(pdf_path)
    
    def open_save_folder(self):
        """저장 폴더 열기"""
        if self.current_save_folder and os.path.exists(self.current_save_folder):
            os.startfile(self.current_save_folder)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    # Windows 멀티프로세싱 지원
    import multiprocessing
    multiprocessing.freeze_support()
    
    main()
