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
    QProgressBar, QGroupBox, QMessageBox, QCheckBox, QSpinBox
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

from io_utils import load_excel, get_output_filenames, save_pdf, save_report, is_text_based_pdf
from matcher import extract_pages, match_rows_to_pages, reorder_pdf


class ProcessingWorker(QThread):
    """백그라운드 작업 스레드"""
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
                    self.progress.emit(
                        f"   엑셀 {row_idx+1}행 ({row['구매자명']}) → "
                        f"PDF {detail['page_idx']+1}페이지 (점수: {detail['score']:.1f})"
                    )
                else:
                    self.progress.emit(f"   엑셀 {row_idx+1}행 ({row['구매자명']}) → 매칭 실패")
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


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Excel Matcher - 구매자 순서로 PDF 자동 정렬")
        self.setMinimumSize(900, 750)
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        layout = QVBoxLayout(central_widget)
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
        file_group = QGroupBox("1️⃣ 파일 선택")
        file_layout = QVBoxLayout()
        
        # 엑셀
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("엑셀 파일:"))
        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText("구매자명, 전화번호, 주소 컬럼이 있는 엑셀 파일...")
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
        
        # 출력 폴더
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("출력 폴더:"))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("결과 파일이 저장될 폴더...")
        output_layout.addWidget(self.output_edit)
        output_btn = QPushButton("찾아보기")
        output_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(output_btn)
        file_layout.addLayout(output_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 옵션 그룹
        option_group = QGroupBox("2️⃣ 매칭 옵션")
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
            "💡 매칭 기준: 구매자명 + 전화번호 + 주소 (3가지 모두 일치해야 함)\n"
            "   • 기본: 정확 일치만 인정 (권장)\n"
            "   • 유사도 매칭: 오타나 표기 차이가 있을 때 보조적으로 사용 (임계값 조정 가능)"
        )
        help_label.setStyleSheet("color: #555; font-size: 9pt; padding: 10px; background: #f5f5f5; border-radius: 5px;")
        option_layout.addWidget(help_label)
        
        option_group.setLayout(option_layout)
        layout.addWidget(option_group)
        
        # 실행 버튼
        self.run_btn = QPushButton("▶️  PDF 정렬 실행")
        self.run_btn.setMinimumHeight(55)
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 15pt;
                font-weight: bold;
                border-radius: 8px;
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
        progress_group = QGroupBox("3️⃣ 진행 상황")
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
    
    def on_fuzzy_changed(self, state):
        """유사도 매칭 체크박스 변경 이벤트"""
        self.threshold_spin.setEnabled(state == Qt.Checked)
    
    def browse_excel(self):
        """엑셀 파일 찾아보기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.excel_edit.setText(file_path)
            self.log(f"✓ 엑셀 선택: {os.path.basename(file_path)}")
    
    def browse_pdf(self):
        """PDF 파일 찾아보기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "PDF 파일 선택", "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        if file_path:
            self.pdf_edit.setText(file_path)
            self.log(f"✓ PDF 선택: {os.path.basename(file_path)}")
    
    def browse_output(self):
        """출력 폴더 찾아보기"""
        folder_path = QFileDialog.getExistingDirectory(self, "출력 폴더 선택")
        if folder_path:
            self.output_edit.setText(folder_path)
            self.log(f"✓ 출력 폴더: {folder_path}")
    
    def log(self, message):
        """로그 메시지 추가"""
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
    
    def run_processing(self):
        """처리 실행"""
        # 입력 검증
        excel_path = self.excel_edit.text().strip()
        pdf_path = self.pdf_edit.text().strip()
        output_dir = self.output_edit.text().strip()
        
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "입력 오류", "유효한 엑셀 파일을 선택하세요.")
            return
        
        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "입력 오류", "유효한 PDF 파일을 선택하세요.")
            return
        
        if not output_dir or not os.path.exists(output_dir):
            QMessageBox.warning(self, "입력 오류", "유효한 출력 폴더를 선택하세요.")
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
