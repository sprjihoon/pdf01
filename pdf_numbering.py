"""
PDF 페이지 넘버링 추가
- 각 페이지 상단 오른쪽에 5pt 크기의 페이지 번호 추가
- 같은 주문번호는 같은 번호 부여
"""

import os
from typing import Dict
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO


def add_page_numbers_by_order(input_pdf_path: str, output_pdf_path: str, 
                               page_to_order_number: Dict[int, int], 
                               font_size: int = 5):
    """
    PDF의 각 페이지 상단 오른쪽에 주문번호 기준 페이지 번호 추가
    같은 주문번호는 같은 번호를 부여
    
    Args:
        input_pdf_path: 입력 PDF 경로
        output_pdf_path: 출력 PDF 경로
        page_to_order_number: {페이지_인덱스: 주문_번호} 매핑 (0-based)
        font_size: 페이지 번호 폰트 크기 (기본값: 5pt)
    """
    # 입력 PDF 읽기
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    total_pages = len(reader.pages)
    
    for page_num in range(total_pages):
        page = reader.pages[page_num]
        
        # 주문번호 기준 페이지 번호 가져오기 (매핑되지 않은 페이지는 None)
        order_number = page_to_order_number.get(page_num, None)
        
        # 매핑된 페이지만 번호 추가
        if order_number is not None:
            # 페이지 크기 가져오기
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            
            # 페이지 번호 오버레이 생성
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=(page_width, page_height))
            
            page_number_text = str(order_number)
            
            # 폰트 설정 (5pt)
            can.setFont("Helvetica", font_size)
            
            # 상단 오른쪽 위치 계산
            # 오른쪽에서 10pt 떨어진 위치, 상단에서 5pt 아래
            x_position = page_width - 15  # 오른쪽 여백
            y_position = page_height - 10  # 상단 여백
            
            # 페이지 번호 그리기
            can.drawRightString(x_position, y_position, page_number_text)
            
            can.save()
            
            # 오버레이 PDF 생성
            packet.seek(0)
            overlay_pdf = PdfReader(packet)
            overlay_page = overlay_pdf.pages[0]
            
            # 원본 페이지와 오버레이 합치기
            page.merge_page(overlay_page)
        
        # 새 PDF에 추가 (번호 있든 없든 페이지는 추가)
        writer.add_page(page)
    
    # 출력 PDF 저장
    with open(output_pdf_path, 'wb') as output_file:
        writer.write(output_file)


def add_page_numbers(input_pdf_path: str, output_pdf_path: str, font_size: int = 5):
    """
    PDF의 각 페이지 상단 오른쪽에 페이지 번호 추가 (순차적)
    
    Args:
        input_pdf_path: 입력 PDF 경로
        output_pdf_path: 출력 PDF 경로
        font_size: 페이지 번호 폰트 크기 (기본값: 5pt)
    """
    # 입력 PDF 읽기
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    total_pages = len(reader.pages)
    
    for page_num in range(total_pages):
        page = reader.pages[page_num]
        
        # 페이지 크기 가져오기
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        # 페이지 번호 오버레이 생성
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=(page_width, page_height))
        
        # 페이지 번호 텍스트 (1-based)
        page_number_text = str(page_num + 1)
        
        # 폰트 설정 (5pt)
        can.setFont("Helvetica", font_size)
        
        # 상단 오른쪽 위치 계산
        # 오른쪽에서 10pt 떨어진 위치, 상단에서 5pt 아래
        x_position = page_width - 15  # 오른쪽 여백
        y_position = page_height - 10  # 상단 여백
        
        # 페이지 번호 그리기
        can.drawRightString(x_position, y_position, page_number_text)
        
        can.save()
        
        # 오버레이 PDF 생성
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        overlay_page = overlay_pdf.pages[0]
        
        # 원본 페이지와 오버레이 합치기
        page.merge_page(overlay_page)
        
        # 새 PDF에 추가
        writer.add_page(page)
    
    # 출력 PDF 저장
    with open(output_pdf_path, 'wb') as output_file:
        writer.write(output_file)


def add_page_numbers_to_temp(input_pdf_path: str, font_size: int = 5) -> str:
    """
    임시 파일로 페이지 번호가 추가된 PDF 생성
    
    Args:
        input_pdf_path: 입력 PDF 경로
        font_size: 페이지 번호 폰트 크기
        
    Returns:
        임시 PDF 파일 경로
    """
    import tempfile
    
    # 임시 파일 생성
    temp_fd, temp_path = tempfile.mkstemp(suffix='_numbered.pdf')
    os.close(temp_fd)
    
    # 페이지 번호 추가
    add_page_numbers(input_pdf_path, temp_path, font_size)
    
    return temp_path

