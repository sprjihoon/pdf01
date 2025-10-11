#!/usr/bin/env python3
import pdfplumber
import pandas as pd
from matcher import normalize_order_number

def find_matching_orders():
    # 파일 경로
    pdf_path = r"C:\Users\one\Documents\카카오톡 받은 파일\구매계약서.pdf"
    excel_path = r"C:\Users\one\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls"
    
    # 엑셀에서 주문번호 추출
    print("엑셀 주문번호 추출 중...")
    df = pd.read_excel(excel_path)
    excel_orders = set()
    for order in df['주문번호']:
        normalized = normalize_order_number(str(order))
        if normalized:
            excel_orders.add(normalized)
    
    print(f"엑셀 고유 주문번호: {len(excel_orders)}개")
    print(f"샘플 (처음 5개): {list(excel_orders)[:5]}")
    
    # PDF에서 주문번호 추출
    print(f"\nPDF 주문번호 추출 중...")
    pdf_orders = set()
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if i >= 5:  # 처음 5페이지만 확인
                break
            text = page.extract_text() or ''
            # 15-20자리 숫자 찾기
            import re
            numbers = re.findall(r'\d{15,20}', text)
            
            for num in numbers:
                normalized = normalize_order_number(num)
                if normalized:
                    pdf_orders.add(normalized)
    
    print(f"PDF 고유 주문번호 (첫 5페이지): {len(pdf_orders)}개")
    print(f"샘플: {list(pdf_orders)[:5]}")
    
    # 일치하는 주문번호 찾기
    matching_orders = excel_orders.intersection(pdf_orders)
    print(f"\n일치하는 주문번호: {len(matching_orders)}개")
    
    if matching_orders:
        print(f"매칭 성공 예시: {list(matching_orders)[:3]}")
    else:
        print("❌ 일치하는 주문번호가 없습니다!")
        print(f"\n엑셀 첫 번째: {list(excel_orders)[0] if excel_orders else 'None'}")
        print(f"PDF 첫 번째: {list(pdf_orders)[0] if pdf_orders else 'None'}")

if __name__ == "__main__":
    find_matching_orders()
