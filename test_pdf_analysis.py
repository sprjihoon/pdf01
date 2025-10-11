#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pdfplumber
import re
import pandas as pd

def analyze_pdf(pdf_path):
    """PDF 파일 분석"""
    print(f"PDF 분석: {pdf_path}")
    
    pattern = re.compile(r'\d{15,20}')
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"총 페이지 수: {len(pdf.pages)}")
            
            # 처음 5페이지 확인
            for i in range(min(5, len(pdf.pages))):
                page = pdf.pages[i]
                text = page.extract_text() or ''
                numbers = pattern.findall(text)
                
                print(f"\n페이지 {i+1}:")
                print(f"  추출된 15-20자리 숫자: {numbers[:3]}")  # 처음 3개만
                
                # 텍스트 샘플 (한글 안전하게)
                sample = text[:200].replace('\n', ' ').strip()
                print(f"  텍스트 샘플: {sample[:100]}...")
                
                # 다른 숫자 패턴도 확인
                all_numbers = re.findall(r'\d{10,}', text)
                print(f"  10자리 이상 숫자들: {all_numbers[:5]}")
                
    except Exception as e:
        print(f"PDF 분석 오류: {e}")

def analyze_excel(excel_path):
    """엑셀 파일 분석"""
    print(f"\n엑셀 분석: {excel_path}")
    
    try:
        # 엑셀 읽기
        df = pd.read_excel(excel_path)
        print(f"총 행 수: {len(df)}")
        print(f"컬럼들: {list(df.columns)}")
        
        # 주문번호 컬럼 확인
        if '주문번호' in df.columns:
            print(f"\n주문번호 샘플 (처음 5개):")
            for i, order_num in enumerate(df['주문번호'].head(5)):
                print(f"  {i+1}: {order_num} (길이: {len(str(order_num))})")
        
        # 처음 3행 출력
        print(f"\n처음 3행:")
        print(df.head(3).to_string())
        
    except Exception as e:
        print(f"엑셀 분석 오류: {e}")

if __name__ == "__main__":
    pdf_path = r"C:\Users\one\Documents\카카오톡 받은 파일\구매계약서.pdf"
    excel_path = r"C:\Users\one\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls"
    
    analyze_pdf(pdf_path)
    analyze_excel(excel_path)
