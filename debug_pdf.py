"""
PDF 디버깅 스크립트
PDF와 엑셀의 실제 데이터를 확인하여 매칭 문제를 진단합니다.
"""

import sys
import pandas as pd
import pdfplumber
from io_utils import load_excel
from matcher import (
    normalize_name, normalize_phone, normalize_addr,
    extract_names_from_text, extract_phones_from_text, 
    extract_addresses_from_text
)

def debug_excel(excel_path):
    """엑셀 데이터 디버깅"""
    print("=" * 80)
    print("📊 엑셀 데이터 디버깅")
    print("=" * 80)
    
    df = load_excel(excel_path)
    print(f"총 {len(df)}행 로드됨\n")
    
    print("처음 5행의 원본 데이터:")
    print("-" * 80)
    for idx in range(min(5, len(df))):
        row = df.iloc[idx]
        print(f"\n[{idx+1}행]")
        print(f"  구매자명: {row['구매자명']}")
        print(f"  전화번호: {row['전화번호']}")
        print(f"  주소: {row['주소']}")
    
    print("\n\n처음 5행의 정규화된 데이터:")
    print("-" * 80)
    for idx in range(min(5, len(df))):
        row = df.iloc[idx]
        print(f"\n[{idx+1}행]")
        print(f"  이름: '{normalize_name(row['구매자명'])}'")
        print(f"  전화: '{normalize_phone(row['전화번호'])}'")
        print(f"  주소: '{normalize_addr(row['주소'])}'")
        
        # 빈 값 체크
        if not normalize_name(row['구매자명']):
            print(f"  ⚠️  이름이 비어있습니다!")
        if not normalize_phone(row['전화번호']):
            print(f"  ⚠️  전화번호가 비어있거나 010으로 시작하지 않습니다!")
        if not normalize_addr(row['주소']):
            print(f"  ⚠️  주소가 비어있습니다!")


def debug_pdf(pdf_path):
    """PDF 데이터 디버깅"""
    print("\n\n" + "=" * 80)
    print("📄 PDF 데이터 디버깅")
    print("=" * 80)
    
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"총 {total_pages}페이지\n")
        
        print("처음 3페이지의 원본 텍스트:")
        print("-" * 80)
        
        for i in range(min(3, total_pages)):
            page = pdf.pages[i]
            text = page.extract_text() or ""
            
            print(f"\n[{i+1}페이지] (텍스트 길이: {len(text)}자)")
            print("원본 텍스트 (처음 500자):")
            print(text[:500])
            print("\n추출된 정보:")
            
            # 이름 추출
            names = extract_names_from_text(text)
            print(f"  이름 후보: {names[:5]}")
            print(f"  정규화: {[normalize_name(n) for n in names[:5]]}")
            
            # 전화번호 추출
            phones = extract_phones_from_text(text)
            print(f"  전화 후보: {phones[:5]}")
            print(f"  정규화: {[normalize_phone(p) for p in phones[:5]]}")
            
            # 주소 추출
            addrs = extract_addresses_from_text(text)
            print(f"  주소 후보: {addrs[:3]}")
            print(f"  정규화: {[normalize_addr(a) for a in addrs[:3]]}")
            
            if not names and not phones and not addrs:
                print("  ⚠️  이 페이지에서 아무 정보도 추출되지 않았습니다!")
            
            print("-" * 80)


def main():
    if len(sys.argv) < 3:
        print("사용법: python debug_pdf.py <엑셀파일> <PDF파일>")
        print("\n예시:")
        print('  python debug_pdf.py "data.xlsx" "document.pdf"')
        return
    
    excel_path = sys.argv[1]
    pdf_path = sys.argv[2]
    
    print("\n🔍 PDF-Excel 매칭 디버그 도구")
    print("=" * 80)
    print(f"엑셀: {excel_path}")
    print(f"PDF: {pdf_path}")
    print("=" * 80)
    
    try:
        # 엑셀 디버깅
        debug_excel(excel_path)
        
        # PDF 디버깅
        debug_pdf(pdf_path)
        
        print("\n\n" + "=" * 80)
        print("✅ 디버깅 완료")
        print("=" * 80)
        print("\n💡 문제 해결 방법:")
        print("1. 엑셀 데이터가 제대로 정규화되었는지 확인")
        print("2. PDF에서 텍스트가 제대로 추출되었는지 확인")
        print("3. 추출된 정보들이 실제 엑셀 데이터와 일치하는지 확인")
        print("4. PDF가 이미지 스캔이라면 OCR 처리 필요")
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()

