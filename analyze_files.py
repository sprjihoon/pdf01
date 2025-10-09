"""파일 분석"""
import pandas as pd
import pdfplumber
from io_utils import load_excel
from matcher import (
    normalize_name, normalize_phone, normalize_addr,
    extract_names_from_text, extract_phones_from_text, 
    extract_addresses_from_text
)

# 파일 경로
excel_path = r'C:\Users\user\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls'
pdf_path = r'C:\Users\user\Documents\카카오톡 받은 파일\구매계약서.pdf'

print("=" * 80)
print("🔍 PDF-Excel 매칭 디버그")
print("=" * 80)
print(f"엑셀: {excel_path}")
print(f"PDF: {pdf_path}")
print("=" * 80)

# 1. 엑셀 분석
print("\n📊 엑셀 데이터 분석")
print("-" * 80)

try:
    df = load_excel(excel_path)
    print(f"✅ 총 {len(df)}행 로드됨\n")
    
    print("컬럼 목록:", df.columns.tolist())
    
    print("\n첫 3행의 데이터:")
    for idx in range(min(3, len(df))):
        row = df.iloc[idx]
        print(f"\n[{idx+1}행]")
        print(f"  구매자명: {row['구매자명']}")
        print(f"  전화번호: {row['전화번호']}")
        print(f"  주소: {str(row['주소'])[:60]}...")
        
        name_norm = normalize_name(row['구매자명'])
        phone_norm = normalize_phone(row['전화번호'])
        addr_norm = normalize_addr(row['주소'])
        
        print(f"  → 정규화:")
        print(f"     이름: '{name_norm}'")
        print(f"     전화: '{phone_norm}'")
        print(f"     주소: '{addr_norm[:40]}...'" if len(addr_norm) > 40 else f"     주소: '{addr_norm}'")
        
        if not name_norm:
            print(f"     ⚠️  이름 비어있음!")
        if not phone_norm:
            print(f"     ⚠️  전화번호 비어있음 (010이 아니거나 형식 불일치)!")
        if not addr_norm:
            print(f"     ⚠️  주소 비어있음!")

except Exception as e:
    print(f"❌ 엑셀 오류: {e}")
    import traceback
    traceback.print_exc()

# 2. PDF 분석
print("\n\n📄 PDF 데이터 분석")
print("-" * 80)

try:
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"✅ 총 {total_pages}페이지\n")
        
        for i in range(min(3, total_pages)):
            page = pdf.pages[i]
            text = page.extract_text() or ""
            
            print(f"\n[{i+1}페이지]")
            print(f"  텍스트 길이: {len(text)}자")
            
            if len(text) == 0:
                print(f"  ⚠️  텍스트가 없습니다! (이미지 스캔 PDF일 가능성)")
                continue
            
            print(f"  원본 텍스트 (처음 200자):")
            print(f"  {text[:200]}")
            
            names = extract_names_from_text(text)
            phones = extract_phones_from_text(text)
            addrs = extract_addresses_from_text(text)
            
            print(f"\n  추출된 정보:")
            print(f"    이름 후보 ({len(names)}개): {names[:5]}")
            print(f"    전화 후보 ({len(phones)}개): {phones[:5]}")
            print(f"    주소 후보 ({len(addrs)}개): {[a[:30]+'...' for a in addrs[:3]]}")
            
            print(f"\n  정규화:")
            print(f"    이름: {[normalize_name(n) for n in names[:5]]}")
            print(f"    전화: {[normalize_phone(p) for p in phones[:5]]}")
            print(f"    주소: {[normalize_addr(a)[:30]+'...' for a in addrs[:3]]}")
            
            if not names and not phones and not addrs:
                print(f"  ⚠️  아무 정보도 추출되지 않았습니다!")

except Exception as e:
    print(f"❌ PDF 오류: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("✅ 분석 완료")
print("=" * 80)

