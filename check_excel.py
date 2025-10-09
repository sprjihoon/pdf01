"""엑셀 파일 빠른 확인"""
from io_utils import load_excel
from matcher import normalize_name, normalize_phone, normalize_addr

excel_path = r'C:\Users\user\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls'

print("=" * 80)
print("📊 엑셀 파일 분석")
print("=" * 80)

df = load_excel(excel_path)
print(f"\n✅ 총 {len(df)}행 로드됨")

print("\n📋 컬럼 목록:")
print(df.columns.tolist())

print("\n📄 첫 3행 원본 데이터:")
print(df.head(3).to_string())

print("\n" + "=" * 80)
print("🔍 정규화 테스트 (첫 3행)")
print("=" * 80)

for idx in range(min(3, len(df))):
    row = df.iloc[idx]
    print(f"\n[{idx+1}행]")
    print(f"  원본:")
    print(f"    구매자명: {row['구매자명']}")
    print(f"    전화번호: {row['전화번호']}")
    print(f"    주소: {row['주소'][:50]}..." if len(str(row['주소'])) > 50 else f"    주소: {row['주소']}")
    
    name_norm = normalize_name(row['구매자명'])
    phone_norm = normalize_phone(row['전화번호'])
    addr_norm = normalize_addr(row['주소'])
    
    print(f"  정규화:")
    print(f"    이름: '{name_norm}'")
    print(f"    전화: '{phone_norm}'")
    print(f"    주소: '{addr_norm[:50]}...'" if len(addr_norm) > 50 else f"    주소: '{addr_norm}'")
    
    # 경고
    if not name_norm:
        print(f"    ⚠️  이름이 비어있습니다!")
    if not phone_norm:
        print(f"    ⚠️  전화번호가 비어있거나 010으로 시작하지 않습니다!")
    if not addr_norm:
        print(f"    ⚠️  주소가 비어있습니다!")

print("\n" + "=" * 80)
print("이제 PDF 파일 경로를 알려주세요!")
print("=" * 80)

