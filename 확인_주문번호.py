"""주문번호 확인"""
from io_utils import load_excel
from matcher import normalize_order_number

excel_path = r'C:\Users\user\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls'

print("=" * 60)
print("주문번호 확인")
print("=" * 60)

try:
    df = load_excel(excel_path)
    print(f"✅ 총 {len(df)}행 로드됨\n")
    
    print("컬럼:", df.columns.tolist())
    
    print("\n첫 5행의 주문번호:")
    for idx in range(min(5, len(df))):
        row = df.iloc[idx]
        order = row['주문번호']
        order_norm = normalize_order_number(order)
        print(f"  [{idx+1}] {order} → '{order_norm}'")
    
    print("\n" + "=" * 60)
    print("✅ 주문번호 컬럼이 제대로 로드되었습니다!")
    print("=" * 60)
    
except Exception as e:
    print(f"❌ 오류: {e}")
    import traceback
    traceback.print_exc()

