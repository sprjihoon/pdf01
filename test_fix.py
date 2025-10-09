"""전화번호 정규화 수정 테스트"""
from matcher import normalize_phone

print("=" * 60)
print("전화번호 정규화 테스트")
print("=" * 60)

test_cases = [
    ("1026417075", "엑셀 형식 (앞 0 누락)"),
    ("01026417075", "정상 형식"),
    ("010-2641-7075", "하이픈 포함"),
    ("1052944324", "엑셀 형식 (앞 0 누락)"),
    ("01052944324", "정상 형식"),
    ("01025847576", "PDF 형식"),
]

for phone, desc in test_cases:
    result = normalize_phone(phone)
    status = "✅" if result else "❌"
    print(f"{status} {desc:20s} '{phone}' → '{result}'")

print("\n" + "=" * 60)
print("수정 완료! 이제 다시 프로그램을 실행하세요.")
print("=" * 60)

