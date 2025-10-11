#!/usr/bin/env python3
from matcher import normalize_order_number

# 테스트
pdf_order = '0100012025100100075'
excel_order = '100012025100900021'

pdf_normalized = normalize_order_number(pdf_order)
excel_normalized = normalize_order_number(excel_order)

print(f'PDF 원본: {pdf_order}')
print(f'PDF 정규화: {pdf_normalized}')
print(f'엑셀 원본: {excel_order}')  
print(f'엑셀 정규화: {excel_normalized}')
print(f'길이 비교: PDF={len(pdf_normalized)}, 엑셀={len(excel_normalized)}')

# 다른 예시들도 테스트
test_cases = [
    '0100012025100100075',  # PDF 형태
    '100012025100900021',   # 엑셀 형태
    '01234567890123456789', # 앞에 01 붙은 긴 숫자
    '1234567890123456789',  # 1로 시작하는 긴 숫자
    '0123456789012345',     # 짧은 숫자
]

print('\n다양한 테스트 케이스:')
for i, case in enumerate(test_cases):
    normalized = normalize_order_number(case)
    print(f'{i+1}: {case} → {normalized}')
