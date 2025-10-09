"""엑셀 컬럼 확인"""
import pandas as pd

excel_path = r'C:\Users\user\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls'

df = pd.read_excel(excel_path, engine='xlrd')

print("=" * 60)
print("엑셀 파일의 모든 컬럼:")
print("=" * 60)
for i, col in enumerate(df.columns, 1):
    print(f"{i}. '{col}'")

print("\n" + "=" * 60)
print("첫 3행 데이터:")
print("=" * 60)
print(df.head(3).to_string())

