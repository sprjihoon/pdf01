"""디버그 실행"""
import subprocess
import sys

excel_path = r'C:\Users\user\Documents\카카오톡 받은 파일\Ordering_data_20251009_spring.xls'
pdf_path = r'C:\Users\user\Documents\카카오톡 받은 파일\구매계약서.pdf'

print(f"엑셀: {excel_path}")
print(f"PDF: {pdf_path}")
print()

# debug_pdf.py를 sys.argv로 실행
sys.argv = ['debug_pdf.py', excel_path, pdf_path]

# debug_pdf.py 내용 실행
exec(open('debug_pdf.py', encoding='utf-8').read())

