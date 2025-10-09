"""
예시 파일 생성 스크립트
엑셀 파일과 샘플 데이터를 생성합니다.
"""

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def create_example_excel():
    """예시 엑셀 파일 생성"""
    data = {
        '순번': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        '이름': ['홍길동', '김철수', '이영희', '박민수', '정수진', 
                '최동욱', '강민지', '윤서연', '임재현', '오지은'],
        '전화번호': [
            '010-1234-5678',
            '010-9876-5432',
            '010-5555-1234',
            '010-3333-7777',
            '010-8888-9999',
            '010-1111-2222',
            '010-4444-5555',
            '010-6666-7777',
            '010-2222-3333',
            '010-9999-0000'
        ],
        '주문번호': [
            'ORD-2024-001',
            'ORD-2024-002',
            'ORD-2024-003',
            'ORD-2024-004',
            'ORD-2024-005',
            'ORD-2024-006',
            'ORD-2024-007',
            'ORD-2024-008',
            'ORD-2024-009',
            'ORD-2024-010'
        ],
        '금액': [50000, 75000, 120000, 95000, 60000, 
                80000, 110000, 85000, 70000, 90000]
    }
    
    df = pd.DataFrame(data)
    df.to_excel('example_purchasers.xlsx', index=False, engine='openpyxl')
    print("✅ 예시 엑셀 파일이 생성되었습니다: example_purchasers.xlsx")

def create_example_pdf_simple():
    """간단한 텍스트 기반 예시 PDF 생성 (한글 폰트 없이)"""
    try:
        from pypdf import PdfWriter
        import io
        
        # 간단한 PDF 생성 (reportlab 없이)
        c = canvas.Canvas('example_document.pdf', pagesize=A4)
        width, height = A4
        
        # 예시 구매자 정보 (영문으로)
        purchasers = [
            {'name': 'Hong Gildong', 'phone': '010-1234-5678', 'order': 'ORD-2024-001'},
            {'name': 'Kim Chulsoo', 'phone': '010-9876-5432', 'order': 'ORD-2024-002'},
            {'name': 'Lee Younghee', 'phone': '010-5555-1234', 'order': 'ORD-2024-003'},
            {'name': 'Park Minsoo', 'phone': '010-3333-7777', 'order': 'ORD-2024-004'},
            {'name': 'Jung Soojin', 'phone': '010-8888-9999', 'order': 'ORD-2024-005'},
        ]
        
        # 각 페이지 생성 (순서를 섞어서)
        import random
        shuffled = purchasers.copy()
        random.shuffle(shuffled)
        
        for i, person in enumerate(shuffled, 1):
            # 페이지 헤더
            c.setFont("Helvetica-Bold", 20)
            c.drawString(50, height - 50, f"Purchase Order - Page {i}")
            
            # 구매자 정보
            c.setFont("Helvetica", 14)
            y_position = height - 100
            
            c.drawString(50, y_position, f"Name: {person['name']}")
            y_position -= 30
            c.drawString(50, y_position, f"Phone: {person['phone']}")
            y_position -= 30
            c.drawString(50, y_position, f"Order Number: {person['order']}")
            y_position -= 50
            
            # 상품 정보
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, "Product Information:")
            y_position -= 25
            
            c.setFont("Helvetica", 11)
            c.drawString(70, y_position, "- Product: Premium Widget")
            y_position -= 20
            c.drawString(70, y_position, "- Quantity: 2")
            y_position -= 20
            c.drawString(70, y_position, f"- Total: ${random.randint(50, 150)},000")
            
            # 푸터
            c.setFont("Helvetica-Oblique", 9)
            c.drawString(50, 50, f"Order: {person['order']} | Contact: {person['phone']}")
            
            c.showPage()
        
        c.save()
        print("✅ 예시 PDF 파일이 생성되었습니다: example_document.pdf")
        print("   (페이지 순서가 섞여 있습니다. 엑셀 순서대로 정렬해보세요!)")
        
    except ImportError as e:
        print(f"⚠️  PDF 생성 실패: {e}")
        print("   reportlab 패키지를 설치해주세요: pip install reportlab")

if __name__ == '__main__':
    print("=" * 60)
    print("예시 파일 생성 중...")
    print("=" * 60)
    print()
    
    # 엑셀 파일 생성
    create_example_excel()
    
    # PDF 파일 생성
    create_example_pdf_simple()
    
    print()
    print("=" * 60)
    print("완료!")
    print("=" * 60)
    print()
    print("생성된 파일:")
    print("  1. example_purchasers.xlsx (구매자 정보)")
    print("  2. example_document.pdf (정렬 전 PDF)")
    print()
    print("테스트 방법:")
    print("  1. main.py 실행")
    print("  2. 엑셀: example_purchasers.xlsx 선택")
    print("  3. PDF: example_document.pdf 선택")
    print("  4. 키 컬럼: 이름 또는 전화번호")
    print("  5. 실행 버튼 클릭!")
    print()

