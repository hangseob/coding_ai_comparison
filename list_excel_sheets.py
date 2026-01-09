import xlwings as xw
import sys

def list_all_excel_sheets():
    try:
        # 한글 출력을 위해 인코딩 설정
        if sys.stdout.encoding != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
            
        print("=" * 50)
        print("현재 열려 있는 모든 엑셀 파일 및 시트 목록")
        print("=" * 50)
        
        books = xw.books
        if not books:
            print("열려 있는 엑셀 파일이 없습니다.")
            return

        for book in books:
            print(f"\n파일 이름: {book.name}")
            print("-" * 30)
            for sheet in book.sheets:
                print(f"  - 시트 이름: {sheet.name}")
        
        print("\n" + "=" * 50)
                
    except Exception as e:
        print(f"오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    list_all_excel_sheets()

