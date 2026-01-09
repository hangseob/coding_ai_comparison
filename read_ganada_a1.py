import xlwings as xw
import sys

def read_ganada_a1():
    try:
        # 한글 출력을 위해 인코딩 설정
        if sys.stdout.encoding != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
        
        # IRS_Bootstrap_Standalone_Final.xlsm 파일 찾기
        target_book = None
        for book in xw.books:
            if 'IRS_Bootstrap_Standalone_Final' in book.name:
                target_book = book
                break
        
        if not target_book:
            print("IRS_Bootstrap_Standalone_Final.xlsm 파일을 찾을 수 없습니다.")
            return
        
        # 가나다 시트 찾기
        ganada_sheet = None
        for sheet in target_book.sheets:
            if sheet.name == '가나다':
                ganada_sheet = sheet
                break
        
        if not ganada_sheet:
            print("'가나다' 시트를 찾을 수 없습니다.")
            return
        
        # A1 셀의 값 읽기
        value = ganada_sheet.range('A1').value
        print(f"'가나다' 시트의 A1 셀 값: {value}")
        
    except Exception as e:
        print(f"오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    read_ganada_a1()

