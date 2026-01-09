import xlwings as xw
import sys

def read_calculation_top():
    try:
        if sys.stdout.encoding != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
        
        target_book = None
        for book in xw.books:
            if 'IRS_Bootstrap_Standalone_Final' in book.name:
                target_book = book
                break
        
        if not target_book:
            print("IRS_Bootstrap_Standalone_Final.xlsm 파일을 찾을 수 없습니다.")
            return
        
        calc_sheet = None
        for sheet in target_book.sheets:
            if sheet.name == 'Calculation':
                calc_sheet = sheet
                break
        
        if not calc_sheet:
            print("'Calculation' 시트를 찾을 수 없습니다.")
            return
        
        a4_value = calc_sheet.range('A4').value
        a5_value = calc_sheet.range('A5').value
        
        print(f"A4 셀의 값: {a4_value}")
        print(f"A5 셀의 값: {a5_value}")
        
    except Exception as e:
        print(f"오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    read_calculation_top()

