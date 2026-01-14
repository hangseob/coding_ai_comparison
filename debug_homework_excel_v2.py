import xlwings as xw
import pandas as pd

def debug_excel():
    print("엑셀 워크북 탐색...")
    for app in xw.apps:
        for book in app.books:
            if "6.4" in book.name:
                print(f"\n[워크북: {book.name}]")
                for sheet in book.sheets:
                    print(f"  - 시트: {sheet.name}")
                    try:
                        for tbl in sheet.api.ListObjects:
                            print(f"    * 테이블: {tbl.Name}")
                            # 테이블의 모든 컬럼 이름 출력
                            df = sheet.range(tbl.Range.Address).options(pd.DataFrame, index=False, header=True).value
                            print(f"      컬럼들: {list(df.columns)}")
                    except Exception as e:
                        print(f"    * 에러: {e}")

if __name__ == "__main__":
    debug_excel()
