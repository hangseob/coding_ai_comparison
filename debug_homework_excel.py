import xlwings as xw

def debug_excel():
    print("엑셀 워크북 탐색...")
    found = False
    for app in xw.apps:
        for book in app.books:
            print(f"워크북 발견: {book.name}")
            if "6.4" in book.name:
                print(f"--- '{book.name}' 상세 정보 ---")
                for sheet in book.sheets:
                    print(f"  시트: {sheet.name}")
                    try:
                        tables = sheet.api.ListObjects
                        for t in tables:
                            print(f"    테이블: {t.name}")
                            df = sheet.range(t.Range.Address).options(pd.DataFrame, index=False, header=True).value
                            print(f"      컬럼: {list(df.columns)}")
                    except:
                        pass
                found = True
    if not found:
        print("'.6.4'를 포함하는 워크북을 찾지 못했습니다.")

if __name__ == "__main__":
    debug_excel()
