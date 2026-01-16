import xlwings as xw
import os

def cleanup_sheets():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    if not os.path.exists(file_path):
        print("파일을 찾을 수 없습니다.")
        return
        
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    try:
        wb = app.books.open(file_path)
        sheets_to_delete = ["Validation_Deposit", "Validation_IRS"]
        
        deleted_count = 0
        for sheet_name in sheets_to_delete:
            try:
                wb.sheets[sheet_name].delete()
                print(f"시트 삭제됨: {sheet_name}")
                deleted_count += 1
            except:
                print(f"시트가 없거나 삭제할 수 없음: {sheet_name}")
        
        if deleted_count > 0:
            wb.save()
            print("파일 저장 완료.")
        else:
            print("삭제할 시트가 없습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    cleanup_sheets()
