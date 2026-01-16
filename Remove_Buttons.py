import xlwings as xw
import os

def remove_excel_buttons():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    if not os.path.exists(file_path):
        print("파일을 찾을 수 없습니다.")
        return
        
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    try:
        wb = app.books.open(file_path)
        ws = wb.sheets["Main"]
        
        count = 0
        # 1. 시트 내의 모든 Shape(도형, 버튼 등) 확인
        for shape in ws.api.Shapes:
            try:
                # Type 8: Form Control (버튼 등), Type 12: OLE Control
                # 또는 매크로(OnAction)가 연결된 경우 삭제
                if shape.Type in [8, 12] or (hasattr(shape, 'OnAction') and shape.OnAction != ""):
                    shape.Delete()
                    count += 1
            except Exception as e:
                # 특정 개체는 삭제 권한이 없을 수 있음 (무시)
                continue
        
        if count > 0:
            wb.save()
            print(f"성공! {count}개의 버튼/도형이 삭제되었습니다.")
        else:
            print("삭제할 버튼이나 매크로 도형을 찾지 못했습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    remove_excel_buttons()
