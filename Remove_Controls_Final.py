import xlwings as xw
import os

def force_remove_all_controls():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    
    try:
        wb = app.books.open(file_path)
        ws = wb.sheets["Main"]
        
        count = 0
        
        # 방식 1: Buttons 컬렉션 (양식 컨트롤 버튼)
        try:
            for btn in ws.api.Buttons():
                btn.Delete()
                count += 1
        except: pass

        # 방식 2: OLEObjects (ActiveX 컨트롤)
        try:
            for ole in ws.api.OLEObjects():
                ole.Delete()
                count += 1
        except: pass
        
        # 방식 3: 일반 Shapes (매크로가 지정된 도형)
        try:
            for i in range(ws.api.Shapes.Count, 0, -1):
                shp = ws.api.Shapes(i)
                if shp.OnAction != "":
                    shp.Delete()
                    count += 1
        except: pass

        if count > 0:
            wb.save()
            print(f"총 {count}개의 컨트롤을 삭제했습니다.")
        else:
            print("삭제할 컨트롤이 없습니다.")
            
    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    force_remove_all_controls()
