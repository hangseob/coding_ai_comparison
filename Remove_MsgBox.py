import xlwings as xw
import os

def remove_vba_msgbox():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    abs_path = os.path.abspath(file_path)
    
    print(f"엑셀 파일 수정 중: {abs_path}")
    
    # 이미 열려있는 엑셀 앱 사용 또는 새로 열기
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    
    try:
        wb = app.books.open(abs_path)
        
        # VBA 프로젝트의 모든 컴포넌트(모듈)를 순회하며 MsgBox 찾기
        modified = False
        for component in wb.api.VBProject.VBComponents:
            vba = component.CodeModule
            count = vba.CountOfLines
            if count > 0:
                content = vba.Lines(1, count)
                if 'MsgBox "Bootstrapping Complete!"' in content:
                    # 해당 라인 주석 처리
                    new_content = content.replace('MsgBox "Bootstrapping Complete!"', "' MsgBox \"Bootstrapping Complete!\"")
                    vba.DeleteLines(1, count)
                    vba.AddFromString(new_content)
                    print(f"[{component.Name}] 모듈에서 메시지 박스 제거 완료.")
                    modified = True
        
        if modified:
            wb.save()
            print("성공적으로 저장되었습니다.")
        else:
            print("수정할 메시지 박스 코드를 찾지 못했습니다. 이미 제거되었을 수 있습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")
        print("참고: 엑셀 옵션에서 'VBA 프로젝트 개체 모델에 대한 신뢰'가 체크되어 있어야 합니다.")
    # finally:
    #     if xw.apps.count == 1 and not app.visible:
    #         app.quit()

if __name__ == "__main__":
    remove_vba_msgbox()
