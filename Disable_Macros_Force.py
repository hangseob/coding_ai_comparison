import xlwings as xw
import os

def disable_all_subs():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    abs_path = os.path.abspath(file_path)
    
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    try:
        wb = app.books.open(abs_path)
        modified = False
        
        for component in wb.api.VBProject.VBComponents:
            vba = component.CodeModule
            count = vba.CountOfLines
            if count <= 0: continue
            
            content = vba.Lines(1, count)
            lines = content.split('\r\n')
            new_lines = []
            
            for line in lines:
                # Sub로 시작하는 라인(매크로)은 모두 주석 처리
                if line.strip().lower().startswith("public sub") or line.strip().lower().startswith("sub "):
                    new_lines.append("' [Python Disabled] " + line)
                    modified = True
                else:
                    new_lines.append(line)
            
            if modified:
                vba.DeleteLines(1, count)
                vba.AddFromString('\r\n'.join(new_lines))
        
        if modified:
            wb.save()
            print("성공! 모든 실행형 매크로(Sub)가 주석 처리되었습니다.")
        else:
            print("주석 처리할 매크로가 없습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    disable_all_subs()
