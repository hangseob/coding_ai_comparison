import xlwings as xw
import os

def remove_unsafe_vba():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    abs_path = os.path.abspath(file_path)
    
    if not os.path.exists(abs_path):
        print(f"파일을 찾을 수 없습니다: {abs_path}")
        return

    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=False)
    
    try:
        wb = app.books.open(abs_path)
        
        # VBA 프로젝트의 모듈 순회
        modified = False
        for component in wb.api.VBProject.VBComponents:
            vba = component.CodeModule
            count = vba.CountOfLines
            if count <= 0: continue
            
            content = vba.Lines(1, count)
            new_lines = []
            lines = content.split('\r\n')
            
            skip_mode = False
            # 삭제할 프로시저 목록
            targets = ["Sub RunBootstrap()", "Sub SolveStep("]
            
            for line in lines:
                trimmed = line.strip()
                
                # 프로시저 시작점 확인
                if any(trimmed.startswith(t) for t in targets):
                    skip_mode = True
                    modified = True
                    print(f"[{component.Name}] 삭제 중: {trimmed}")
                    continue
                
                # 프로시저 종료점 확인
                if skip_mode and (trimmed == "End Sub"):
                    skip_mode = False
                    continue
                
                if not skip_mode:
                    new_lines.append(line)
            
            if modified:
                vba.DeleteLines(1, count)
                if new_lines:
                    vba.AddFromString('\r\n'.join(new_lines))
        
        if modified:
            wb.save()
            print("성공! 오류 유발 매크로가 제거되었습니다.")
        else:
            print("삭제할 매크로를 찾지 못했습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")
        print("참고: 엑셀 옵션에서 'VBA 프로젝트 개체 모델에 대한 신뢰'가 체크되어 있어야 합니다.")

if __name__ == "__main__":
    remove_unsafe_vba()
