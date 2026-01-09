import xlwings as xw
import os

def create_bootstrap_excel():
    # 1. 엑셀 앱 실행
    app = xw.App(visible=True)
    wb = app.books.add()
    
    # 2. VBA 코드 작성 (해찾기 및 보간 로직)
    vba_code = """
Option Explicit

'--- 1. Tenor 문자열을 숫자(연)로 변환 ---
Function ParseTenor(ByVal tStr As Variant) As Double
    Dim v As Double
    Dim unit As String
    Dim s As String
    
    If IsNumeric(tStr) Then
        ParseTenor = CDbl(tStr)
        Exit Function
    End If
    
    s = UCase(Trim(CStr(tStr)))
    v = Val(s)
    unit = Right(s, 1)
    
    Select Case unit
        Case "D": ParseTenor = v / 365
        Case "W": ParseTenor = (v * 7) / 365
        Case "M": ParseTenor = v / 12
        Case "Y": ParseTenor = v
        Case Else: ParseTenor = v
    End Select
End Function

'--- 2. Log-Linear Interpolation (UDF) ---
Function LogLinearDF(t As Double, jumpNodes As Range, solvedForwards As Range) As Double
    Dim n As Long
    Dim totalIntegral As Double
    Dim prevNode As Double
    Dim i As Long
    Dim currentFwd As Double
    Dim node As Double
    
    If t <= 0 Then
        LogLinearDF = 1#
        Exit Function
    End If
    
    n = jumpNodes.Cells.Count
    totalIntegral = 0
    prevNode = 0
    currentFwd = 0
    
    For i = 1 To n
        node = jumpNodes.Cells(i).Value
        If node <= 0 Or IsEmpty(jumpNodes.Cells(i)) Then Exit For
        
        currentFwd = solvedForwards.Cells(i).Value
        
        If t <= node Then
            totalIntegral = totalIntegral + currentFwd * (t - prevNode)
            LogLinearDF = Exp(-totalIntegral)
            Exit Function
        Else
            totalIntegral = totalIntegral + currentFwd * (node - prevNode)
            prevNode = node
        End If
    Next i
    
    LogLinearDF = Exp(-(totalIntegral + currentFwd * (t - prevNode)))
End Function

'--- 3. 순차적 해찾기 엔진 ---
Sub RunBootstrap()
    Dim wsMain As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim stepIdx As Long
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "No data found to bootstrap.", vbExclamation
        Exit Sub
    End If
    
    ' 결과값 초기화
    wsMain.Range("H5:H" & lastRow).Value = 0.03
    
    For stepIdx = 1 To (lastRow - 4)
        SolveStep stepIdx
    Next stepIdx
    
    MsgBox "Bootstrapping Complete!", vbInformation
End Sub

Private Sub SolveStep(idx As Long)
    Dim wsMain As Worksheet
    Dim wsCalc As Worksheet
    Dim targetCell As Range
    Dim f As Double
    Dim errorVal As Double
    Dim f_plus As Double
    Dim error_plus As Double
    Dim deriv As Double
    Dim delta As Double
    Dim iter As Integer
    Dim instType As String
    Dim tblSummary As ListObject
    
    Dim tempVal As Variant
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    instType = wsMain.Range("B" & (idx + 4)).Value
    
    If instType = "Deposit" Then
        Set wsCalc = ThisWorkbook.Sheets("검증(Deposit)")
    Else
        Set wsCalc = ThisWorkbook.Sheets("검증(IRS)")
    End If
    
    Set tblSummary = wsCalc.ListObjects(1) ' 상단 요약 테이블
    Set targetCell = wsMain.Range("H" & (idx + 4))
    
    ' 현재 계산 타겟 테너 설정
    tblSummary.ListColumns("Target Tenor").DataBodyRange.Cells(1).Value = wsMain.Range("C" & (idx + 4)).Value
    
    delta = 0.00001
    f = targetCell.Value
    
    For iter = 1 To 30
        targetCell.Value = f
        Application.CalculateFull ' 수식 강제 재계산
        
        tempVal = tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value
        If IsError(tempVal) Then
            errorVal = 1000000# ' 에러 발생 시 매우 큰 값으로 대체
        Else
            errorVal = CDbl(tempVal)
        End If
        
        If Abs(errorVal) < 0.000000000001 Then Exit For
        
        targetCell.Value = f + delta
        Application.CalculateFull
        
        tempVal = tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value
        if IsError(tempVal) Then
            error_plus = 1000000#
        Else
            error_plus = CDbl(tempVal)
        End If
        
        deriv = (error_plus - errorVal) / delta
        
        If Abs(deriv) < 1E-15 Then Exit For
        
        f = f - errorVal / deriv
        If f < -0.1 Then f = 0.01
    Next iter
    
    targetCell.Value = f
End Sub
"""

    # 3. VBA 모듈 주입
    try:
        vba_module = wb.api.VBProject.VBComponents.Add(1)
        vba_module.CodeModule.AddFromString(vba_code)
    except Exception as e:
        print(f"VBA 주입 실패: {e}")

    # 4. Main 시트 구성
    ws_main = wb.sheets[0]
    ws_main.name = "Main"
    
    ws_main.range("A1").value = "KRW IRS Standalone Bootstrapper"
    ws_main.range("A1").api.Font.Size = 16
    ws_main.range("A1").api.Font.Bold = True
    
    headers = ["No", "Type", "Inst. Tenor", "Tenor(Y)", "Market Rate", "Jump Node", "Node(Y)", "Solved Forward"]
    ws_main.range("A4").value = headers
    ws_main.range("A4:H4").color = (200, 200, 200)
    ws_main.range("A4:H4").api.Font.Bold = True
    
    inputs = [
        [1, "Deposit", "1D", 0, 0.05, "1M", 0, 0.05],
        [2, "Deposit", "3M", 0, 0.05, "4M", 0, 0.05],
        [3, "IRS", "6M", 0, 0.05, "7M", 0, 0.05],
        [4, "IRS", "1Y", 0, 0.05, "13M", 0, 0.05],
        [5, "IRS", "2Y", 0, 0.05, "25M", 0, 0.05],
        [6, "IRS", "3Y", 0, 0.05, "37M", 0, 0.05]
    ]
    ws_main.range("A5").value = inputs

    table_range = ws_main.range("A4").expand()
    tbl = ws_main.api.ListObjects.Add(1, table_range.api, None, 1)
    tbl.Name = "MarketTable"
    
    ws_main.range("MarketTable[Tenor(Y)]").formula = "=ParseTenor([@[Inst. Tenor]])"
    ws_main.range("MarketTable[Node(Y)]").formula = "=ParseTenor([@[Jump Node]])"
    
    dv_range = ws_main.range("MarketTable[Type]")
    dv_range.api.Validation.Delete()
    dv_range.api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="Deposit,IRS")

    btn = ws_main.api.Buttons().Add(ws_main.range("J4").left, ws_main.range("J4").top, 150, 40)
    btn.Characters.Text = "Run Bootstrapping"
    btn.OnAction = "RunBootstrap"

    # 5. 검증 시트 생성 함수
    def setup_calc_sheet(sheet_name, is_deposit, after_sheet=None):
        ws = wb.sheets.add(sheet_name, after=after_sheet)
        
        # 상단 요약 테이블 구성
        if is_deposit:
            summary_headers = ["Target Tenor", "Target Index", "Target Year", "NPV", "NPV Error"]
        else:
            summary_headers = ["Target Tenor", "Target Index", "Target Year", "NPV fixed", "NPV floating", "NPV Error"]
            
        ws.range("A1").value = summary_headers
        ws.range("A1").expand('right').api.Font.Bold = True
        ws.range("A2").value = "1D" # 초기값
        
        stbl_name = f"Summary_{sheet_name.replace('(', '').replace(')', '')}"
        summary_range = ws.range("A1").expand()
        stbl = ws.api.ListObjects.Add(1, summary_range.api, None, 1)
        stbl.Name = stbl_name
        
        # 요약 테이블 기본 수식 (다른 테이블 참조 전)
        ws.range(f"{stbl_name}[Target Index]").formula = "=IFERROR(MATCH([@[Target Tenor]], MarketTable[Inst. Tenor], 0), 1)"
        ws.range(f"{stbl_name}[Target Year]").formula = "=INDEX(MarketTable[Tenor(Y)], [@[Target Index]])"
        
        # 하단 상세 CF 테이블 구성
        cf_start_row = 5
        ws.range(f"A{cf_start_row}").value = ["시간(연)", "예상현금흐름", "할인계수", "할인현금흐름"]
        ws.range(f"A{cf_start_row}:D{cf_start_row}").api.Font.Bold = True
        
        if is_deposit:
            # INDEX(테이블[컬럼], 1)을 사용하여 행 위치에 상관없이 첫 번째 행의 값을 가져옴
            ws.range(f"A{cf_start_row+1}").formula = f"=INDEX({stbl_name}[Target Year], 1)"
            rows_count = 1
        else:
            rows_count = 80
            times = [[i*0.25] for i in range(1, rows_count + 1)]
            ws.range(f"A{cf_start_row+1}").value = times

        ctbl_name = f"CalcTable_{sheet_name.replace('(', '').replace(')', '')}"
        cf_table_range = ws.range(f"A{cf_start_row}").expand()
        ctbl = ws.api.ListObjects.Add(1, cf_table_range.api, None, 1)
        ctbl.Name = ctbl_name
        
        # CF 테이블 수식 (요약 테이블 참조 시 INDEX 사용)
        ws.range(f"{ctbl_name}[예상현금흐름]").formula = f"=INDEX(MarketTable[Market Rate], INDEX({stbl_name}[Target Index], 1)) * IF({str(is_deposit).upper()}, INDEX({stbl_name}[Target Year], 1), 0.25)"
        ws.range(f"{ctbl_name}[할인계수]").formula = "=LogLinearDF([@[시간(연)]], MarketTable[Node(Y)], MarketTable[Solved Forward])"
        ws.range(f"{ctbl_name}[할인현금흐름]").formula = "=[@예상현금흐름]*[@할인계수]"

        # 요약 테이블 NPV/Error 수식 (CF 테이블 참조)
        if is_deposit:
            ws.range(f"{stbl_name}[NPV]").formula = f"=(1 + INDEX(MarketTable[Market Rate], [@[Target Index]]) * [@[Target Year]]) * LogLinearDF([@[Target Year]], MarketTable[Node(Y)], MarketTable[Solved Forward])"
            ws.range(f"{stbl_name}[NPV Error]").formula = "=[@NPV] - 1"
        else:
            ws.range(f"{stbl_name}[NPV fixed]").formula = f"=SUMIFS({ctbl_name}[할인현금흐름], {ctbl_name}[시간(연)], \"<=\" & [@[Target Year]] + 0.0001)"
            ws.range(f"{stbl_name}[NPV floating]").formula = "=1 - LogLinearDF([@[Target Year]], MarketTable[Node(Y)], MarketTable[Solved Forward])"
            ws.range(f"{stbl_name}[NPV Error]").formula = "=[@[NPV fixed]] - [@[NPV floating]]"
        
        ws.autofit()
        return ws

    ws_depo = setup_calc_sheet("검증(Deposit)", True, after_sheet=ws_main)
    ws_irs = setup_calc_sheet("검증(IRS)", False, after_sheet=ws_depo)

    # 6. Info 시트 구성
    ws_info = wb.sheets.add("Info", after=ws_irs)
    from datetime import datetime
    ws_info.range("A1").value = "Excel Bootstrapper 생성 정보"
    ws_info.range("A1").api.Font.Size = 14
    ws_info.range("A1").api.Font.Bold = True
    
    info_data = [
        ["생성 시점", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["작업 도구", "Python (xlwings) + Excel VBA"],
        ["", ""],
        ["[부트스트래핑 로직 요약]", ""],
        ["1. 보간법", "Log-Linear Interpolation (Linear on Log DF)"],
        ["2. 해찾기", "Newton-Raphson Method 기반 VBA 엔진 (구조적 참조 적용)"],
        ["3. 평가 방식", "Deposit(NPV-1) vs IRS(Fixed-Floating) 시트 분리"],
        ["4. 구조", "모든 계산은 엑셀 테이블(Structured Reference) 기반으로 작동"]
    ]
    ws_info.range("A3").value = info_data
    ws_info.autofit()

    # 7. 저장 및 마무리
    ws_main.activate()
    ws_main.autofit()
    
    save_path = os.path.join(os.getcwd(), "IRS_Bootstrap_Standalone_Final.xlsm")
    try:
        for b in app.books:
            if b.name == os.path.basename(save_path):
                b.close()
        wb.save(save_path)
        print(f"최종 독립형 엑셀 파일이 생성되었습니다: {save_path}")
    except Exception as e:
        print(f"저장 중 오류 발생: {e}")

if __name__ == "__main__":
    create_bootstrap_excel()
