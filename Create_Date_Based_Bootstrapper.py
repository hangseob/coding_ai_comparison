"""
Date-Based IRS Bootstrapper - Excel Standalone Solution
Creates an Excel file with VBA for bootstrapping based on actual dates.
"""

import xlwings as xw
from datetime import datetime, timedelta
import os

def create_date_based_bootstrapper():
    # ========== Configuration ==========
    output_filename = "IRS_Bootstrap_DateBased.xlsm"
    output_path = os.path.join(os.path.dirname(__file__), output_filename)
    
    # ========== VBA Code ==========
    vba_code = '''
' ============================================
' Date-Based IRS Bootstrapping VBA Functions
' ============================================

' Calculate Year Fraction based on Day Count Basis
Public Function CalcYF(startDate As Date, endDate As Date, basis As String) As Double
    Dim days As Long
    Dim startYear As Integer, endYear As Integer
    Dim startMonth As Integer, endMonth As Integer
    Dim startDay As Integer, endDay As Integer
    
    days = endDate - startDate
    
    Select Case LCase(Trim(basis))
        Case "act/365"
            CalcYF = days / 365#
        Case "act/360"
            CalcYF = days / 360#
        Case "act/act"
            ' Simplified: use actual days / actual days in year
            Dim avgYearDays As Double
            avgYearDays = 365.25
            CalcYF = days / avgYearDays
        Case "30/360"
            startYear = Year(startDate)
            startMonth = Month(startDate)
            startDay = Day(startDate)
            endYear = Year(endDate)
            endMonth = Month(endDate)
            endDay = Day(endDate)
            
            If startDay = 31 Then startDay = 30
            If endDay = 31 And startDay = 30 Then endDay = 30
            
            CalcYF = (360 * (endYear - startYear) + 30 * (endMonth - startMonth) + (endDay - startDay)) / 360#
        Case Else
            CalcYF = days / 365#
    End Select
End Function

' Parse tenor string and add to date
Public Function AddTenor(baseDate As Date, tenor As String) As Date
    Dim num As Long
    Dim unit As String
    Dim cleanTenor As String
    
    cleanTenor = UCase(Trim(tenor))
    
    ' Extract number and unit
    num = Val(cleanTenor)
    unit = Right(cleanTenor, 1)
    
    Select Case unit
        Case "D"
            AddTenor = baseDate + num
        Case "W"
            AddTenor = baseDate + num * 7
        Case "M"
            AddTenor = DateAdd("m", num, baseDate)
        Case "Y"
            AddTenor = DateAdd("yyyy", num, baseDate)
        Case Else
            ' Try to parse as number (years)
            If IsNumeric(cleanTenor) Then
                AddTenor = DateAdd("yyyy", CLng(cleanTenor), baseDate)
            Else
                AddTenor = baseDate
            End If
    End Select
End Function

' Get Discount Factor using Log-Linear Interpolation (Date-Based)
Public Function LogLinearDF_Date(targetDate As Date, todayDate As Date, _
                                  jumpDates As Range, solvedFwds As Range, _
                                  basis As String) As Double
    Application.Volatile
    Dim n As Long
    Dim i As Long
    Dim targetYF As Double
    Dim prevYF As Double, currYF As Double
    Dim logDF As Double
    Dim jumpDate As Date
    Dim prevDate As Date
    Dim fwdRate As Double
    Dim lastValidFwd As Double
    
    n = jumpDates.Rows.Count
    targetYF = CalcYF(todayDate, targetDate, basis)
    
    If targetYF <= 0 Then
        LogLinearDF_Date = 1#
        Exit Function
    End If
    
    logDF = 0#
    prevYF = 0#
    lastValidFwd = 0#
    
    For i = 1 To n
        If IsEmpty(jumpDates.Cells(i, 1).Value) Or jumpDates.Cells(i, 1).Value = 0 Then Exit For
        
        jumpDate = jumpDates.Cells(i, 1).Value
        currYF = CalcYF(todayDate, jumpDate, basis)
        fwdRate = solvedFwds.Cells(i, 1).Value
        lastValidFwd = fwdRate
        
        If targetYF <= currYF Then
            ' Target is before or at this jump node
            logDF = logDF - fwdRate * (targetYF - prevYF)
            LogLinearDF_Date = Exp(logDF)
            Exit Function
        Else
            ' Target is after this jump node
            logDF = logDF - fwdRate * (currYF - prevYF)
            prevYF = currYF
        End If
    Next i
    
    ' Target is beyond current jump nodes - extrapolate with last valid rate
    logDF = logDF - lastValidFwd * (targetYF - prevYF)
    LogLinearDF_Date = Exp(logDF)
End Function

' Run Bootstrap Process
Public Sub RunBootstrap()
    Dim wsMain As Worksheet
    Dim tblMarket As ListObject
    Dim tblCommon As ListObject
    Dim todayDate As Date
    Dim basis As String
    Dim i As Long
    Dim n As Long
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set tblMarket = wsMain.ListObjects("MarketTable")
    Set tblCommon = wsMain.ListObjects("Common")
    
    todayDate = tblCommon.ListColumns("Today").DataBodyRange.Cells(1).Value
    basis = tblCommon.ListColumns("DayCount Basis").DataBodyRange.Cells(1).Value
    
    n = tblMarket.ListRows.Count
    
    ' Initialize Solved Forward to 0
    For i = 1 To n
        tblMarket.ListColumns("Solved Forward").DataBodyRange.Cells(i).Value = 0
    Next i
    
    Application.Calculate
    DoEvents
    
    ' Solve each instrument
    For i = 1 To n
        Call SolveStep(i)
        Application.Calculate
        DoEvents
    Next i
    
    MsgBox "Bootstrapping Complete!", vbInformation
End Sub

' Solve for one instrument
Public Sub SolveStep(stepIdx As Long)
    Dim wsMain As Worksheet
    Dim tblMarket As ListObject
    Dim tblCommon As ListObject
    Dim instType As String
    Dim f As Double
    Dim errorVal As Double
    Dim deriv As Double
    Dim epsilon As Double
    Dim maxIter As Long
    Dim iter As Long
    Dim fUp As Double, fDown As Double
    Dim errUp As Double, errDown As Double
    Dim h As Double
    Dim tblSummary As ListObject
    Dim wsSummary As Worksheet
    Dim marketRate As Double
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set tblMarket = wsMain.ListObjects("MarketTable")
    Set tblCommon = wsMain.ListObjects("Common")
    
    instType = LCase(Trim(tblMarket.ListColumns("Type").DataBodyRange.Cells(stepIdx).Value))
    marketRate = tblMarket.ListColumns("Market Rate").DataBodyRange.Cells(stepIdx).Value
    
    ' Select appropriate validation sheet
    If instType = "deposit" Then
        Set wsSummary = ThisWorkbook.Sheets("Validation_Deposit")
        Set tblSummary = wsSummary.ListObjects("Summary_Deposit")
    Else
        Set wsSummary = ThisWorkbook.Sheets("Validation_IRS")
        Set tblSummary = wsSummary.ListObjects("Summary_IRS")
    End If
    
    ' Set target index
    tblSummary.ListColumns("Target Index").DataBodyRange.Cells(1).Value = stepIdx
    
    Application.Calculate
    DoEvents
    
    ' Initial guess
    f = marketRate
    If f = 0 Then f = 0.03
    
    epsilon = 0.0000000001
    maxIter = 100
    h = 0.00001
    
    For iter = 1 To maxIter
        tblMarket.ListColumns("Solved Forward").DataBodyRange.Cells(stepIdx).Value = f
        Application.Calculate
        DoEvents
        
        If IsError(tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value) Then
            errorVal = 1000
        Else
            errorVal = tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value
        End If
        
        If Abs(errorVal) < epsilon Then Exit For
        
        ' Numerical derivative
        fUp = f + h
        tblMarket.ListColumns("Solved Forward").DataBodyRange.Cells(stepIdx).Value = fUp
        Application.Calculate
        DoEvents
        
        If IsError(tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value) Then
            errUp = 1000
        Else
            errUp = tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value
        End If
        
        fDown = f - h
        tblMarket.ListColumns("Solved Forward").DataBodyRange.Cells(stepIdx).Value = fDown
        Application.Calculate
        DoEvents
        
        If IsError(tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value) Then
            errDown = 1000
        Else
            errDown = tblSummary.ListColumns("NPV Error").DataBodyRange.Cells(1).Value
        End If
        
        deriv = (errUp - errDown) / (2 * h)
        
        If Abs(deriv) < 0.0000001 Then
            deriv = 0.0000001
        End If
        
        f = f - errorVal / deriv
        
        ' Bound check
        If f < 0.0001 Then f = 0.0001
        If f > 0.5 Then f = 0.5
    Next iter
    
    tblMarket.ListColumns("Solved Forward").DataBodyRange.Cells(stepIdx).Value = f
    Application.Calculate
End Sub
'''

    # ========== Create Excel Application ==========
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    
    try:
        # Close existing file if open
        for b in app.books:
            if b.name == output_filename:
                b.close()
        
        wb = app.books.add()
        
        # ========== Create Sheets ==========
        ws_main = wb.sheets[0]
        ws_main.name = "Main"
        
        ws_deposit = wb.sheets.add("Validation_Deposit", after=ws_main)
        ws_irs = wb.sheets.add("Validation_IRS", after=ws_deposit)
        ws_info = wb.sheets.add("Info", after=ws_irs)
        
        # ========== Add VBA Module ==========
        vb_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.CodeModule.AddFromString(vba_code)
        
        # ========== Main Sheet Setup ==========
        
        # --- Common Table (Row 1-2) ---
        ws_main.range("A1").value = [["Today", "DayCount Basis", "IRS Coupon Freq"]]
        ws_main.range("A2").value = [[datetime(2026, 1, 8), "ACT/365", 4]]
        
        common_range = ws_main.range("A1:C2")
        common_table = ws_main.api.ListObjects.Add(1, common_range.api, None, 1)  # xlSrcRange=1, xlYes=1
        common_table.Name = "Common"
        
        # Data Validation for DayCount Basis
        try:
            cell_b2 = ws_main.range("B2").api
            cell_b2.Validation.Delete()
            cell_b2.Validation.Add(Type=3, AlertStyle=1, Formula1="ACT/365,ACT/360,ACT/ACT,30/360")
        except:
            pass
        
        # Data Validation for IRS Coupon Freq
        try:
            cell_c2 = ws_main.range("C2").api
            cell_c2.Validation.Delete()
            cell_c2.Validation.Add(Type=3, AlertStyle=1, Formula1="1,2,4,12,Daily")
        except:
            pass
        
        # --- JumpDates Table (Column O-P, to avoid overlap with MarketTable) ---
        ws_main.range("O1").value = [["No", "Jump Date"]]
        
        # Generate jump dates from provided list
        jump_dates_data = [
            [1, datetime(2026, 1, 15)],
            [2, datetime(2026, 2, 26)],
            [3, datetime(2026, 4, 10)],
            [4, datetime(2026, 5, 28)],
            [5, datetime(2026, 7, 16)],
            [6, datetime(2026, 8, 27)],
            [7, datetime(2026, 10, 22)],
            [8, datetime(2026, 11, 26)],
            [9, datetime(2027, 1, 14)],
            [10, datetime(2027, 2, 25)],
            [11, datetime(2027, 4, 15)],
            [12, datetime(2027, 5, 27)],
            [13, datetime(2027, 7, 15)],
            [14, datetime(2027, 8, 26)],
            [15, datetime(2027, 10, 21)],
            [16, datetime(2027, 11, 25)],
            [17, datetime(2028, 1, 15)]
        ]
        
        ws_main.range("O2").value = jump_dates_data
        
        jump_end_row = 1 + len(jump_dates_data)
        jump_range = ws_main.range(f"O1:P{jump_end_row}")
        jump_table = ws_main.api.ListObjects.Add(1, jump_range.api, None, 1)
        jump_table.Name = "JumpDates"
        
        # Format Jump Date column as date
        ws_main.range(f"P2:P{jump_end_row}").number_format = "YYYY-MM-DD"
        
        # --- MarketTable (Row 5 onwards) ---
        market_headers = ["No", "Inst. Tenor", "Type", "Mty Date", "Mty YearFrac", 
                          "Jump Date", "Jump YearFrac", "Market Rate", "Solved Forward",
                          "Jump Date DCF", "Mty Date DCF", "Jump Zero Rate", "Mty Zero Rate"]
        ws_main.range("A5").value = [market_headers]
        
        # Market data from provided list
        market_data = [
            [1, "1D", "Deposit", "", "", "", "", 0.02500, 0, 0, 0, 0, 0],
            [2, "3M", "Deposit", "", "", "", "", 0.02700, 0, 0, 0, 0, 0],
            [3, "06M", "IRS", "", "", "", "", 0.02705, 0, 0, 0, 0, 0],
            [4, "09M", "IRS", "", "", "", "", 0.027125, 0, 0, 0, 0, 0],
            [5, "01Y", "IRS", "", "", "", "", 0.02735, 0, 0, 0, 0, 0],
            [6, "18M", "IRS", "", "", "", "", 0.028025, 0, 0, 0, 0, 0],
            [7, "02Y", "IRS", "", "", "", "", 0.028925, 0, 0, 0, 0, 0],
        ]
        
        ws_main.range("A6").value = market_data
        
        market_end_row = 5 + len(market_data)
        market_range = ws_main.range(f"A5:M{market_end_row}")
        market_table = ws_main.api.ListObjects.Add(1, market_range.api, None, 1)
        market_table.Name = "MarketTable"
        
        # Data Validation for Type column
        try:
            type_col_range = ws_main.range(f"C6:C{market_end_row}").api
            type_col_range.Validation.Delete()
            type_col_range.Validation.Add(Type=3, AlertStyle=1, Formula1="Deposit,IRS")
        except:
            pass
        
        # Formulas for MarketTable
        # Mty Date = AddTenor(Common[Today], [@[Inst. Tenor]])
        for row in range(6, market_end_row + 1):
            ws_main.range(f"D{row}").formula = '=AddTenor(INDEX(Common[Today],1), [@[Inst. Tenor]])'
            ws_main.range(f"E{row}").formula = '=CalcYF(INDEX(Common[Today],1), [@[Mty Date]], INDEX(Common[DayCount Basis],1))'
            # Jump Date = smallest date in JumpDates >= Mty Date
            ws_main.range(f"F{row}").formula = '=MINIFS(JumpDates[Jump Date], JumpDates[Jump Date], ">=" & [@[Mty Date]])'
            ws_main.range(f"G{row}").formula = '=CalcYF(INDEX(Common[Today],1), [@[Jump Date]], INDEX(Common[DayCount Basis],1))'
            # New Columns Formulas
            ws_main.range(f"J{row}").formula = '=LogLinearDF_Date([@[Jump Date]], INDEX(Common[Today],1), MarketTable[Jump Date], MarketTable[Solved Forward], INDEX(Common[DayCount Basis],1))'
            ws_main.range(f"K{row}").formula = '=LogLinearDF_Date([@[Mty Date]], INDEX(Common[Today],1), MarketTable[Jump Date], MarketTable[Solved Forward], INDEX(Common[DayCount Basis],1))'
            ws_main.range(f"L{row}").formula = '=IF([@[Jump YearFrac]]>0, -LN([@[Jump Date DCF]])/[@[Jump YearFrac]], 0)'
            ws_main.range(f"M{row}").formula = '=IF([@[Mty YearFrac]]>0, -LN([@[Mty Date DCF]])/[@[Mty YearFrac]], 0)'
        
        # Format date columns
        ws_main.range(f"D6:D{market_end_row}").number_format = "YYYY-MM-DD"
        ws_main.range(f"F6:F{market_end_row}").number_format = "YYYY-MM-DD"
        
        # Format rate columns as percentage
        ws_main.range(f"H6:I{market_end_row}").number_format = "0.0000%"
        ws_main.range(f"L6:M{market_end_row}").number_format = "0.0000%"
        # Format DCF columns
        ws_main.range(f"J6:K{market_end_row}").number_format = "0.00000000"
        
        # Add Run Button
        btn = ws_main.api.Buttons().Add(400, 5, 120, 30)
        btn.OnAction = "RunBootstrap"
        btn.Caption = "Run Bootstrapping"
        
        # ========== 검증(Deposit) Sheet ==========
        # Summary table (horizontal)
        ws_deposit.range("A1").value = [["Target Tenor", "Target Index", "Target Date", "Target YearFrac", "Market Rate", "NPV", "NPV Error"]]
        ws_deposit.range("A2").value = [["1D", 1, "", "", "", "", ""]]
        
        summary_dep_range = ws_deposit.range("A1:G2")
        summary_dep_table = ws_deposit.api.ListObjects.Add(1, summary_dep_range.api, None, 1)
        summary_dep_table.Name = "Summary_Deposit"
        
        # Formulas for Summary_Deposit
        ws_deposit.range("A2").formula = '=INDEX(MarketTable[Inst. Tenor], [@[Target Index]])'
        ws_deposit.range("C2").formula = '=INDEX(MarketTable[Mty Date], [@[Target Index]])'
        ws_deposit.range("D2").formula = '=INDEX(MarketTable[Mty YearFrac], [@[Target Index]])'
        ws_deposit.range("E2").formula = '=INDEX(MarketTable[Market Rate], [@[Target Index]])'
        # NPV = (1 + rate * yearfrac) * DF
        ws_deposit.range("F2").formula = '=(1 + [@[Market Rate]] * [@[Target YearFrac]]) * LogLinearDF_Date([@[Target Date]], INDEX(Common[Today],1), MarketTable[Jump Date], MarketTable[Solved Forward], INDEX(Common[DayCount Basis],1))'
        ws_deposit.range("G2").formula = '=[@NPV] - 1'
        
        ws_deposit.range("C2").number_format = "YYYY-MM-DD"
        ws_deposit.range("E2").number_format = "0.0000%"
        
        # ========== 검증(IRS) Sheet ==========
        # First create CalcTable_IRS (needs to exist before Summary_IRS references it)
        calc_headers = ["No", "Payment Date", "YearFrac", "Expected CashFlow", "DF", "DCF"]
        ws_irs.range("A5").value = [calc_headers]
        
        calc_data = []
        today = datetime(2026, 1, 8)
        for i in range(1, 41):  # 40 quarters = 10 years
            payment_date = today
            months_to_add = i * 3
            years_add = months_to_add // 12
            months_add = months_to_add % 12
            new_month = today.month + months_add
            if new_month > 12:
                years_add += 1
                new_month -= 12
            payment_date = today.replace(year=today.year + years_add, month=new_month)
            calc_data.append([i, payment_date, 0, 0, 1, 0])  # Use placeholder values
        
        ws_irs.range("A6").value = calc_data
        
        calc_end_row = 5 + len(calc_data)
        calc_range = ws_irs.range(f"A5:F{calc_end_row}")
        calc_table = ws_irs.api.ListObjects.Add(1, calc_range.api, None, 1)
        calc_table.Name = "CalcTable_IRS"
        ws_irs.range(f"B6:B{calc_end_row}").number_format = "YYYY-MM-DD"
        
        # Now create Summary table (horizontal)
        ws_irs.range("A1").value = [["Target Tenor", "Target Index", "Target Date", "Target YearFrac", "Market Rate", "NPV Fixed", "NPV Floating", "NPV Error"]]
        ws_irs.range("A2").value = [["1Y", 6, today, 1.0, 0.0375, 0, 0, 0]]  # Use placeholder values
        
        summary_irs_range = ws_irs.range("A1:H2")
        summary_irs_table = ws_irs.api.ListObjects.Add(1, summary_irs_range.api, None, 1)
        summary_irs_table.Name = "Summary_IRS"
        
        ws_irs.range("C2").number_format = "YYYY-MM-DD"
        ws_irs.range("E2").number_format = "0.0000%"
        
        # Now apply formulas after both tables exist
        # Formulas for Summary_IRS
        ws_irs.range("A2").formula = '=INDEX(MarketTable[Inst. Tenor], [@[Target Index]])'
        ws_irs.range("C2").formula = '=INDEX(MarketTable[Mty Date], [@[Target Index]])'
        ws_irs.range("D2").formula = '=INDEX(MarketTable[Mty YearFrac], [@[Target Index]])'
        ws_irs.range("E2").formula = '=INDEX(MarketTable[Market Rate], [@[Target Index]])'
        ws_irs.range("F2").formula = '=SUMIFS(CalcTable_IRS[DCF], CalcTable_IRS[Payment Date], "<=" & [@[Target Date]] + 0.5)'
        ws_irs.range("G2").formula = '=1 - LogLinearDF_Date([@[Target Date]], INDEX(Common[Today],1), MarketTable[Jump Date], MarketTable[Solved Forward], INDEX(Common[DayCount Basis],1))'
        ws_irs.range("H2").formula = '=[@[NPV Fixed]] - [@[NPV Floating]]'
        
        # Formulas for CalcTable_IRS
        for row in range(6, calc_end_row + 1):
            ws_irs.range(f"C{row}").formula = '=CalcYF(INDEX(Common[Today],1), [@[Payment Date]], INDEX(Common[DayCount Basis],1))'
            # Expected CashFlow: Market Rate * YearFraction(Accrual Period)
            # No=1이면 Today를 시작일로, 그 외에는 이전 행의 Payment Date를 시작일로 사용
            ws_irs.range(f"D{row}").formula = '=INDEX(Summary_IRS[Market Rate],1) * CalcYF(IF([@No]=1, INDEX(Common[Today],1), OFFSET([@[Payment Date]],-1,0)), [@[Payment Date]], INDEX(Common[DayCount Basis],1))'
            ws_irs.range(f"E{row}").formula = '=LogLinearDF_Date([@[Payment Date]], INDEX(Common[Today],1), MarketTable[Jump Date], MarketTable[Solved Forward], INDEX(Common[DayCount Basis],1))'
            ws_irs.range(f"F{row}").formula = '=[@[Expected CashFlow]] * [@DF]'
        
        # ========== Info Sheet ==========
        creation_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            ws_info.range("A1").value = "Creation Timestamp"
            ws_info.range("B1").value = creation_time
            ws_info.range("A3").value = "Date-Based IRS Bootstrapper"
            ws_info.range("A5").value = "Features:"
            ws_info.range("A6").value = "- All calculations are date-based"
            ws_info.range("A7").value = "- Day count: ACT/365, ACT/360, ACT/ACT, 30/360"
            ws_info.range("A8").value = "- IRS freq: Annual, Semi, Quarterly, Monthly, Daily"
            ws_info.range("A10").value = "VBA Functions:"
            ws_info.range("A11").value = "CalcYF, AddTenor, LogLinearDF_Date, RunBootstrap"
            ws_info.range("A13").value = "Usage:"
            ws_info.range("A14").value = "1. Set Today and DayCount Basis"
            ws_info.range("A15").value = "2. Enter market instruments"
            ws_info.range("A16").value = "3. Click Run Bootstrapping button"
        except Exception as e:
            print(f"Warning: Could not write Info sheet: {e}")
        
        # ========== Final Adjustments ==========
        # Auto-fit columns
        ws_main.autofit()
        ws_deposit.autofit()
        ws_irs.autofit()
        ws_info.autofit()
        
        # Save workbook
        wb.save(output_path)
        print(f"Successfully created: {output_path}")
        
        # Force calculation
        app.api.Calculate()
        
        return wb, app
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        app.quit()
        raise

if __name__ == "__main__":
    wb, app = create_date_based_bootstrapper()
    print("Excel file created. Please test the bootstrapping manually.")

