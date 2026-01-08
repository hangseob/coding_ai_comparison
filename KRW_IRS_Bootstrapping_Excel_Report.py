import numpy as np
import pandas as pd
from scipy.optimize import fsolve
import os

try:
    import xlwings as xw
except ImportError:
    print("xlwings 라이브러리가 설치되어 있지 않습니다. 'pip install xlwings'를 실행해 주세요.")
    xw = None

# 1. 환경 설정 및 시장 데이터
maturities = {
    '1D Call': 1/365,
    '3M Depo': 3/12,
    '6M IRS': 6/12,
    '9M IRS': 9/12,
    '1Y IRS': 12/12
}

market_rates = {
    '1D Call': 0.0350,
    '3M Depo': 0.0360,
    '6M IRS': 0.0370,
    '9M IRS': 0.0380,
    '1Y IRS': 0.0390
}

# Jumps (Knots) 위치: 2M, 5M, 7M, 10M, 13M
nodes = np.array([2, 5, 7, 10, 13]) / 12.0

def get_df(t, nodes, forwards):
    """Python 내부 계산용 DF 함수"""
    if t <= 0: return 1.0
    integral = 0
    prev_node = 0
    for i in range(len(nodes)):
        node = nodes[i]
        f = forwards[i]
        if t <= node:
            integral += f * (t - prev_node)
            return np.exp(-integral)
        else:
            integral += f * (node - prev_node)
            prev_node = node
    integral += forwards[-1] * (t - prev_node)
    return np.exp(-integral)

# 2. 부트스트래핑 수행
solved_forwards = []

# 1) 1D Call
f1 = fsolve(lambda f: get_df(maturities['1D Call'], nodes[:1], [f]) - 1/(1 + market_rates['1D Call'] * maturities['1D Call']), market_rates['1D Call'])[0]
solved_forwards.append(f1)

# 2) 3M Depo
f2 = fsolve(lambda f: get_df(maturities['3M Depo'], nodes[:2], solved_forwards + [f]) - 1/(1 + market_rates['3M Depo'] * maturities['3M Depo']), market_rates['3M Depo'])[0]
solved_forwards.append(f2)

# 3) 6M IRS
def obj_6m(f):
    fwds = solved_forwards + [f]
    dfs = [get_df(t, nodes[:3], fwds) for t in [3/12, 6/12]]
    return market_rates['6M IRS'] * 0.25 * sum(dfs) - (1 - dfs[-1])
f3 = fsolve(obj_6m, market_rates['6M IRS'])[0]
solved_forwards.append(f3)

# 4) 9M IRS
def obj_9m(f):
    fwds = solved_forwards + [f]
    dfs = [get_df(t, nodes[:4], fwds) for t in [3/12, 6/12, 9/12]]
    return market_rates['9M IRS'] * 0.25 * sum(dfs) - (1 - dfs[-1])
f4 = fsolve(obj_9m, market_rates['9M IRS'])[0]
solved_forwards.append(f4)

# 5) 1Y IRS
def obj_1y(f):
    fwds = solved_forwards + [f]
    dfs = [get_df(t, nodes[:5], fwds) for t in [3/12, 6/12, 9/12, 12/12]]
    return market_rates['1Y IRS'] * 0.25 * sum(dfs) - (1 - dfs[-1])
f5 = fsolve(obj_1y, market_rates['1Y IRS'])[0]
solved_forwards.append(f5)

# 3. 엑셀 출력 (xlwings)
def generate_excel_report():
    if not xw: return
    
    # 엑셀 앱 실행
    app = xw.App(visible=True)
    wb = app.books.add()
    
    # VBA 매크로 추가 (xlsm으로 저장해야 함)
    vba_code = """
Function xlLinearInt(xRange As Range, yRange As Range, x As Double) As Variant
    Dim i As Long
    Dim n As Long
    n = xRange.Cells.Count
    If n <> yRange.Cells.Count Then
        xlLinearInt = CVErr(xlErrRef)
        Exit Function
    End If
    If x <= xRange(1).Value Then
        xlLinearInt = yRange(1).Value
        Exit Function
    End If
    If x >= xRange(n).Value Then
        xlLinearInt = yRange(n).Value
        Exit Function
    End If
    For i = 1 To n - 1
        If x >= xRange(i).Value And x <= xRange(i + 1).Value Then
            xlLinearInt = yRange(i).Value + (yRange(i + 1).Value - yRange(i).Value) * _
                          (x - xRange(i).Value) / (xRange(i + 1).Value - xRange(i).Value)
            Exit Function
        End If
    Next i
    xlLinearInt = CVErr(xlErrValue)
End Function
    """
    
    # VBA 모듈 주입 (보안 설정에 따라 실패할 수 있음)
    try:
        vba_module = wb.api.VBProject.VBComponents.Add(1)
        vba_module.CodeModule.AddFromString(vba_code)
    except Exception as e:
        print("VBA 매크로 주입 실패: 엑셀 설정에서 'VBA 프로젝트 개체 모델에 안전하게 액세스'가 체크되어 있어야 합니다.")
        print("수동으로 VBE(Alt+F11)에 모듈을 추가해 주세요.")

    # --- 시트 1: 요약_및_계산로직 ---
    sheet1 = wb.sheets[0]
    sheet1.name = "요약_및_계산로직"
    sheet1.range("A1").value = "KRW IRS Bootstrapping 요약 (Log-Linear Interpolation)"
    
    # 데이터 준비 (Knot 0 시점 추가)
    summary_data = []
    # Knot 0 추가
    summary_data.append(["T=0", 0.0, 0.0, 0.0, 0.0, 1.0, 0.0])
    
    labels = list(maturities.keys())
    for i in range(len(labels)):
        name = labels[i]
        t_knot = nodes[i]
        f_val = solved_forwards[i]
        df_val = get_df(t_knot, nodes, solved_forwards)
        log_df = np.log(df_val)
        summary_data.append([name, maturities[name], market_rates[name], [2, 5, 7, 10, 13][i], t_knot, df_val, log_df])
    
    cols = ["인스트루먼트", "만기(년)", "시장금리", "Knot(개월)", "Knot(년)", "DiscountFactor", "LogDF"]
    df_summary = pd.DataFrame(summary_data, columns=cols)
    
    sheet1.range("A4").value = df_summary
    
    # 엑셀 테이블로 변환
    table_range = sheet1.range("A4").expand()
    tbl = sheet1.api.ListObjects.Add(1, table_range.api, None, 1)
    tbl.Name = "SummaryTable"
    
    # --- 시트 2: 상세_캐쉬플로우 ---
    sheet2 = wb.sheets.add("상세_캐쉬플로우", after=sheet1)
    sheet2.range("A1").value = "상세 캐쉬플로우 (xlLinearInt 및 테이블 참조)"
    
    current_row = 3
    for name, t_end in maturities.items():
        sheet2.range(f"A{current_row}").value = f"[{name}] 상세 내역"
        sheet2.range(f"A{current_row}").api.Font.Bold = True
        current_row += 1
        
        headers = ["No", "현금흐름일(년)", "현금흐름액", "할인계수(DF)", "할인현금흐름(DCF)"]
        sheet2.range(f"A{current_row}").value = headers
        sheet2.range(f"A{current_row}:E{current_row}").color = (230, 230, 230)
        start_data_row = current_row + 1
        
        pay_times = [maturities[name]] if name in ['1D Call', '3M Depo'] else np.arange(0.25, t_end + 1e-6, 0.25)
        
        for i, t in enumerate(pay_times):
            row = start_data_row + i
            sheet2.range(f"A{row}").value = i + 1
            sheet2.range(f"B{row}").value = t
            
            if name in ['1D Call', '3M Depo']:
                sheet2.range(f"C{row}").value = 1 + market_rates[name] * t
            else:
                sheet2.range(f"C{row}").value = market_rates[name] * 0.25
            
            # VBA 함수와 테이블 이름을 사용한 가독성 높은 수식
            # log(DF)를 선형보간한 후 EXP를 취함
            formula_df = f"=EXP(xlLinearInt(SummaryTable[Knot(년)], SummaryTable[LogDF], B{row}))"
            sheet2.range(f"D{row}").formula = formula_df
            sheet2.range(f"E{row}").formula = f"=C{row}*D{row}"
            
        current_row += len(pay_times) + 1
        
        if 'IRS' in name:
            last_row = current_row - 2
            sheet2.range(f"D{current_row}").value = "Fixed Leg PV:"
            sheet2.range(f"E{current_row}").formula = f"=SUM(E{start_data_row}:E{last_row})"
            current_row += 1
            sheet2.range(f"D{current_row}").value = "Floating Leg PV (1-DF_end):"
            sheet2.range(f"E{current_row}").formula = f"=1-D{last_row}"
            current_row += 2
        else:
            sheet2.range(f"D{current_row}").value = "Total PV:"
            sheet2.range(f"E{current_row}").formula = f"=E{start_data_row}"
            current_row += 2

    sheet1.autofit(axis='columns')
    sheet2.autofit(axis='columns')
    
    # 매크로 포함 파일(.xlsm)로 저장
    save_path = os.path.join(os.getcwd(), "IRS_Bootstrapping_Readable.xlsm")
    try:
        wb.save(save_path)
        print(f"리포트가 생성되었습니다: {save_path}")
    except Exception as e:
        # xlsm 저장 실패 시 (보안 문제 등) 일반 xlsx로 저장 시도
        save_path_xlsx = save_path.replace(".xlsm", ".xlsx")
        wb.save(save_path_xlsx)
        print(f"매크로 제외 버전으로 저장되었습니다: {save_path_xlsx}")

if __name__ == "__main__":
    generate_excel_report()
