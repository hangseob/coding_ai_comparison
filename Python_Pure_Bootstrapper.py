import xlwings as xw
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
from scipy.optimize import newton
import os

class HybridReporter:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = None
        self.today = None
        self.basis = "ACT/365"
        self.freq = 4
        self.market_data = None
        
    def load_data(self):
        print(f"데이터 로드 중: {self.file_path}")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True, add_book=False)
        self.wb = app.books.open(self.file_path)
        ws_main = self.wb.sheets["Main"]
        
        # 설정 로드
        tbl_common = ws_main.api.ListObjects("Common")
        df_common = ws_main.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_common.columns = [str(c).strip() for c in df_common.columns]
        self.today = pd.to_datetime(df_common['Today'].iloc[0])
        self.basis = str(df_common['DayCount Basis'].iloc[0]).upper()
        self.freq = int(df_common['IRS Coupon Freq'].iloc[0])
        
        # 시장 데이터 로드
        tbl_market = ws_main.api.ListObjects("MarketTable")
        self.market_data = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
        self.market_data.columns = [str(c).strip() for c in self.market_data.columns]
        
        # JumpDates 테이블 로드
        tbl_jumpdates = ws_main.api.ListObjects("JumpDates")
        jump_dates_df = ws_main.range(tbl_jumpdates.Range.Address).options(pd.DataFrame, index=False, header=True).value
        jump_dates_df.columns = [str(c).strip() for c in jump_dates_df.columns]
        self.jump_dates_list = sorted(pd.to_datetime(jump_dates_df['Jump Date']).tolist())
        
        # Python에서 Today + Tenor로 Mty Date 재계산
        self.market_data['Mty Date'] = self.market_data['Inst. Tenor'].apply(
            lambda t: self.calc_mty_date(self.today, t))
        
        # Jump Date = JumpDates 중 Mty Date 이상인 날짜 중 최소값
        def find_jump_date(mty_date):
            for jd in self.jump_dates_list:
                if jd >= mty_date:
                    return jd
            return self.jump_dates_list[-1]
        
        self.market_data['Jump Date'] = self.market_data['Mty Date'].apply(find_jump_date)

    def year_frac(self, start, end):
        days = (end - start).days
        return days / 365.0 # Simplified for internal solver

    def get_df_internal(self, target_date, jump_dates, forward_rates):
        target_yf = self.year_frac(self.today, target_date)
        if target_yf <= 0: return 1.0
        log_df, prev_yf = 0.0, 0.0
        for i, jump_date in enumerate(jump_dates):
            curr_yf = self.year_frac(self.today, jump_date)
            if target_yf <= curr_yf:
                log_df -= forward_rates[i] * (target_yf - prev_yf)
                return np.exp(log_df)
            log_df -= forward_rates[i] * (curr_yf - prev_yf)
            prev_yf = curr_yf
        log_df -= forward_rates[-1] * (target_yf - prev_yf)
        return np.exp(log_df)

    def parse_tenor(self, tenor_str):
        s = str(tenor_str).upper()
        num = float(''.join(filter(lambda x: x.isdigit() or x=='.', s)))
        if 'M' in s: return num / 12.0
        return num
    
    def calc_mty_date(self, today, tenor_str):
        """Today + Tenor로 만기일 계산"""
        s = str(tenor_str).upper().strip()
        num = int(''.join(filter(str.isdigit, s)))
        
        if 'W' in s:
            return today + timedelta(weeks=num)
        elif 'M' in s:
            new_month = today.month + num
            new_year = today.year + (new_month - 1) // 12
            new_month = (new_month - 1) % 12 + 1
            try:
                return today.replace(year=new_year, month=new_month)
            except ValueError:
                import calendar
                last_day = calendar.monthrange(new_year, new_month)[1]
                return today.replace(year=new_year, month=new_month, day=min(today.day, last_day))
        elif 'Y' in s:
            try:
                return today.replace(year=today.year + num)
            except ValueError:
                return today.replace(year=today.year + num, day=28)
        else:
            return today + timedelta(days=num)

    def npv_error_internal(self, fwd_guess, step_idx, solved_fwds):
        row = self.market_data.iloc[step_idx]
        current_fwds = solved_fwds.copy()
        current_fwds[step_idx] = fwd_guess
        jump_dates = self.market_data['Jump Date'].iloc[:step_idx+1].tolist()
        
        mkt_rate, mty_date = row['Market Rate'], row['Mty Date']
        if str(row['Type']).lower() == "deposit":
            df = self.get_df_internal(mty_date, jump_dates, current_fwds)
            return (1 + mkt_rate * self.year_frac(self.today, mty_date)) * df - 1.0
        else:
            tenor_yf = self.parse_tenor(row['Inst. Tenor'])
            num_coupons = int(round(tenor_yf * self.freq))
            fixed_pv = 0.0
            p_date = self.today
            for j in range(1, num_coupons + 1):
                c_date = mty_date if j == num_coupons else self.today + timedelta(days=int(j * (365/self.freq)))
                yf = self.year_frac(p_date, c_date)
                fixed_pv += mkt_rate * yf * self.get_df_internal(c_date, jump_dates, current_fwds)
                p_date = c_date
            return fixed_pv - (1.0 - self.get_df_internal(mty_date, jump_dates, current_fwds))

    def run_bootstrap(self):
        print("파이썬 내부 부트스트랩 계산 중...")
        n = len(self.market_data)
        self.solved_fwds = np.zeros(n)
        for i in range(n):
            self.solved_fwds[i] = newton(self.npv_error_internal, self.market_data.iloc[i]['Market Rate'], args=(i, self.solved_fwds), tol=1e-12)
        
        # 엑셀 메인 테이블 업데이트 (Mty Date, Jump Date, Solved Forward)
        ws_main = self.wb.sheets["Main"]
        tbl_market = ws_main.api.ListObjects("MarketTable")
        ws_main.range(tbl_market.ListColumns("Mty Date").DataBodyRange.Address).value = [[d] for d in self.market_data['Mty Date']]
        ws_main.range(tbl_market.ListColumns("Jump Date").DataBodyRange.Address).value = [[d] for d in self.market_data['Jump Date']]
        ws_main.range(tbl_market.ListColumns("Solved Forward").DataBodyRange.Address).value = self.solved_fwds.reshape(-1, 1)
        self.wb.app.calculate()
        print("계산 완료 및 메인 테이블 업데이트 성공.")

    def write_validation_sheets(self):
        print("검증 시트 작성 중 (Excel 수식 적용)...")
        
        # Main 시트의 핵심 셀 주소 파악 (A2: Today, B2: Basis 가정이나 테이블에서 동적 추출)
        ws_main = self.wb.sheets["Main"]
        tbl_common = ws_main.api.ListObjects("Common")
        today_addr = f"Main!{ws_main.range(tbl_common.DataBodyRange.Cells(1, 1).Address).address}"
        basis_addr = f"Main!{ws_main.range(tbl_common.DataBodyRange.Cells(1, 2).Address).address}"
        
        for name in ["Validation_Deposit", "Validation_IRS"]:
            if name not in [s.name for s in self.wb.sheets]:
                self.wb.sheets.add(name, after=self.wb.sheets["Main"])
            ws = self.wb.sheets[name]
            ws.clear()
            
            curr_row = 1
            inst_type_filter = "deposit" if "Deposit" in name else "irs"
            
            for i, row in self.market_data.iterrows():
                if str(row['Type']).lower() != inst_type_filter: continue
                
                # 테너별 리포트 블록 작성
                tenor = row['Inst. Tenor']
                mkt_rate = row['Market Rate']
                mty_date = row['Mty Date']
                
                ws.range(f"A{curr_row}").value = f"Tenor: {tenor} | Market Rate: {mkt_rate:.4%}"
                ws.range(f"A{curr_row}").font.bold = True
                curr_row += 1
                
                headers = ["Date", "Cpn YF", "DF", "CF Amount", "DCF"]
                ws.range(f"A{curr_row}").value = headers
                ws.range(f"A{curr_row}:E{curr_row}").color = (200, 200, 200)
                start_data_row = curr_row + 1
                
                # 현금흐름 날짜 생성
                cf_dates = []
                if inst_type_filter == "deposit":
                    cf_dates = [mty_date]
                else:
                    num_coupons = int(round(self.parse_tenor(tenor) * self.freq))
                    for j in range(1, num_coupons + 1):
                        cf_dates.append(mty_date if j == num_coupons else self.today + timedelta(days=int(j * (365/self.freq))))
                
                # 데이터 입력 및 수식 설정
                for j, d in enumerate(cf_dates):
                    r = start_data_row + j
                    ws.range(f"A{r}").value = d
                    
                    # 1. Cpn YF 계산 (IRS는 직전일 기준, Deposit은 Today 기준)
                    if inst_type_filter == "irs":
                        prev_date_ref = today_addr if j == 0 else f"A{r-1}"
                        ws.range(f"B{r}").formula = f"=CalcYF({prev_date_ref}, A{r}, {basis_addr})"
                    else:
                        ws.range(f"B{r}").formula = f"=CalcYF({today_addr}, A{r}, {basis_addr})"
                    
                    # 2. DF 수식 (항상 Today 기준으로 할인)
                    ws.range(f"C{r}").formula = f"=LogLinearDF_Date(A{r}, {today_addr}, MarketTable[Jump Date], MarketTable[Solved Forward], {basis_addr})"
                    
                    # 3. CF Amount 및 DCF
                    if inst_type_filter == "deposit":
                        ws.range(f"D{r}").formula = f"=1 + {mkt_rate} * B{r}"
                    else:
                        ws.range(f"D{r}").formula = f"={mkt_rate} * B{r}"
                    ws.range(f"E{r}").formula = f"=D{r} * C{r}"
                
                end_data_row = start_data_row + len(cf_dates) - 1
                
                # NPV Error 수식
                ws.range(f"G{start_data_row}").value = "NPV Error"
                if inst_type_filter == "deposit":
                    ws.range(f"G{start_data_row+1}").formula = f"=SUM(E{start_data_row}:E{end_data_row}) - 1"
                else:
                    ws.range(f"G{start_data_row+1}").formula = f"=SUM(E{start_data_row}:E{end_data_row}) - (1 - C{end_data_row})"
                
                curr_row = end_data_row + 3 # 다음 블록을 위해 간격 띄움

        print("검증 시트 모든 테너 리포트 작성 완료.")
        self.wb.save()

    def plot_results(self):
        print("결과 차트 생성 중...")
        df_plot = self.market_data.copy()
        
        fig = go.Figure()

        # 1. Market Rate Points
        fig.add_trace(go.Scatter(
            x=df_plot['Mty Date'], y=df_plot['Market Rate'],
            mode='markers+text', name='Market Rate',
            text=[f"{r:.2%}" for r in df_plot['Market Rate']],
            textposition="top center",
            marker=dict(size=10, color='gray'),
            hovertemplate="<b>Market Rate</b><br>Date: %{x|%Y-%m-%d}<br>Rate: %{y:.4%}<extra></extra>"
        ))

        # 2. Forward Curve (Step Line with Daily Sampling)
        step_x, step_y, step_text = [], [], []
        p_date = self.today
        for _, row in df_plot.iterrows():
            c_date = row['Jump Date']
            f_rate = self.solved_fwds[df_plot.index[df_plot['Jump Date'] == c_date][0]]
            
            # 일 단위 샘플링으로 선 전체 호버 가능하게 함
            num_days = (c_date - p_date).days
            period_str = f"<b>{p_date.strftime('%Y-%m-%d')} ~ {c_date.strftime('%Y-%m-%d')}</b>"
            
            for d in range(num_days + 1):
                current_d = p_date + timedelta(days=d)
                step_x.append(current_d)
                step_y.append(f_rate)
                step_text.append(period_str)
            
            step_x.append(None); step_y.append(None); step_text.append(None)
            p_date = c_date

        fig.add_trace(go.Scatter(
            x=step_x, y=step_y,
            mode='lines', name='Solved Forward',
            line=dict(color='red', width=3),
            customdata=step_text,
            hovertemplate="<b>Forward Rate</b><br>Rate: %{y:.4%}<br>Period: %{customdata}<extra></extra>"
        ))

        # Layout 설정
        fig.update_layout(
            title=dict(
                text=f"IRS Forward Curve Analysis ({self.today.strftime('%Y-%m-%d')})",
                x=0.5, font=dict(size=22)
            ),
            xaxis=dict(title="Timeline", type='date', tickformat='%Y-%m-%d', tickangle=-45, gridcolor='#eee'),
            yaxis=dict(title="Rate (%)", tickformat=".2%", gridcolor='#eee'),
            template="plotly_white", width=1200, height=700,
            hovermode="closest"
        )

        fig.show()
        print("차트가 브라우저에 표시되었습니다.")

if __name__ == "__main__":
    runner = HybridReporter("IRS_Bootstrap_DateBased.xlsm")
    runner.load_data()
    runner.run_bootstrap()
    runner.write_validation_sheets()
    runner.plot_results()