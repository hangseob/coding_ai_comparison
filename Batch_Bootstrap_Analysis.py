import xlwings as xw
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
from scipy.optimize import fsolve
import os

class BatchBootstrapper:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = None
        self.app = None
        self.basis = "ACT/365"
        self.freq = 4
        
    def year_frac(self, start, end):
        return (end - start).days / 365.0

    def get_df_internal(self, target_date, today, jump_dates, forward_rates):
        target_yf = self.year_frac(today, target_date)
        if target_yf <= 0: return 1.0
        log_df, prev_yf = 0.0, 0.0
        for i, jump_date in enumerate(jump_dates):
            curr_yf = self.year_frac(today, jump_date)
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
            # 월 단위: 대략적인 계산 (30일 기준이 아닌 실제 월 계산)
            new_month = today.month + num
            new_year = today.year + (new_month - 1) // 12
            new_month = (new_month - 1) % 12 + 1
            try:
                return today.replace(year=new_year, month=new_month)
            except ValueError:
                # 월말 처리 (예: 1/31 + 1M = 2/28)
                import calendar
                last_day = calendar.monthrange(new_year, new_month)[1]
                return today.replace(year=new_year, month=new_month, day=min(today.day, last_day))
        elif 'Y' in s:
            try:
                return today.replace(year=today.year + num)
            except ValueError:
                # 윤년 처리 (2/29)
                return today.replace(year=today.year + num, day=28)
        else:
            # 기본: 일 단위
            return today + timedelta(days=num)

    def npv_error(self, fwd_guess, step_idx, solved_fwds, market_data, today):
        row = market_data.iloc[step_idx]
        current_fwds = solved_fwds.copy()
        current_fwds[step_idx] = float(fwd_guess) if hasattr(fwd_guess, '__iter__') else fwd_guess
        jump_dates = market_data['Jump Date'].iloc[:step_idx+1].tolist()
        
        mkt_rate, mty_date = row['Market Rate'], row['Mty Date']
        if str(row['Type']).lower() == "deposit":
            df = self.get_df_internal(mty_date, today, jump_dates, current_fwds)
            return (1 + mkt_rate * self.year_frac(today, mty_date)) * df - 1.0
        else:
            tenor_yf = self.parse_tenor(row['Inst. Tenor'])
            num_coupons = int(round(tenor_yf * self.freq))
            fixed_pv = 0.0
            p_date = today
            for j in range(1, num_coupons + 1):
                c_date = mty_date if j == num_coupons else today + timedelta(days=int(j * (365/self.freq)))
                yf = self.year_frac(p_date, c_date)
                fixed_pv += mkt_rate * yf * self.get_df_internal(c_date, today, jump_dates, current_fwds)
                p_date = c_date
            return fixed_pv - (1.0 - self.get_df_internal(mty_date, today, jump_dates, current_fwds))

    def run_batch(self, start_date, end_date):
        output_dir = "Batch_Results"
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"배치 작업 시작: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
        
        self.app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True, add_book=False)
        self.wb = self.app.books.open(self.file_path)
        ws_main = self.wb.sheets["Main"]
        
        current_date = start_date
        while current_date <= end_date:
            date_str = current_date.strftime('%Y-%m-%d')
            print(f"\n[{date_str}] 처리 중...")
            
            try:
                # 1. Today 날짜 업데이트
                tbl_common = ws_main.api.ListObjects("Common")
                ws_main.range(tbl_common.DataBodyRange.Cells(1, 1).Address).value = current_date
                self.app.calculate()
                
                # 2. 시장 데이터 및 JumpDates 테이블 로드
                tbl_market = ws_main.api.ListObjects("MarketTable")
                market_data = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
                market_data.columns = [str(c).strip() for c in market_data.columns]
                
                tbl_jumpdates = ws_main.api.ListObjects("JumpDates")
                jump_dates_df = ws_main.range(tbl_jumpdates.Range.Address).options(pd.DataFrame, index=False, header=True).value
                jump_dates_df.columns = [str(c).strip() for c in jump_dates_df.columns]
                jump_dates_list = sorted(pd.to_datetime(jump_dates_df['Jump Date']).tolist())
                
                # Python에서 Today + Tenor로 Mty Date 재계산
                market_data['Mty Date'] = market_data['Inst. Tenor'].apply(
                    lambda t: self.calc_mty_date(current_date, t))
                
                # Jump Date = JumpDates 중 Mty Date 이상인 날짜 중 최소값
                def find_jump_date(mty_date):
                    for jd in jump_dates_list:
                        if jd >= mty_date:
                            return jd
                    return jump_dates_list[-1]  # 없으면 마지막 날짜
                
                market_data['Jump Date'] = market_data['Mty Date'].apply(find_jump_date)
                
                # 3. Python 부트스트랩 실행
                n = len(market_data)
                solved_fwds = np.zeros(n)
                for i in range(n):
                    x0 = float(market_data.iloc[i]['Market Rate'])
                    result = fsolve(self.npv_error, x0, 
                                   args=(i, solved_fwds, market_data, current_date), 
                                   full_output=True)
                    solved_fwds[i] = result[0][0]
                
                # 4. 엑셀에 결과 저장 (Mty Date, Jump Date, Solved Forward)
                ws_main.range(tbl_market.ListColumns("Mty Date").DataBodyRange.Address).value = [[d] for d in market_data['Mty Date']]
                ws_main.range(tbl_market.ListColumns("Jump Date").DataBodyRange.Address).value = [[d] for d in market_data['Jump Date']]
                ws_main.range(tbl_market.ListColumns("Solved Forward").DataBodyRange.Address).value = solved_fwds.reshape(-1, 1)
                self.app.calculate()
                
                # 5. 차트 생성
                fig = go.Figure()
                
                fig.add_trace(go.Scatter(
                    x=market_data['Mty Date'], y=market_data['Market Rate'],
                    mode='markers+text', name='Market Rate',
                    text=[f"{r:.2%}" for r in market_data['Market Rate']],
                    textposition="top center",
                    marker=dict(size=8, color='gray')
                ))
                
                step_x, step_y, step_text = [], [], []
                p_date = current_date
                for idx, row in market_data.iterrows():
                    c_date = row['Jump Date']
                    f_rate = solved_fwds[idx]
                    num_days = (c_date - p_date).days
                    period_str = f"{p_date.strftime('%Y-%m-%d')} ~ {c_date.strftime('%Y-%m-%d')}"
                    for d in range(num_days + 1):
                        step_x.append(p_date + timedelta(days=d))
                        step_y.append(f_rate)
                        step_text.append(period_str)
                    step_x.append(None); step_y.append(None); step_text.append(None)
                    p_date = c_date
                
                fig.add_trace(go.Scatter(
                    x=step_x, y=step_y,
                    mode='lines', name='Forward Curve',
                    line=dict(color='red', width=3),
                    customdata=step_text,
                    hovertemplate="Rate: %{y:.4%}<br>Period: %{customdata}<extra></extra>"
                ))
                
                fig.update_layout(
                    title=f"IRS Forward Curve - {date_str}",
                    xaxis=dict(title="Date", type='date', tickformat='%Y-%m-%d'),
                    yaxis=dict(title="Rate (%)", tickformat=".2%"),
                    template="plotly_white", width=1200, height=700
                )
                
                output_file = os.path.join(output_dir, f"Bootstrap_{date_str}.html")
                fig.write_html(output_file)
                print(f"  -> OK: {output_file}")
                
            except Exception as e:
                print(f"  -> ERROR: {e}")
            
            current_date += timedelta(days=1)
        
        self.wb.save()
        print(f"\n모든 작업 완료! 결과: '{output_dir}' 폴더")

if __name__ == "__main__":
    runner = BatchBootstrapper("IRS_Bootstrap_DateBased.xlsm")
    runner.run_batch(datetime(2026, 1, 14), datetime(2026, 1, 25))
