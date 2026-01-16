import xlwings as xw
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os

def run_bootstrap_and_chart():
    file_path = "IRS_Bootstrap_DateBased.xlsm"
    
    if not os.path.exists(file_path):
        print(f"Error: {file_path} 파일을 찾을 수 없습니다.")
        return

    print(f"엑셀 파일 연결 중: {file_path}")
    
    # 이미 열려있는 엑셀 앱 찾기 또는 새로 열기
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
        else:
            app = xw.App(visible=True, add_book=False)
            
        wb = app.books.open(file_path)
    except Exception as e:
        print(f"엑셀 연결 실패: {e}")
        return
    
    try:
        # 1. 엑셀 매크로 실행 (RunBootstrap)
        print("엑셀 매크로(RunBootstrap) 실행 중... (계산량이 많으면 수 초 걸릴 수 있습니다)")
        # 참고: 매크로 마지막에 MsgBox가 있다면 엑셀에서 확인 버튼을 눌러야 다음 단계로 진행됩니다.
        macro = wb.macro("RunBootstrap")
        macro()
        
        # MarketTable 수식 결과값 강제 업데이트 (부트스트랩 결과 반영을 위함)
        print("수식 결과값 최종 업데이트 중 (Calculate)...")
        app.calculate()
        
        # 계산이 완료될 때까지 잠시 대기 (안정성 확보)
        import time
        time.sleep(1) 
        
        print("계산 및 업데이트 완료!")

        # 2. 데이터 시트 및 테이블 로드
        ws_main = wb.sheets["Main"]
        
        # Common 테이블 읽기 (Today 날짜 추출)
        tbl_common = ws_main.api.ListObjects("Common")
        df_common = ws_main.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_common.columns = [str(c).strip() for c in df_common.columns] # 컬럼명 정리
        
        if 'Today' in df_common.columns:
            today = df_common['Today'].iloc[0]
        else:
            # 컬럼명이 다를 경우 첫 번째 셀값 사용 시도
            today = ws_main.range("A2").value 
            
        if not isinstance(today, datetime):
            today = pd.to_datetime(today)
            
        print(f"기준일(Today): {today.strftime('%Y-%m-%d')}")

        # MarketTable 읽기
        tbl_market = ws_main.api.ListObjects("MarketTable")
        df_market = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_market.columns = [str(c).strip() for c in df_market.columns] # 컬럼명 정리

        # 3. 데이터 필터링 및 차트용 가공
        # 필요한 컬럼 확인
        required_cols = ['Mty Date', 'Solved Forward', 'Market Rate', 'Jump Date']
        for col in required_cols:
            if col not in df_market.columns:
                print(f"Error: '{col}' 컬럼을 테이블에서 찾을 수 없습니다.")
                print(f"현재 컬럼들: {list(df_market.columns)}")
                return

        # 날짜 형식 확실히 변환
        df_market['Jump Date'] = pd.to_datetime(df_market['Jump Date'])
        df_market['Mty Date'] = pd.to_datetime(df_market['Mty Date'])

        # Today 이후의 유효한 데이터만 필터링
        df_plot = df_market.dropna(subset=['Jump Date', 'Solved Forward']).copy()
        df_plot = df_plot[df_plot['Jump Date'] >= today].sort_values('Jump Date')

        if df_plot.empty:
            print("Error: 시각화할 유효한 데이터가 없습니다. (Today 이후 데이터 확인 필요)")
            return

        # X축 범위 계산 (Today ~ 마지막 Jump Date)
        max_jump_date = df_plot['Jump Date'].max()
        x_range_start = today - timedelta(days=5) # 시작점에 약간의 여유
        x_range_end = max_jump_date + timedelta(days=20) # 끝점에 약간의 여유

        # 4. Plotly 차트 생성
        fig = go.Figure()

        # Market Rate 포인트
        fig.add_trace(go.Scatter(
            x=df_plot['Mty Date'], y=df_plot['Market Rate'],
            mode='markers+text', name='Market Rate',
            text=[f"{r:.2%}" for r in df_plot['Market Rate']],
            textposition="top center",
            marker=dict(size=10, color='gray', symbol='circle'),
            hoverlabel=dict(bgcolor="white"),
            hovertemplate="<b>Market Rate</b><br>Date: %{x|%Y-%m-%d}<br>Rate: %{y:.4%}<extra></extra>"
        ))

        # Solved Forward (Daily Sampling Step Line)
        step_x, step_y, step_text = [], [], []
        p_date = today
        for _, row in df_plot.iterrows():
            c_date = row['Jump Date']
            f_rate = row['Solved Forward']
            
            if pd.isna(c_date) or pd.isna(f_rate): continue
            
            # 구간 내 매일 포인트 생성하여 선 전체 호버 가능하게 함
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
            mode='lines', name='Solved Forward (Step)',
            line=dict(color='red', width=4),
            customdata=step_text,
            hoverlabel=dict(bgcolor="white"),
            hovertemplate="<b>Forward Rate</b><br>Rate: %{y:.4%}<br>Period: %{customdata}<extra></extra>"
        ))

        # 레이아웃 설정
        fig.update_layout(
            title=dict(
                text=f"IRS Bootstrapping: Forward Curve Analysis<br><span style='font-size:14px; color:gray;'>Base Date: {today.strftime('%Y-%m-%d')}</span>",
                x=0.5, font=dict(size=22)
            ),
            xaxis=dict(
                title="Date", 
                type='date', 
                tickformat='%Y-%m-%d', 
                gridcolor='#eee', 
                showline=True,
                range=[x_range_start, x_range_end], # 스케일 고정
                tickangle=-45,
                nticks=20 # 틱 개수 최적화
            ),
            yaxis=dict(
                title="Rate (%)", 
                tickformat=".2%", 
                gridcolor='#eee', 
                showline=True,
                zeroline=False
            ),
            template="plotly_white",
            width=1300, height=750,
            margin=dict(t=150, b=100), # 하단 여유 공간 확보
            hovermode="closest",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0.5, xanchor='center')
        )

        # 차트 출력
        fig.show()
        print("성공! 차트가 브라우저에 표시되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_bootstrap_and_chart()
