import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import xlwings as xw
from datetime import datetime
import os
import numpy as np

def generate_homework_analysis():
    # 1. 엑셀 연결 (이름이 '숙제.6.4'로 시작하는 파일 찾기)
    target_wb = None
    target_app = None
    
    print("엑셀 인스턴스 탐색 중...")
    for app in xw.apps:
        for book in app.books:
            # 한글 인코딩 문제 대응을 위해 startswith와 '6.4' 포함 여부 병행 확인
            if book.name.startswith("숙제.6.4") or "6.4" in book.name:
                target_wb = book
                target_app = app
                break
        if target_wb: break
        
    if not target_wb:
        print("Error: '숙제.6.4'로 시작하는 엑셀 파일을 찾을 수 없습니다.")
        return

    print(f"연결된 파일: {target_wb.name}")
    
    try:
        # 시트 이름 확인 (디버깅 결과에 근거: '1.메인화면', 'JumpDates')
        # 시트 이름이 다를 수 있으므로 인덱스나 테이블 검색으로 보완
        ws_main = None
        ws_jump = None
        
        for sheet in target_wb.sheets:
            tables = [t.name for t in sheet.api.ListObjects]
            if "MarketTable" in tables:
                ws_main = sheet
            if "JumpDates" in tables:
                ws_jump = sheet
        
        if not ws_main or not ws_jump:
            print("Error: 필요한 테이블(MarketTable 또는 JumpDates)을 시트에서 찾을 수 없습니다.")
            return

        # 데이터 읽기
        tbl_common = ws_main.api.ListObjects("Common")
        df_common = ws_main.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
        today = df_common['Today'].iloc[0]

        tbl_market = ws_main.api.ListObjects("MarketTable")
        df_market = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_market.columns = [str(c).strip() for c in df_market.columns]
        
        # 위치 기반 매핑 (한글 인코딩 문제 회피)
        if len(df_market.columns) >= 10:
            cols = list(df_market.columns)
            # 6: 금통위 날짜, 7: YearFrac to 금통위, 8: fwd, 9: DCF(금통위)
            mapping = {
                cols[6]: 'Jump Date',
                cols[7]: 'Jump YearFrac',
                cols[8]: 'Solved Forward',
                cols[9]: 'Jump Date DCF'
            }
            df_market = df_market.rename(columns=mapping)
        else:
            # 기존 이름 기반 매핑 시도
            mapping = {
                '금통위 날짜': 'Jump Date',
                'fwd': 'Solved Forward',
                'DCF(금통위)': 'Jump Date DCF',
                'YearFrac to 금통위': 'Jump YearFrac'
            }
            df_market = df_market.rename(columns=mapping)

        tbl_jump = ws_jump.api.ListObjects("JumpDates")
        df_jump = ws_jump.range(tbl_jump.Range.Address).options(pd.DataFrame, index=False, header=True).value
        # JumpDates 컬럼 매핑 (위치 기반)
        if len(df_jump.columns) >= 2:
            df_jump = df_jump.rename(columns={df_jump.columns[1]: 'Jump Date'})
        else:
            df_jump = df_jump.rename(columns={'금통위': 'Jump Date'})
        
    except Exception as e:
        print(f"데이터 읽기 중 오류 발생: {e}")
        return

    # 데이터 정제
    df_plot = df_market.dropna(subset=['Mty Date', 'Solved Forward']).copy()
    df_plot['Tenor_Str'] = df_plot['Mty YearFrac'].apply(lambda x: f"{x:.2f}y" if pd.notnull(x) else "")
    
    # 챠트 레이아웃 설정 (테이블 모두 제외, 차트만 포함)
    chart_height = 800
    total_height = chart_height + 200
    
    fig = go.Figure()

    # 1. Chart - Market Rate (Zero Rate 제외)
    fig.add_trace(go.Scatter(
        x=df_plot['Mty Date'], y=df_plot['Market Rate'],
        mode='markers+text', name='Market Rate',
        marker=dict(size=12, color='#1f77b4', line=dict(width=1, color='white')),
        text=df_plot['Tenor_Str'],
        textposition="bottom center", textfont=dict(size=10, color="black"),
        hovertemplate=(
            "<b>Maturity Date: %{x|%Y-%m-%d}</b><br>" +
            "<b>Market Rate: %{y:.4%}</b><br>" +
            "Tenor: %{text}<extra></extra>"
        )
    ))

    # Forward Step Lines (Solved Forward)
    prev_date = today
    all_ticks = [today]
    
    # Jump Date 수직선 및 Forward 선 그리기
    for i, row in df_plot.iterrows():
        curr_date = row['Jump Date']
        fwd_rate = row['Solved Forward']
        mid_date = prev_date + (curr_date - prev_date) / 2
        
        fig.add_trace(go.Scatter(
            x=[prev_date, mid_date, curr_date], y=[fwd_rate, fwd_rate, fwd_rate],
            mode='lines+text', line=dict(color='#2ca02c', width=3),
            text=["", f"<b>{fwd_rate:.2%}</b>", ""],
            textposition="top center", textfont=dict(size=11, color="green"),
            showlegend=False,
            hovertemplate=f"<b>Forward Rate: {fwd_rate:.4%}</b><br>Period: <b>{prev_date.strftime('%Y-%m-%d')}</b> ~ <b>{curr_date.strftime('%Y-%m-%d')}</b><extra></extra>"
        ))
        
        # Jump Date 수직선 (점선)
        fig.add_vline(x=curr_date, line_width=1, line_dash="dash", line_color="rgba(150,150,150,0.5)")
        all_ticks.append(curr_date)
        prev_date = curr_date

    # 범례용 가짜 트레이스 (Forward)
    fig.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='#2ca02c', width=3), name='Solved Forward'))

    # Layout 및 파일 저장
    all_ticks = sorted(list(set(all_ticks + df_plot['Mty Date'].tolist())))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"숙제.6.4 모범답안 CHART_{timestamp}.html"

    fig.update_layout(
        title=dict(
            text=f"IRS Bootstrap 모범답안 분석 (금통위 Jump Node 반영)<br><span style='font-size:14px; color:gray;'>Target: {target_wb.name} | Base Date: {today.strftime('%Y-%m-%d')}</span>",
            x=0.5, y=0.97, xanchor='center', yanchor='top', font=dict(size=22)
        ),
        xaxis=dict(
            title="Date (Jump Node & Maturity)", type='date', tickmode='array', tickvals=all_ticks,
            ticktext=[d.strftime('%y-%m-%d') for d in all_ticks], tickangle=-90, tickfont=dict(size=9),
            gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True
        ),
        yaxis=dict(
            title="Rate (%)", tickformat=".2%", gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True, zeroline=False
        ),
        template="plotly_white", width=1450, height=chart_height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0.5, xanchor='center'),
        margin=dict(l=80, r=80, t=150, b=120),
        hovermode="closest"
    )

    fig.write_html(output_filename, full_html=True, include_plotlyjs='cdn')
    print(f"\n성공! HTML 차트가 생성되었습니다: {output_filename}")

if __name__ == "__main__":
    generate_homework_analysis()
