import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import xlwings as xw
from datetime import datetime
import os
import numpy as np

def generate_homework_chart():
    # 1. 엑셀 연결 (이름이 '숙제.6.4'로 시작하는 파일 찾기)
    target_wb = None
    target_app = None
    
    print("엑셀 인스턴스 탐색 중...")
    for app in xw.apps:
        for book in app.books:
            if book.name.startswith("숙제.6.4"):
                target_wb = book
                target_app = app
                break
        if target_wb: break
        
    if not target_wb:
        print("Error: '숙제.6.4'로 시작하는 엑셀 파일을 찾을 수 없습니다.")
        # 디버깅을 위해 현재 열려있는 파일 목록 출력
        print("현재 열려 있는 파일들:")
        for app in xw.apps:
            for book in app.books:
                print(f" - {book.name}")
        return

    print(f"연결된 파일: {target_wb.name}")
    
    try:
        ws_main = target_wb.sheets['Main']
        
        # 데이터 읽기
        tbl_common = ws_main.api.ListObjects("Common")
        df_common = ws_main.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
        today = df_common['Today'].iloc[0]

        tbl_market = ws_main.api.ListObjects("MarketTable")
        df_market = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_market.columns = [str(c).strip() for c in df_market.columns]

        # 테이블 2개를 제외하라고 했으므로 Common과 JumpDates는 차트에 포함하지 않음
        # MarketTable만 하단에 표시함
        
    except Exception as e:
        print(f"데이터 읽기 중 오류 발생: {e}")
        return

    df_plot = df_market.dropna(subset=['Mty Date', 'Solved Forward']).copy()
    df_plot['Tenor_Str'] = df_plot['Mty YearFrac'].apply(lambda x: f"{x:.2f}y")
    
    # 챠트 레이아웃 설정
    chart_height = 750
    market_table_height = (len(df_market) + 1) * 35 + 50
    total_height = chart_height + market_table_height + 200
    
    row_heights = [chart_height/total_height, market_table_height/total_height]

    fig = make_subplots(
        rows=2, cols=1,
        vertical_spacing=0.1,
        specs=[[{"type": "scatter"}], [{"type": "table"}]],
        row_heights=row_heights
    )

    # 1. Chart - Market Rate (Zero Rate 제외)
    # Market Rate trace (Zero Rate 정보를 포함한 단일 호버 박스)
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
    ), row=1, col=1)

    # Forward Step Lines
    prev_date = today
    all_ticks = [today]
    for i, row in df_plot.iterrows():
        curr_date = row['Jump Date']
        fwd_rate = row['Solved Forward']
        # 선 중앙에 텍스트를 배치하기 위한 mid_date 계산
        mid_date = prev_date + (curr_date - prev_date) / 2
        
        fig.add_trace(go.Scatter(
            x=[prev_date, mid_date, curr_date], y=[fwd_rate, fwd_rate, fwd_rate],
            mode='lines+text', line=dict(color='#2ca02c', width=3),
            text=["", f"<b>{fwd_rate:.2%}</b>", ""],
            textposition="top center", textfont=dict(size=11, color="green"),
            showlegend=False,
            hovertemplate=f"<b>Forward Rate: {fwd_rate:.4%}</b><br>Period: <b>{prev_date.strftime('%Y-%m-%d')}</b> ~ <b>{curr_date.strftime('%Y-%m-%d')}</b><extra></extra>"
        ), row=1, col=1)
        
        # Jump Date 수직선
        fig.add_vline(x=curr_date, line_width=1, line_dash="dash", line_color="rgba(200,200,200,0.5)", row=1, col=1)
        all_ticks.append(curr_date)
        prev_date = curr_date

    # 범례용 가짜 트레이스 (Forward)
    fig.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='#2ca02c', width=3), name='Solved Forward'), row=1, col=1)

    # 2. Table - MarketTable만 포함 (Common, JumpDates 제외)
    mkt_disp = df_market.copy()
    # 포맷팅
    date_cols = [c for c in ['Mty Date', 'Jump Date'] if c in mkt_disp.columns]
    for c in date_cols: mkt_disp[c] = mkt_disp[c].dt.strftime('%Y-%m-%d')
    
    rate_cols = [c for c in ['Market Rate', 'Solved Forward', 'Jump Zero Rate', 'Mty Zero Rate'] if c in mkt_disp.columns]
    for c in rate_cols: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.4%}" if pd.notnull(x) else "")
    
    dcf_cols = [c for c in ['Jump Date DCF', 'Mty Date DCF'] if c in mkt_disp.columns]
    for c in dcf_cols: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.6f}" if pd.notnull(x) else "")
    
    yf_cols = [c for c in ['Mty YearFrac', 'Jump YearFrac'] if c in mkt_disp.columns]
    for c in yf_cols: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.4f}" if pd.notnull(x) else "")

    # Zero Rate 컬럼 제거 (사용자 요청에 따라 데이터에서도 제외)
    cols_to_show = [c for c in mkt_disp.columns if 'Zero Rate' not in str(c)]
    mkt_disp = mkt_disp[cols_to_show]

    fig.add_trace(go.Table(
        columnwidth=[0.4, 0.7, 1, 0.7, 1, 0.7, 0.9, 0.9, 0.9, 1],
        header=dict(values=[f"<b>{c}</b>" for c in mkt_disp.columns], fill_color='#4472c4', font=dict(color='white', size=11), align='center'),
        cells=dict(values=[mkt_disp[k].tolist() for k in mkt_disp.columns], fill_color=[['white', '#f9f9f9']*len(mkt_disp)], height=28, align='center', font=dict(size=10))
    ), row=2, col=1)

    # Layout 조정
    all_ticks = sorted(list(set(all_ticks + df_plot['Mty Date'].tolist())))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"숙제.6.4 모범답안 CHART_{timestamp}.html"

    fig.update_layout(
        title=dict(
            text=f"IRS Bootstrap 모범답안 분석 차트<br><span style='font-size:14px; color:gray;'>파일: {target_wb.name} | 기준일: {today.strftime('%Y-%m-%d')}</span>",
            x=0.5, y=0.97, xanchor='center', yanchor='top', font=dict(size=24)
        ),
        xaxis=dict(
            title="Date", type='date', tickmode='array', tickvals=all_ticks,
            ticktext=[d.strftime('%y-%m-%d') for d in all_ticks], tickangle=-90, tickfont=dict(size=9),
            gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True
        ),
        yaxis=dict(
            title="Rate (%)", tickformat=".2%", gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True, zeroline=False
        ),
        template="plotly_white", width=1400, height=total_height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0.5, xanchor='center'),
        margin=dict(l=80, r=80, t=180, b=80),
        hovermode="closest"
    )

    fig.write_html(output_filename, full_html=True, include_plotlyjs='cdn')
    print(f"HTML 차트 생성 완료: {output_filename}")
    # fig.show() # 로컬 실행 시 확인용

if __name__ == "__main__":
    generate_homework_chart()
