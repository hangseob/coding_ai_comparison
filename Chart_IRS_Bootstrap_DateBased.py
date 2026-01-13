import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import xlwings as xw
from datetime import datetime, timedelta
import os
import numpy as np

def generate_irs_chart():
    target_excel = "IRS_Bootstrap_DateBased.xlsm"
    print(f"'{target_excel}' 연결 시도 중...")

    wb = None
    app = None
    
    try:
        if xw.apps.count > 0:
            for a in xw.apps:
                for b in a.books:
                    if b.name == target_excel:
                        wb = b
                        app = a
                        break
                if wb: break
        
        if wb is None:
            if not os.path.exists(target_excel):
                print(f"Error: '{target_excel}' 파일을 찾을 수 없습니다.")
                return
            app = xw.App(visible=False)
            wb = app.books.open(target_excel)
            
        ws_main = wb.sheets['Main']
        
        # 데이터 읽기
        tbl_common = ws_main.api.ListObjects("Common")
        df_common = ws_main.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
        today = df_common['Today'].iloc[0]

        tbl_market = ws_main.api.ListObjects("MarketTable")
        df_market = ws_main.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
        df_market.columns = [str(c).strip() for c in df_market.columns]

        tbl_jump = ws_main.api.ListObjects("JumpDates")
        df_jump = ws_main.range(tbl_jump.Range.Address).options(pd.DataFrame, index=False, header=True).value
        
    finally:
        if wb and app and not app.visible:
            wb.close()
            app.quit()

    df_plot = df_market.dropna(subset=['Mty Date', 'Solved Forward']).copy()
    
    # 테너 라벨 미리 생성
    df_plot['Tenor_Str'] = df_plot['Mty YearFrac'].apply(lambda x: f"{x:.2f}y")
    
    # 챠트 높이 및 간격 설정 (간격 더 확대)
    chart_height = 700 
    common_table_height = (len(df_common) + 1) * 50
    market_table_height = (len(df_market) + 1) * 35
    jump_table_height = (len(df_jump) + 1) * 35
    total_height = chart_height + common_table_height + market_table_height + jump_table_height + 500
    
    row_heights = [chart_height/total_height, common_table_height/total_height, market_table_height/total_height, jump_table_height/total_height]

    fig = make_subplots(
        rows=4, cols=1,
        vertical_spacing=0.08, # 간격 더 확대
        specs=[[{"type": "scatter"}], [{"type": "table"}], [{"type": "table"}], [{"type": "table"}]],
        row_heights=row_heights
    )

    # 1. Chart - 통합 호버 데이터 구성
    custom_data_array = df_plot[['Market Rate', 'Mty Zero Rate', 'Tenor_Str']].values

    # Market Rate
    fig.add_trace(go.Scatter(
        x=df_plot['Mty Date'], y=df_plot['Market Rate'],
        mode='markers+text', name='Market Rate',
        marker=dict(size=10, color='#1f77b4'),
        text=df_plot['Tenor_Str'],
        textposition="bottom center", textfont=dict(size=9),
        customdata=custom_data_array,
        hovertemplate=(
            "<b>Maturity Date: %{x|%Y-%m-%d}</b><br>" +
            "<b>Market Rate: %{customdata[0]:.4%}</b><br>" +
            "Zero Rate: %{customdata[1]:.4%}<br>" +
            "Tenor: %{customdata[2]}<extra></extra>"
        )
    ), row=1, col=1)

    # Zero Rate
    fig.add_trace(go.Scatter(
        x=df_plot['Mty Date'], y=df_plot['Mty Zero Rate'],
        mode='markers', name='Zero Rate',
        marker=dict(size=8, color='#d62728', symbol='diamond'),
        customdata=custom_data_array,
        hovertemplate=(
            "<b>Maturity Date: %{x|%Y-%m-%d}</b><br>" +
            "Market Rate: %{customdata[0]:.4%}<br>" +
            "<b>Zero Rate: %{customdata[1]:.4%}</b><br>" +
            "Tenor: %{customdata[2]}<extra></extra>"
        )
    ), row=1, col=1)

    # Forward Step Lines
    prev_date = today
    all_ticks = [today]
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
            hovertemplate=f"<b>Forward Rate: {fwd_rate:.4%}</b><br>Period: {prev_date.strftime('%Y-%m-%d')} ~ {curr_date.strftime('%Y-%m-%d')}<extra></extra>"
        ), row=1, col=1)
        fig.add_vline(x=curr_date, line_width=1, line_dash="dot", line_color="#ddd", row=1, col=1)
        all_ticks.append(curr_date)
        prev_date = curr_date

    fig.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='#2ca02c', width=3), name='Solved Forward (Step)'), row=1, col=1)

    # --- Tables Setup ---
    # 2. Common Table
    fig.add_trace(go.Table(
        columnwidth=[1, 1, 1],
        header=dict(values=[f"<b>{c}</b>" for c in df_common.columns], fill_color='#f2f2f2', align='center'),
        cells=dict(values=[[today.strftime('%Y-%m-%d')], [df_common['DayCount Basis'].iloc[0]], [int(df_common['IRS Coupon Freq'].iloc[0])]], height=30, align='center')
    ), row=2, col=1)

    # 3. MarketTable (포맷팅 적용)
    mkt_disp = df_market.copy()
    for c in ['Mty Date', 'Jump Date']: mkt_disp[c] = mkt_disp[c].dt.strftime('%Y-%m-%d')
    for c in ['Market Rate', 'Solved Forward', 'Jump Zero Rate', 'Mty Zero Rate']:
        if c in mkt_disp.columns: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.4%}" if pd.notnull(x) else "")
    for c in ['Jump Date DCF', 'Mty Date DCF']:
        if c in mkt_disp.columns: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.6f}" if pd.notnull(x) else "")
    for c in ['Mty YearFrac', 'Jump YearFrac']:
        if c in mkt_disp.columns: mkt_disp[c] = mkt_disp[c].apply(lambda x: f"{x:.4f}" if pd.notnull(x) else "")
    if 'Bootstrap Error' in mkt_disp.columns:
        mkt_disp['Bootstrap Error'] = mkt_disp['Bootstrap Error'].apply(lambda x: f"{x:.2e}" if pd.notnull(x) else "")

    fig.add_trace(go.Table(
        columnwidth=[0.4, 0.7, 0.7, 1, 0.7, 1, 0.7, 0.9, 0.9, 0.9, 1, 1, 1, 1],
        header=dict(values=[f"<b>{c}</b>" for c in mkt_disp.columns], fill_color='#4472c4', font=dict(color='white', size=10), align='center'),
        cells=dict(values=[mkt_disp[k].tolist() for k in mkt_disp.columns], fill_color=[['white', '#f9f9f9']*len(mkt_disp)], height=25, align='center', font=dict(size=10))
    ), row=3, col=1)

    # 4. JumpDates Table
    df_jump_disp = df_jump.copy()
    df_jump_disp['Jump Date'] = df_jump_disp['Jump_Date_Str'] = df_jump_disp['Jump Date'].dt.strftime('%Y-%m-%d')
    fig.add_trace(go.Table(
        columnwidth=[0.5, 1],
        header=dict(values=[f"<b>{c}</b>" for c in df_jump.columns], fill_color='#70ad47', font=dict(color='white'), align='center'),
        cells=dict(values=[df_jump_disp['No'].tolist(), df_jump_disp['Jump_Date_Str'].tolist()], height=25, align='center')
    ), row=4, col=1)

    # Layout
    all_ticks = sorted(list(set(all_ticks + df_plot['Mty Date'].tolist())))
    fig.update_layout(
        title=dict(text=f"IRS Bootstrapping Results Analysis<br><span style='font-size:14px; color:gray;'>Target: {target_excel} | Base Date: {today.strftime('%Y-%m-%d')}</span>", x=0.5, y=0.98, xanchor='center', yanchor='top', font=dict(size=22)),
        xaxis=dict(title="Date (Actual Time Scale)", type='date', tickmode='array', tickvals=all_ticks, ticktext=[d.strftime('%y-%m-%d') for d in all_ticks], tickangle=-90, tickfont=dict(size=8), gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True),
        yaxis=dict(title="Rate (%)", tickformat=".2%", gridcolor='#eee', showline=True, linewidth=1, linecolor='black', mirror=True, zeroline=False),
        template="plotly_white", width=1500, height=total_height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0.5, xanchor='center', bgcolor='rgba(255, 255, 255, 0.5)'),
        margin=dict(l=70, r=70, t=200, b=100),
        hovermode="closest"
    )

    output_html = "IRS_Bootstrap_Analysis.html"
    fig.write_html(output_html, full_html=True, include_plotlyjs='cdn')
    print(f"챠트 생성 및 HTML 저장 완료: {output_html}")
    fig.show()

if __name__ == "__main__": generate_irs_chart()
