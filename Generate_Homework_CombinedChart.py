import pandas as pd
import plotly.graph_objects as go
import xlwings as xw
from datetime import datetime, timedelta
import os

def generate_combined_forward_chart():
    # 1. 엑셀 연결
    target_wb = None
    print("엑셀 인스턴스 탐색 중...")
    for app in xw.apps:
        for book in app.books:
            if "6.4" in book.name or "6.5" in book.name:
                target_wb = book
                break
        if target_wb: break
        
    if not target_wb:
        print("Error: 관련 엑셀 파일을 찾을 수 없습니다.")
        return

    print(f"연결된 파일: {target_wb.name}")
    
    fig = go.Figure()
    all_ticks = []
    colors = ['#1f77b4', '#d62728'] # 파랑(금통위), 빨강(채권매칭)
    line_styles = ['solid', 'dash']
    scenarios = []

    # 두 번째(index 1)와 세 번째(index 2) 시트 처리
    for i, sheet_idx in enumerate([1, 2]):
        try:
            ws = target_wb.sheets[sheet_idx]
            # 사용자 요청에 따른 시나리오 이름 설정
            scenario_name = "금통위 노드" if i == 0 else "채권만기 노드"
            print(f"시트 로드 중: {ws.name} ({scenario_name})")
            
            # 테이블 찾기
            tbl_market = next(t for t in ws.api.ListObjects if "MarketTable" in t.Name)
            tbl_common = next(t for t in ws.api.ListObjects if "Common" in t.Name)
            
            # 데이터 로드
            df_common = ws.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
            today = df_common['Today'].iloc[0]
            all_ticks.append(today)

            df_market = ws.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
            df_market.columns = [str(c).strip() for c in df_market.columns]
            
            # 위치 기반 컬럼 매핑
            cols = list(df_market.columns)
            mapping = {cols[3]: 'Mty Date', cols[4]: 'Market Rate', cols[6]: 'Jump Date', cols[8]: 'Solved Forward'}
            df_market = df_market.rename(columns=mapping)
            
            df_plot = df_market.dropna(subset=['Mty Date', 'Solved Forward']).copy()
            df_plot = df_plot.sort_values('Jump Date')

            # --- Trace 구성 ---
            
            # 1. Market Rate (Scatter) - 첫 번째 루프에서만 범례에 표시
            fig.add_trace(go.Scatter(
                x=df_plot['Mty Date'], y=df_plot['Market Rate'],
                mode='markers',
                name="Market Rate",
                marker=dict(size=10, color='#555', symbol='circle'), # 조금 더 진한 회색
                showlegend=(i == 0), # 한 번만 표시
                hoverlabel=dict(bgcolor="white", font_size=13, font_family="Arial"), # 호버 박스 배경 흰색으로 고정
                hovertemplate="<b>Market Rate</b><br>Date: %{x|%Y-%m-%d}<br>Rate: %{y:.4%}<extra></extra>"
            ))

            # 2. Forward Horizontal Lines (일 단위 샘플링으로 전체 호버 가능하게 수정)
            step_x = []
            step_y = []
            step_text = []
            
            p_date = today
            for _, row in df_plot.iterrows():
                c_date = row['Jump Date']
                f_rate = row['Solved Forward']
                
                # 일 단위로 포인트 생성하여 선 전체를 호버 가능하게 함
                num_days = (c_date - p_date).days
                period_str = f"<b>{p_date.strftime('%Y-%m-%d')} ~ {c_date.strftime('%Y-%m-%d')}</b>"
                
                for d in range(num_days + 1):
                    current_d = p_date + timedelta(days=d)
                    step_x.append(current_d)
                    step_y.append(f_rate)
                    step_text.append(period_str)
                
                step_x.append(None) # 구간 분리
                step_y.append(None)
                step_text.append(None)
                
                p_date = c_date
            
            fig.add_trace(go.Scatter(
                x=step_x, y=step_y,
                mode='lines',
                name=f"Forward ({scenario_name})",
                line=dict(color=colors[i], width=4, dash=line_styles[i]),
                legendgroup=scenario_name,
                showlegend=True,
                connectgaps=False,
                customdata=step_text,
                hoverlabel=dict(bgcolor="white", font_size=13, font_family="Arial"),
                hovertemplate=f"<span style='color:{colors[i]};'><b>{scenario_name} Forward</b></span><br>Rate: %{{y:.4%}}<br>Period: %{{customdata}}<extra></extra>"
            ))

            # 수직선 날짜 수집
            all_ticks.extend(df_plot['Jump Date'].tolist())
            all_ticks.extend(df_plot['Mty Date'].tolist())

        except Exception as e:
            print(f"Error processing sheet {sheet_idx}: {e}")

    # Layout 설정
    unique_ticks = sorted(list(set(all_ticks)))
    timestamp = datetime.now().strftime("%H%M%S")
    output_filename = f"Forward_Chart_Combined_Analysis_{timestamp}.html"

    fig.update_layout(
        title=dict(
            text=f"Combined Forward Curve Analysis<br><span style='font-size:15px; color:gray;'>File: {target_wb.name}</span>",
            x=0.5, y=0.96, xanchor='center', yanchor='top', font=dict(size=24)
        ),
        xaxis=dict(
            title="Date", type='date', tickvals=unique_ticks, tickformat='%y-%m-%d', 
            tickangle=-90, tickfont=dict(size=18), gridcolor='#eee'
        ),
        yaxis=dict(title="Rate (%)", tickformat=".2%", tickfont=dict(size=14), gridcolor='#eee'),
        template="plotly_white", width=1550, height=950,
        margin=dict(t=220, b=180, l=80, r=80),
        hovermode="closest",
        legend=dict(
            orientation="h", 
            yanchor="bottom", 
            y=1.02, 
            x=0.5, 
            xanchor='center', 
            font=dict(size=15),
            bgcolor='rgba(255, 255, 255, 0.7)',
            bordercolor="LightGray",
            borderwidth=1
        )
    )

    # 금통위 날짜 가이드라인 (옵션)
    for tick in unique_ticks:
        fig.add_vline(x=tick, line_width=1, line_dash="dot", line_color="rgba(200,200,200,0.3)")

    fig.write_html(output_filename, full_html=True, include_plotlyjs='cdn')
    print(f"\n성공! 통합 차트가 생성되었습니다: {output_filename}")

if __name__ == "__main__":
    generate_combined_forward_chart()
