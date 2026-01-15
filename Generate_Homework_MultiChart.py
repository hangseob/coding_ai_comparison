import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import xlwings as xw
from datetime import datetime
import os

def generate_multi_sheet_forward_charts():
    # 1. 엑셀 연결 (이름이 '숙제 6.4' 또는 '숙제 6.5'로 시작하는 파일 찾기)
    target_wb = None
    
    print("엑셀 인스턴스 탐색 중...")
    for app in xw.apps:
        for book in app.books:
            if "6.4" in book.name or "6.5" in book.name:
                target_wb = book
                break
        if target_wb: break
        
    if not target_wb:
        print("Error: '숙제 6.4' 또는 '숙제 6.5' 관련 엑셀 파일을 찾을 수 없습니다.")
        return

    print(f"연결된 파일: {target_wb.name}")
    
    # 두 번째(index 1)와 세 번째(index 2) 시트 처리
    for sheet_idx in [1, 2]:
        try:
            ws = target_wb.sheets[sheet_idx]
            print(f"\n시트 처리 중: {ws.name}")
            
            # 테이블 찾기
            tbl_market = None
            tbl_common = None
            
            for tbl in ws.api.ListObjects:
                if "MarketTable" in tbl.Name:
                    tbl_market = tbl
                elif "Common" in tbl.Name:
                    tbl_common = tbl
            
            if not tbl_market or not tbl_common:
                print(f"  - Skip: 필요한 테이블을 {ws.name} 시트에서 찾을 수 없습니다.")
                continue

            # 데이터 로드
            df_common = ws.range(tbl_common.Range.Address).options(pd.DataFrame, index=False, header=True).value
            today = df_common['Today'].iloc[0]

            df_market = ws.range(tbl_market.Range.Address).options(pd.DataFrame, index=False, header=True).value
            df_market.columns = [str(c).strip() for c in df_market.columns]
            
            # 위치 기반 컬럼 매핑 (한글 인코딩 및 자동 번호 대응)
            cols = list(df_market.columns)
            if len(cols) >= 10:
                # 3: Mty Date, 4: Market Rate, 6: Jump Date, 8: Solved Forward (fwd)
                mapping = {
                    cols[3]: 'Mty Date',
                    cols[4]: 'Market Rate',
                    cols[6]: 'Jump Date',
                    cols[8]: 'Solved Forward'
                }
                df_market = df_market.rename(columns=mapping)
            
            # 데이터 정제
            df_plot = df_market.dropna(subset=['Mty Date', 'Solved Forward']).copy()
            df_plot['Tenor_Str'] = df_plot.apply(lambda x: f"{x['Inst. Tenor']}" if 'Inst. Tenor' in x else "", axis=1)

            # 차트 생성
            fig = go.Figure()

            # Market Rate
            fig.add_trace(go.Scatter(
                x=df_plot['Mty Date'], y=df_plot['Market Rate'],
                mode='markers+text', name='Market Rate',
                marker=dict(size=12, color='#1f77b4', line=dict(width=1, color='white')),
                text=df_plot['Tenor_Str'],
                textposition="bottom center", textfont=dict(size=10),
                hovertemplate="<b>Mty Date: %{x|%Y-%m-%d}</b><br>Rate: %{y:.4%}<extra></extra>"
            ))

            # Forward Step Lines
            prev_date = today
            all_ticks = [today]
            for _, row in df_plot.iterrows():
                curr_date = row['Jump Date']
                fwd_rate = row['Solved Forward']
                mid_date = prev_date + (curr_date - prev_date) / 2
                
                fig.add_trace(go.Scatter(
                    x=[prev_date, mid_date, curr_date], y=[fwd_rate, fwd_rate, fwd_rate],
                    mode='lines+text', line=dict(color='#2ca02c', width=3),
                    text=["", f"<b>{fwd_rate:.2%}</b>", ""],
                    textposition="top center", textfont=dict(size=11, color="green"),
                    showlegend=False,
                    hovertemplate=f"<b>Forward: {fwd_rate:.4%}</b><br>{prev_date.strftime('%Y-%m-%d')} ~ {curr_date.strftime('%Y-%m-%d')}<extra></extra>"
                ))
                fig.add_vline(x=curr_date, line_width=1, line_dash="dash", line_color="rgba(150,150,150,0.5)")
                all_ticks.append(curr_date)
                prev_date = curr_date

            fig.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color='#2ca02c', width=3), name='Solved Forward'))

            # 레이아웃
            all_ticks = sorted(list(set(all_ticks + df_plot['Mty Date'].tolist())))
            timestamp = datetime.now().strftime("%H%M%S")
            output_filename = f"Forward_Chart_{ws.name}_{timestamp}.html"

            fig.update_layout(
                title=dict(text=f"Forward Chart Analysis - {ws.name}<br><span style='font-size:14px; color:gray;'>File: {target_wb.name}</span>", x=0.5),
                xaxis=dict(type='date', tickvals=all_ticks, tickformat='%y-%m-%d', tickangle=-90, tickfont=dict(size=18)),
                yaxis=dict(title="Rate (%)", tickformat=".2%", tickfont=dict(size=14)),
                template="plotly_white", width=1400, height=800,
                margin=dict(t=120, b=180),
                hovermode="closest"
            )

            fig.write_html(output_filename, full_html=True, include_plotlyjs='cdn')
            print(f"  - 성공: {output_filename}")

        except Exception as e:
            print(f"  - 에러 발생 ({sheet_idx}번째 시트): {e}")

if __name__ == "__main__":
    generate_multi_sheet_forward_charts()
