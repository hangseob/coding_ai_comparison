import plotly.graph_objects as go
import pandas as pd
from datetime import datetime

def create_timeline_chart():
    # 1. 데이터 정의
    mat_dates_str = [
        "2026-01-22", "2026-01-29", "2026-02-05", "2026-02-15", 
        "2026-03-15", "2026-04-15", "2026-07-15", "2026-10-15", 
        "2027-01-15", "2027-07-15", "2028-01-15"
    ]
    
    bok_dates_str = [
        "2026-01-15", "2026-02-26", "2026-04-16", "2026-05-28", "2026-07-16", "2026-08-27", "2026-10-22", "2026-11-26",
        "2027-01-14", "2027-02-25", "2027-04-15", "2027-05-27", "2027-07-15", "2027-08-26", "2027-10-21", "2027-11-25",
        "2028-01-13", "2028-02-24"
    ]

    today_str = "2026-01-15" # 현재 기준일 문자열

    fig = go.Figure()

    # 2. 금통위 일정 (Y=0)
    fig.add_trace(go.Scatter(
        x=bok_dates_str,
        y=[0] * len(bok_dates_str),
        mode='markers+text',
        name='금통위 (BOK Meeting)',
        text=[d[5:] for d in bok_dates_str], # MM-DD만 표시
        textposition="bottom center",
        marker=dict(size=12, color='royalblue', symbol='diamond'),
        hovertemplate="<b>금통위</b><br>날짜: %{x}<extra></extra>"
    ))

    # 3. 만기일 일정 (Y=1)
    fig.add_trace(go.Scatter(
        x=mat_dates_str,
        y=[1] * len(mat_dates_str),
        mode='markers+text',
        name='만기일 (Mat Date)',
        text=[d[5:] for d in mat_dates_str], # MM-DD만 표시
        textposition="top center",
        marker=dict(size=12, color='firebrick', symbol='circle'),
        hovertemplate="<b>만기일</b><br>날짜: %{x}<extra></extra>"
    ))

    # 4. 오늘 기준선 표시 (텍스트 분리)
    fig.add_vline(x=today_str, line_width=3, line_dash="dash", line_color="green")
    fig.add_annotation(
        x=today_str, y=1.4, 
        text="Today (2026-01-15)", 
        showarrow=False, 
        font=dict(color="green", size=14),
        xanchor="left"
    )

    # 5. 수직 가이드라인 (모든 날짜에 대해)
    for d in mat_dates_str + bok_dates_str:
        fig.add_vline(x=d, line_width=0.5, line_dash="dot", line_color="rgba(200,200,200,0.3)")

    # 6. 레이아웃 설정
    fig.update_layout(
        title=dict(text="Timeline Comparison: Mat Dates vs. BOK Meetings", x=0.5, font=dict(size=20)),
        xaxis=dict(
            title="Timeline",
            type='date',
            tickformat='%Y-%m-%d',
            dtick="M3", # 3개월 단위 메인 그리드
            gridcolor='lightgray',
            showgrid=True,
            rangeslider=dict(visible=True) # 하단 슬라이더로 기간 조절 가능
        ),
        yaxis=dict(
            tickvals=[0, 1],
            ticktext=["금통위", "만기일"],
            range=[-0.5, 1.5],
            fixedrange=True
        ),
        template="plotly_white",
        height=500,
        margin=dict(l=100, r=50, t=100, b=100),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    # 차트 즉시 표시
    fig.show()

if __name__ == "__main__":
    create_timeline_chart()
