import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy.optimize import newton

# 1. 입력 데이터 설정
market_tenors = np.array([1, 2, 3, 5])
market_rates = np.array([0.02, 0.03, 0.04, 0.05])
dt = 0.25  # Act/365 기준 1년을 365일로 보고 분기별 0.25년 가정

def get_df(t, nodes, dfs):
    """t시점의 Discount Factor 계산 (Linear on Log DF)"""
    if t <= 0: return 1.0
    # nodes와 dfs는 0시점(t=0, DF=1.0)을 포함해야 함
    log_dfs = np.log(dfs)
    interp_log_df = np.interp(t, nodes, log_dfs)
    return np.exp(interp_log_df)

# --- [Phase 1] 부트스트래핑 및 데이터 축적 (DataFrame) ---
print("Phase 1: Bootstrapping and accumulating data...")
rows_list = []
temp_forwards = []
temp_dfs = [1.0]   # t=0일 때 DF=1.0
temp_nodes = [0.0] # t=0 노드 포함

for step_idx in range(len(market_tenors)):
    target_tenor = market_tenors[step_idx]
    swap_rate = market_rates[step_idx]
    iter_context = {'count': 0}
    
    def objective(f_current):
        f_val = float(f_current)
        # 현재 시도하는 f_val에 따른 target_tenor에서의 DF 계산
        df_prev = temp_dfs[-1]
        t_prev = temp_nodes[-1]
        df_target = df_prev * np.exp(-f_val * (target_tenor - t_prev))
        
        current_dfs = temp_dfs + [df_target]
        current_nodes = temp_nodes + [target_tenor]
        
        payment_times = np.arange(dt, target_tenor + 1e-9, dt)
        # Linear on Log DF 방식으로 중간 DF들 계산
        fixed_leg = sum([swap_rate * dt * get_df(t, current_nodes, current_dfs) for t in payment_times])
        df_end = current_dfs[-1]
        fixed_bond = fixed_leg + (1.0 * df_end)
        
        iter_context['count'] += 1
        
        rows_list.append({
            'Step': step_idx + 1,
            'Iteration': iter_context['count'],
            'Target_Tenor': target_tenor,
            'Fwd_Attempted': f_val,
            'Fixed_Bond_Value': fixed_bond,
            'All_Fwds': temp_forwards + [f_val],
            'All_Dfs': list(current_dfs),
            'All_Nodes': list(current_nodes)
        })
        return fixed_bond - 1.0

    f_sol = newton(objective, market_rates[step_idx], tol=1e-7)
    temp_forwards.append(f_sol)
    # 확정된 DF와 노드 추가
    df_prev = temp_dfs[-1]
    t_prev = temp_nodes[-1]
    temp_dfs.append(df_prev * np.exp(-f_sol * (target_tenor - t_prev)))
    temp_nodes.append(target_tenor)

bootstrapping_df = pd.DataFrame(rows_list)

# --- [Phase 2] Plotly 인터랙티브 시각화 준비 ---
# 엑셀과 동일한 시간 그리드 생성 (0.05 간격 고정)
fixed_t_grid = np.linspace(0, max(market_tenors), 101)

# --- [Phase 3] 엑셀 데이터 저장 ---
print("Phase 3: Saving data to Excel...")
final_row = bootstrapping_df.iloc[-1]
final_nodes = final_row['All_Nodes']
final_dfs = final_row['All_Dfs']

# 차트에 그려지는 곡선 포인트 (100개) 생성
t_fine = np.linspace(0, max(market_tenors), 101)
df_fine = [get_df(t, final_nodes, final_dfs) for t in t_fine]
log_df_fine = [np.log(d) for d in df_fine]

curve_export_df = pd.DataFrame({
    'Time(T)': t_fine,
    'Discount_Factor': df_fine,
    'Log_DF': log_df_fine
})

# 리스트 형태의 컬럼은 엑셀 저장 시 문자열로 변환
export_iterations_df = bootstrapping_df.copy()
for col in ['All_Fwds', 'All_Dfs', 'All_Nodes']:
    export_iterations_df[col] = export_iterations_df[col].apply(lambda x: str(x))

try:
    with pd.ExcelWriter('bootstrapping_data.xlsx', engine='openpyxl') as writer:
        export_iterations_df.to_excel(writer, sheet_name='Iteration_History', index=False)
        curve_export_df.to_excel(writer, sheet_name='Final_DF_Curve', index=False)
    print("Excel file 'bootstrapping_data.xlsx' has been created.")
except Exception as e:
    print(f"Excel save failed: {e}. (Make sure 'openpyxl' is installed)")

# --- [Phase 4] Plotly 인터랙티브 시각화 생성 ---
print("Phase 4: Creating Interactive HTML...")

fig = make_subplots(
    rows=3, cols=1, 
    subplot_titles=("1. Instantaneous Forward Rate (%)", "2. Discount Factor Curve", "3. IRS Cash Flows"),
    vertical_spacing=0.1,
    specs=[[{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": True}]]
)

# 각 Iteration을 프레임으로 추가
frames = []
for i in range(len(bootstrapping_df)):
    row = bootstrapping_df.iloc[i]
    fwds = row['All_Fwds']
    dfs = row['All_Dfs']
    nodes = row['All_Nodes']
    
    # 1. Fwd Data
    current_market_nodes = market_tenors[:len(fwds)].tolist()
    fwd_x = [0] + current_market_nodes
    fwd_y = fwds + [fwds[-1]]
    
    # 텍스트 위치 계산 (각 구간의 중앙)
    text_x = []
    prev_node = 0
    for node in current_market_nodes:
        text_x.append((prev_node + node) / 2)
        prev_node = node
    text_y = fwds
    
    # 2. DF Data (엑셀과 동일한 고정 그리드 사용)
    # 현재 타겟 만기 이하의 노드들만 필터링하여 일관성 유지
    t_display = fixed_t_grid[fixed_t_grid <= row['Target_Tenor'] + 1e-9]
    df_y = [get_df(t, nodes, dfs) for t in t_display]
    
    # 3. Cash Flow Data
    pay_times = np.arange(dt, row['Target_Tenor'] + 1e-9, dt)
    coupons = [market_rates[int(row['Step'])-1] * dt] * len(pay_times)
    
    # 프레임별 데이터 트레이스
    frame_traces = [
        # 선만 그리는 트레이스
        go.Scatter(
            x=fwd_x, y=fwd_y, 
            line_shape='hv', 
            name='Fwd Rate', 
            line=dict(color='green', width=3),
            mode='lines'
        ),
        # 텍스트만 표시하는 트레이스 (구간 중앙)
        go.Scatter(
            x=text_x, y=text_y,
            mode='text',
            text=[f"{v:.4%}" for v in fwds],
            textposition="top center",
            showlegend=False
        ),
        go.Scatter(x=t_display, y=df_y, name='Discount Factor', line=dict(color='blue', width=3)),
        go.Bar(x=pay_times, y=coupons, name='Coupon', marker_color='orange', opacity=0.7, width=0.1),
        go.Bar(x=[pay_times[-1]], y=[1.0], name='Principal', marker_color='red', opacity=0.5, width=0.15)
    ]
    
    # Y축 범위 계산
    current_max_fwd = max(max(fwds), 0.08)
    current_min_df = min(min(dfs), 0.7)

    frames.append(go.Frame(
        data=frame_traces,
        name=f"frame{i}",
        layout=go.Layout(
            title_text=f"<b>Step {int(row['Step'])} | Iteration {int(row['Iteration'])} | fwd rate: {row['Fwd_Attempted']:.4%} | Bond Value: {row['Fixed_Bond_Value']:.6f} (error = {row['Fixed_Bond_Value']-1.0:.6f})</b>",
            yaxis=dict(range=[0, current_max_fwd * 1.1]), # Fwd 차트
            yaxis2=dict(range=[current_min_df * 0.95, 1.05]) # DCF 차트
        )
    ))

# 초기 빈 데이터 추가 (시작 시 빈 차트)
fig.add_trace(go.Scatter(x=[], y=[], line_shape='hv', name='Fwd Rate', line=dict(color='green', width=3), mode='lines'), row=1, col=1)
fig.add_trace(go.Scatter(x=[], y=[], mode='text', showlegend=False), row=1, col=1)
fig.add_trace(go.Scatter(x=[], y=[], name='Discount Factor', line=dict(color='blue', width=3)), row=2, col=1)
fig.add_trace(go.Bar(x=[], y=[], name='Coupon', marker_color='orange', opacity=0.7, width=0.1), row=3, col=1, secondary_y=False)
fig.add_trace(go.Bar(x=[], y=[], name='Principal', marker_color='red', opacity=0.5, width=0.15), row=3, col=1, secondary_y=True)

# 프레임 구성 (0번 프레임은 빈 상태 추가)
empty_frame = go.Frame(
    data=[
        go.Scatter(x=[], y=[], line_shape='hv', name='Fwd Rate', mode='lines'),
        go.Scatter(x=[], y=[], mode='text'),
        go.Scatter(x=[], y=[], name='Discount Factor'),
        go.Bar(x=[], y=[], name='Coupon'),
        go.Bar(x=[], y=[], name='Principal')
    ],
    name="frame0",
    layout=go.Layout(title_text="")
)

# 기존 프레임들의 이름을 frame1, frame2... 로 변경하고 앞에 empty_frame 삽입
for i, frame in enumerate(frames):
    frame.name = f"frame{i+1}"
frames.insert(0, empty_frame)

# 레이아웃 설정
fig.update_layout(
    height=800,  # 1000에서 800으로 축소
    title="<b>IRS Bootstrapping Interactive Process (Left Click: Prev, Right Click: Next)</b>",
    updatemenus=[], # Step 드롭다운 삭제
    sliders=[]      # 하단 슬라이더 삭제
)

# 시간축(X축) 일치 및 0부터 5.5로 고정
fig.update_xaxes(range=[0, 5.5], row=1, col=1)
fig.update_xaxes(range=[0, 5.5], row=2, col=1)
fig.update_xaxes(range=[0, 5.5], row=3, col=1)

# Fwd Y축 % 포맷 적용 및 축 범위 설정
fig.update_yaxes(tickformat=".1%", range=[0, 0.08], row=1, col=1)
fig.update_yaxes(range=[0.7, 1.05], row=2, col=1)

# Cash Flow 축 라벨 및 범위 설정
fig.update_yaxes(title_text="이표금액", range=[0, 0.03], row=3, col=1, secondary_y=False)
fig.update_yaxes(title_text="원금", range=[0, 1.2], tickvals=[1], ticktext=["1"], row=3, col=1, secondary_y=True)

# HTML 저장
fig.frames = frames

# JavaScript 삽입: 내비게이션 및 제어 버튼
html_content = fig.to_html(include_plotlyjs=True, full_html=True)

# 시장 데이터 테이블 HTML 생성
market_table_html = f'''
<div class="market-data-container">
    <table class="market-table">
        <tr>
            <th>만기 (Tenor)</th>
            {''.join([f'<td>{t}Y</td>' for t in market_tenors])}
        </tr>
        <tr>
            <th>시장금리 (Market Rate)</th>
            {''.join([f'<td>{r:.4%}</td>' for r in market_rates])}
        </tr>
    </table>
</div>
'''

custom_js = '''
<style>
    .market-data-container {
        width: 100%;
        display: flex;
        justify-content: center;
        margin-top: 20px;
        margin-bottom: 0px;
    }
    .market-table {
        border-collapse: collapse;
        width: 60%;
        font-family: sans-serif;
        box-shadow: 0 0 15px rgba(0,0,0,0.1);
        border-radius: 8px;
        overflow: hidden;
        font-size: 13px;
    }
    .market-table th, .market-table td {
        border: 1px solid #eee;
        padding: 8px 12px;
        text-align: center;
    }
    .market-table th {
        background-color: #f4f4f4;
        color: #333;
        font-weight: bold;
        width: 25%;
    }
    .market-table td {
        background-color: #ffffff;
        font-weight: 500;
    }
    .control-btn {
        position: fixed;
        right: 20px;
        padding: 10px 15px;
        font-size: 14px;
        font-weight: bold;
        color: white;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        z-index: 1000;
        width: 100px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
    }
    #start-btn { top: 35%; background-color: #4CAF50; }
    #stop-btn { top: 41%; background-color: #f44336; }
    #prev-btn { top: 47%; background-color: #2196F3; }
    .control-btn:hover { opacity: 0.8; }
    .instruction-text {
        position: fixed;
        right: 20px;
        top: 54%;
        font-size: 12px;
        color: #555;
        text-align: right;
        z-index: 1000;
        line-height: 1.5;
        font-weight: bold;
    }
</style>
<button id="start-btn" class="control-btn">시작 (Start)</button>
<button id="stop-btn" class="control-btn">정지 (Stop)</button>
<button id="prev-btn" class="control-btn">전단계로</button>
<div class="instruction-text">
    화면 클릭시 다음 단계로
</div>

<script>
    var currentFrame = 0;
    var totalFrames = ''' + str(len(frames)) + ''';
    var isPlaying = false;
    var playInterval = null;
    
    function goToFrame(f, plotDiv) {
        if (!plotDiv) return;
        currentFrame = f;
        if (currentFrame < 0) currentFrame = totalFrames - 1;
        if (currentFrame >= totalFrames) currentFrame = 0;
        
        Plotly.animate(plotDiv, ['frame' + currentFrame], {
            frame: {duration: 0, redraw: true},
            transition: {duration: 0},
            mode: 'immediate'
        });
    }

    document.addEventListener('DOMContentLoaded', function() {
        var plotDiv = document.querySelector('.plotly-graph-div');
        
        // 자동 재생 방지 및 초기 프레임 고정
        if (plotDiv) {
            // Plotly 내부 애니메이션 중단
            if (window.Plotly && Plotly.Animations) {
                Plotly.Animations.terminate(plotDiv);
            }
            // 0번 프레임(빈 화면)으로 강제 이동
            setTimeout(function() {
                goToFrame(0, plotDiv);
            }, 100);
        }

        // 시작 버튼
        document.getElementById('start-btn').addEventListener('click', function(e) {
            e.stopPropagation();
            if (!isPlaying) {
                isPlaying = true;
                playInterval = setInterval(function() {
                    goToFrame(currentFrame + 1, plotDiv);
                }, 500);
            }
        });

        // 정지 버튼
        document.getElementById('stop-btn').addEventListener('click', function(e) {
            e.stopPropagation();
            if (isPlaying) {
                clearInterval(playInterval);
                isPlaying = false;
            }
        });

        // 전단계로 버튼
        document.getElementById('prev-btn').addEventListener('click', function(e) {
            e.stopPropagation();
            if (isPlaying) {
                clearInterval(playInterval);
                isPlaying = false;
            }
            goToFrame(currentFrame - 1, plotDiv);
        });

        if (plotDiv) {
            plotDiv.addEventListener('click', function(e) {
                // 클릭 시 재생 중이면 정지
                if (isPlaying) {
                    clearInterval(playInterval);
                    isPlaying = false;
                    return;
                }

                // 버튼이나 모드바 클릭 시 무시
                if (e.target.closest('.modebar') || e.target.closest('.control-btn')) {
                    return;
                }
                
                // 화면 어디든 클릭하면 다음 프레임으로
                goToFrame(currentFrame + 1, plotDiv);
            });
        }
    });
</script>
'''

# <body> 태그 뒤에 테이블 삽입 및 </body> 태그 앞에 커스텀 JavaScript 삽입
html_content = html_content.replace('<body>', '<body>' + market_table_html)
html_content = html_content.replace('</body>', custom_js + '</body>')

with open('irs_bootstrapping_interactive.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("\nSuccess! Updated Interactive HTML saved as 'irs_bootstrapping_interactive.html'.")
print("Features:")
print("  - Any plot click: Next Iteration")
print("  - '전단계로' button: Previous Iteration")
print("  - Start/Stop buttons on the right middle to control playback")
print("  - Smooth transition effects removed (transition duration set to 0)")
