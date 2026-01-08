import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy.optimize import newton

# 1. 입력 데이터 설정
market_tenors = np.array([1, 2, 3, 5])
market_rates = np.array([0.02, 0.03, 0.04, 0.05])
dt = 0.25

def get_df(t, nodes, forwards):
    """t시점의 Discount Factor 계산"""
    if t <= 0: return 1.0
    total_integral = 0
    prev_node = 0
    for i in range(len(forwards)):
        node = nodes[i]
        f = forwards[i]
        if t <= node:
            total_integral += f * (t - prev_node)
            return np.exp(-total_integral)
        else:
            total_integral += f * (node - prev_node)
            prev_node = node
    total_integral += forwards[-1] * (t - prev_node)
    return np.exp(-total_integral)

# --- [Phase 1] 부트스트래핑 및 데이터 축적 (DataFrame) ---
print("Phase 1: Bootstrapping and accumulating data...")
rows_list = []
temp_forwards = []

for step_idx in range(len(market_tenors)):
    target_tenor = market_tenors[step_idx]
    swap_rate = market_rates[step_idx]
    iter_context = {'count': 0}
    
    def objective(f_current):
        f_val = float(f_current)
        current_fwd_set = temp_forwards + [f_val]
        current_nodes = market_tenors[:len(current_fwd_set)].tolist()
        
        payment_times = np.arange(dt, target_tenor + 1e-9, dt)
        fixed_leg = sum([swap_rate * dt * get_df(t, current_nodes, current_fwd_set) for t in payment_times])
        df_end = get_df(target_tenor, current_nodes, current_fwd_set)
        fixed_bond = fixed_leg + (1.0 * df_end)
        
        iter_context['count'] += 1
        
        rows_list.append({
            'Step': step_idx + 1,
            'Iteration': iter_context['count'],
            'Target_Tenor': target_tenor,
            'Fwd_Attempted': f_val,
            'Fixed_Bond_Value': fixed_bond,
            'All_Fwds': list(current_fwd_set)
        })
        return fixed_bond - 1.0

    f_sol = newton(objective, market_rates[step_idx], tol=1e-7)
    temp_forwards.append(f_sol)

bootstrapping_df = pd.DataFrame(rows_list)

# --- [Phase 2] Plotly 인터랙티브 시각화 생성 ---
print("Phase 2: Creating Interactive HTML...")

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
    nodes = market_tenors[:len(fwds)].tolist()
    
    # 1. Fwd Data
    fwd_x = [0] + nodes
    fwd_y = fwds + [fwds[-1]]
    
    # 텍스트 위치 계산 (각 구간의 중앙)
    text_x = []
    prev_node = 0
    for node in nodes:
        text_x.append((prev_node + node) / 2)
        prev_node = node
    text_y = fwds
    
    # 2. DF Data (항상 0부터 시작)
    t_range = np.linspace(0, row['Target_Tenor'], 100)
    df_y = [get_df(t, nodes, fwds) for t in t_range]
    
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
        go.Scatter(x=t_range, y=df_y, name='Discount Factor', line=dict(color='blue', width=3)),
        go.Bar(x=pay_times, y=coupons, name='Coupon', marker_color='orange', opacity=0.7, width=0.1),
        go.Bar(x=[pay_times[-1]], y=[1.0], name='Principal', marker_color='red', opacity=0.5, width=0.15)
    ]
    
    frames.append(go.Frame(
        data=frame_traces,
        name=f"frame{i}",
        layout=go.Layout(
            title_text=f"<b>Step {int(row['Step'])} | Iteration {int(row['Iteration'])} | fwd rate: {row['Fwd_Attempted']:.4%} | Bond Value: {row['Fixed_Bond_Value']:.6f} (error = {row['Fixed_Bond_Value']-1.0:.6f})</b>"
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
    height=1000,
    title="<b>IRS Bootstrapping Interactive Process (Left Click: Prev, Right Click: Next)</b>",
    updatemenus=[], # Step 드롭다운 삭제
    sliders=[]      # 하단 슬라이더 삭제
)

# 시간축(X축) 일치 및 0부터 5.5로 고정
fig.update_xaxes(range=[0, 5.5], row=1, col=1)
fig.update_xaxes(range=[0, 5.5], row=2, col=1)
fig.update_xaxes(range=[0, 5.5], row=3, col=1)

# Fwd Y축 % 포맷 적용
fig.update_yaxes(tickformat=".1%", range=[0.01, 0.08], row=1, col=1)
fig.update_yaxes(range=[0.7, 1.05], row=2, col=1)

# Cash Flow 축 라벨 및 범위 설정
fig.update_yaxes(title_text="이표금액", range=[0, 0.03], row=3, col=1, secondary_y=False)
fig.update_yaxes(title_text="원금", range=[0, 1.2], tickvals=[1], ticktext=["1"], row=3, col=1, secondary_y=True)

# HTML 저장
fig.frames = frames

# JavaScript 삽입: 내비게이션 및 제어 버튼
html_content = fig.to_html(include_plotlyjs=True, full_html=True)

custom_js = '''
<style>
    .control-btn {
        position: fixed;
        right: 30px;
        padding: 12px 20px;
        font-size: 16px;
        font-weight: bold;
        color: white;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        z-index: 1000;
        width: 120px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
    }
    #start-btn { top: 40%; background-color: #4CAF50; }
    #stop-btn { top: 47%; background-color: #f44336; }
    #prev-btn { top: 54%; background-color: #2196F3; }
    .control-btn:hover { opacity: 0.8; }
    .instruction-text {
        position: fixed;
        right: 30px;
        top: 62%;
        font-size: 14px;
        color: #555;
        text-align: right;
        z-index: 1000;
        line-height: 1.6;
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

# </body> 태그 앞에 커스텀 JavaScript 삽입
html_content = html_content.replace('</body>', custom_js + '</body>')

with open('irs_bootstrapping_interactive.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("\nSuccess! Updated Interactive HTML saved as 'irs_bootstrapping_interactive.html'.")
print("Features:")
print("  - Any plot click: Next Iteration")
print("  - '전단계로' button: Previous Iteration")
print("  - Start/Stop buttons on the right middle to control playback")
print("  - Smooth transition effects removed (transition duration set to 0)")
