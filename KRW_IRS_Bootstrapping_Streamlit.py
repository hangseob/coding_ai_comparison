import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy.optimize import newton
import streamlit.components.v1 as components

# 페이지 설정
st.set_page_config(page_title="IRS Bootstrapping Tool", layout="wide")

st.title("📈 KRW IRS Bootstrapping Interactive Tool")
st.markdown("""
이 도구는 시장 금리를 입력받아 **Log-Linear Interpolation** 방식으로 부트스트래핑을 수행합니다.
입력 완료 후 하단의 **[Run Bootstrapping]** 버튼을 눌러주세요.
""")

# 1. 입력 섹션
st.sidebar.header("1. Market Data Input")
num_tenors = st.sidebar.number_input("만기 개수 (Number of Tenors)", min_value=1, max_value=20, value=4)

# 초기 데이터 프레임 생성
default_data = {
    "Tenor (Year)": [1.0, 2.0, 3.0, 5.0] + [0.0] * (num_tenors - 4) if num_tenors >= 4 else [1.0, 2.0, 3.0, 5.0][:num_tenors],
    "Market Rate (%)": [10.0, 15.0, 20.0, 30.0] + [0.0] * (num_tenors - 4) if num_tenors >= 4 else [10.0, 15.0, 20.0, 30.0][:num_tenors]
}
df_input = pd.DataFrame(default_data)

st.subheader("Market Rates Table")
edited_df = st.data_editor(df_input, num_rows="fixed", use_container_width=True)

# 계산 설정
dt = 0.25

def get_df(t, nodes, dfs):
    """t시점의 Discount Factor 계산 (Linear on Log DF)"""
    if t <= 0: return 1.0
    log_dfs = np.log(dfs)
    interp_log_df = np.interp(t, nodes, log_dfs)
    return np.exp(interp_log_df)

if st.button("🚀 Run Bootstrapping", use_container_width=True):
    # 입력 데이터 정리
    market_tenors = edited_df["Tenor (Year)"].values
    market_rates = edited_df["Market Rate (%)"].values / 100.0  # %를 소수로 변환

    # --- [Phase 1] 부트스트래핑 ---
    rows_list = []
    temp_forwards = []
    temp_dfs = [1.0]
    temp_nodes = [0.0]

    for step_idx in range(len(market_tenors)):
        target_tenor = market_tenors[step_idx]
        swap_rate = market_rates[step_idx]
        iter_context = {'count': 0}
        
        def objective(f_current):
            f_val = float(f_current)
            df_prev = temp_dfs[-1]
            t_prev = temp_nodes[-1]
            df_target = df_prev * np.exp(-f_val * (target_tenor - t_prev))
            
            current_dfs = temp_dfs + [df_target]
            current_nodes = temp_nodes + [target_tenor]
            
            payment_times = np.arange(dt, target_tenor + 1e-9, dt)
            fixed_leg = sum([swap_rate * dt * get_df(t, current_nodes, current_dfs) for t in payment_times])
            fixed_bond = fixed_leg + (1.0 * current_dfs[-1])
            
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
        df_target_final = temp_dfs[-1] * np.exp(-f_sol * (target_tenor - temp_nodes[-1]))
        temp_dfs.append(df_target_final)
        temp_nodes.append(target_tenor)

    bootstrapping_df = pd.DataFrame(rows_list)
    fixed_t_grid = np.linspace(0, max(market_tenors), 101)

    # --- [Phase 2] Plotly 생성 ---
    fig = make_subplots(
        rows=3, cols=1, 
        subplot_titles=("1. Instantaneous Forward Rate (%)", "2. Discount Factor Curve", "3. IRS Cash Flows"),
        vertical_spacing=0.1,
        specs=[[{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": True}]]
    )

    frames = []
    for i in range(len(bootstrapping_df)):
        row = bootstrapping_df.iloc[i]
        fwds, dfs, nodes = row['All_Fwds'], row['All_Dfs'], row['All_Nodes']
        
        # 1. Fwd Data
        current_market_nodes = market_tenors[:len(fwds)].tolist()
        fwd_x = [0] + current_market_nodes
        fwd_y = fwds + [fwds[-1]]
        
        text_x = []
        p_node = 0
        for n in current_market_nodes:
            text_x.append((p_node + n) / 2)
            p_node = n
        
        # 2. DF Data
        t_disp = fixed_t_grid[fixed_t_grid <= row['Target_Tenor'] + 1e-9]
        df_y = [get_df(t, nodes, dfs) for t in t_disp]
        
        # 3. Cash Flow
        pay_times = np.arange(dt, row['Target_Tenor'] + 1e-9, dt)
        coupons = [market_rates[int(row['Step'])-1] * dt] * len(pay_times)
        
        frame_traces = [
            go.Scatter(x=fwd_x, y=fwd_y, line_shape='hv', name='Fwd Rate', line=dict(color='green', width=3), mode='lines'),
            go.Scatter(x=text_x, y=fwds, mode='text', text=[f"{v:.4%}" for v in fwds], textposition="top center", showlegend=False),
            go.Scatter(x=t_disp, y=df_y, name='Discount Factor', line=dict(color='blue', width=3)),
            go.Bar(x=pay_times, y=coupons, name='Coupon', marker_color='orange', opacity=0.7, width=0.1),
            go.Bar(x=[pay_times[-1]], y=[1.0], name='Principal', marker_color='red', opacity=0.5, width=0.15)
        ]
        
        c_max_fwd = max(max(fwds), 0.08)
        c_min_df = min(min(dfs), 0.7)

        frames.append(go.Frame(
            data=frame_traces, name=f"frame{i}",
            layout=go.Layout(
                title_text=f"<b>Step {int(row['Step'])} | Iteration {int(row['Iteration'])} | fwd rate: {row['Fwd_Attempted']:.4%} | Bond Value: {row['Fixed_Bond_Value']:.6f} (error = {row['Fixed_Bond_Value']-1.0:.6f})</b>",
                yaxis=dict(range=[0, c_max_fwd * 1.1]),
                yaxis2=dict(range=[c_min_df * 0.95, 1.05])
            )
        ))

    # 초기 빈 트레이스
    fig.add_trace(go.Scatter(x=[], y=[], line_shape='hv', name='Fwd Rate', line=dict(color='green', width=3), mode='lines'), row=1, col=1)
    fig.add_trace(go.Scatter(x=[], y=[], mode='text', showlegend=False), row=1, col=1)
    fig.add_trace(go.Scatter(x=[], y=[], name='Discount Factor', line=dict(color='blue', width=3)), row=2, col=1)
    fig.add_trace(go.Bar(x=[], y=[], name='Coupon', marker_color='orange', opacity=0.7, width=0.1), row=3, col=1, secondary_y=False)
    fig.add_trace(go.Bar(x=[], y=[], name='Principal', marker_color='red', opacity=0.5, width=0.15), row=3, col=1, secondary_y=True)

    empty_frame = go.Frame(data=[go.Scatter(x=[], y=[]), go.Scatter(x=[], y=[]), go.Scatter(x=[], y=[]), go.Bar(x=[], y=[]), go.Bar(x=[], y=[])], name="frame0", layout=go.Layout(title_text=""))
    for i, f in enumerate(frames): f.name = f"frame{i+1}"
    frames.insert(0, empty_frame)
    fig.frames = frames

    fig.update_layout(height=800, title="<b>IRS Bootstrapping Interactive Process</b>", updatemenus=[], sliders=[])
    fig.update_xaxes(range=[0, max(market_tenors)*1.1], row=1, col=1)
    fig.update_xaxes(range=[0, max(market_tenors)*1.1], row=2, col=1)
    fig.update_xaxes(range=[0, max(market_tenors)*1.1], row=3, col=1)
    fig.update_yaxes(tickformat=".1%", range=[0, 0.08], row=1, col=1)
    fig.update_yaxes(range=[0.7, 1.05], row=2, col=1)
    fig.update_yaxes(title_text="이표금액", range=[0, max(market_rates)*dt*1.5], row=3, col=1, secondary_y=False)
    fig.update_yaxes(title_text="원금", range=[0, 1.2], tickvals=[1], ticktext=["1"], row=3, col=1, secondary_y=True)

    # HTML 변환 및 JS 삽입
    html_content = fig.to_html(include_plotlyjs=True, full_html=True)
    
    custom_js = '''
    <style>
        .control-btn { position: fixed; right: 20px; padding: 10px 15px; font-size: 14px; font-weight: bold; color: white; border: none; border-radius: 6px; cursor: pointer; z-index: 1000; width: 100px; box-shadow: 2px 2px 5px rgba(0,0,0,0.2); }
        #start-btn { top: 35%; background-color: #4CAF50; }
        #stop-btn { top: 41%; background-color: #f44336; }
        #prev-btn { top: 47%; background-color: #2196F3; }
        .instruction-text { position: fixed; right: 20px; top: 54%; font-size: 12px; color: #555; text-align: right; z-index: 1000; line-height: 1.5; font-weight: bold; }
    </style>
    <button id="start-btn" class="control-btn">시작 (Start)</button>
    <button id="stop-btn" class="control-btn">정지 (Stop)</button>
    <button id="prev-btn" class="control-btn">전단계로</button>
    <div class="instruction-text">화면 클릭시 다음 단계로</div>
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
            Plotly.animate(plotDiv, ['frame' + currentFrame], { frame: {duration: 0, redraw: true}, transition: {duration: 0}, mode: 'immediate' });
        }
        document.addEventListener('DOMContentLoaded', function() {
            var plotDiv = document.querySelector('.plotly-graph-div');
            if (plotDiv) {
                if (window.Plotly && Plotly.Animations) Plotly.Animations.terminate(plotDiv);
                setTimeout(function() { goToFrame(0, plotDiv); }, 100);
            }
            document.getElementById('start-btn').addEventListener('click', function(e) {
                e.stopPropagation();
                if (!isPlaying) {
                    isPlaying = true;
                    playInterval = setInterval(function() { goToFrame(currentFrame + 1, plotDiv); }, 500);
                }
            });
            document.getElementById('stop-btn').addEventListener('click', function(e) {
                e.stopPropagation();
                if (isPlaying) { clearInterval(playInterval); isPlaying = false; }
            });
            document.getElementById('prev-btn').addEventListener('click', function(e) {
                e.stopPropagation();
                if (isPlaying) { clearInterval(playInterval); isPlaying = false; }
                goToFrame(currentFrame - 1, plotDiv);
            });
            if (plotDiv) {
                plotDiv.addEventListener('click', function(e) {
                    if (isPlaying) { clearInterval(playInterval); isPlaying = false; return; }
                    if (e.target.closest('.modebar') || e.target.closest('.control-btn')) return;
                    goToFrame(currentFrame + 1, plotDiv);
                });
            }
        });
    </script>
    '''
    html_content = html_content.replace('</body>', custom_js + '</body>')
    
    # Streamlit에 HTML 출력
    components.html(html_content, height=850, scrolling=True)
    st.success("Bootstrapping complete! Use the interactive chart above.")

