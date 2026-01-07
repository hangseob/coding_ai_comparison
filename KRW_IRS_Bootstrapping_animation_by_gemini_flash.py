import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from scipy.optimize import fsolve

# 1. 입력 데이터 설정
market_tenors = np.array([1, 2, 3, 5])
market_rates = np.array([0.02, 0.03, 0.04, 0.05])
dt = 0.25  # 분기별 지급 가정

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
    # 마지막 노드 이후는 마지막 선도금리 유지 가정
    total_integral += forwards[-1] * (t - prev_node)
    return np.exp(-total_integral)

# 애니메이션을 위한 초기 설정
fig = plt.figure(figsize=(10, 10))
gs = fig.add_gridspec(3, 1)
ax1 = fig.add_subplot(gs[0, 0]) # Instantaneous Forward Rate
ax2 = fig.add_subplot(gs[1, 0]) # Discount Factor Curve
ax3 = fig.add_subplot(gs[2, 0]) # IRS Cash Flow
nodes = market_tenors.tolist()

# 그래프 초기화 함수
def init():
    # 1. Instantaneous Forward Rate (Top)
    ax1.set_xlim(0, 5.5)
    ax1.set_ylim(0.01, 0.08)
    ax1.set_title('Instantaneous Forward Rate')
    ax1.set_xlabel('Years')
    ax1.grid(True)
    
    # 2. Discount Factor Curve (Middle)
    ax2.set_xlim(0, 5.5)
    ax2.set_ylim(0.7, 1.05)
    ax2.set_title('Discount Factor Curve')
    ax2.set_xlabel('Years')
    ax2.grid(True)

    # 3. IRS Cash Flow (Bottom)
    ax3.set_xlim(0, 5.5)
    ax3.set_ylim(0, 1.1) # 원금 1.0을 포함하기 위해 범위를 1.1까지 확대
    ax3.set_title('IRS Fixed Leg Cash Flows (Coupon + Principal)')
    ax3.set_xlabel('Years')
    ax3.set_ylabel('Payment Amount')
    ax3.grid(True)
    return []

# 각 프레임마다 호출될 함수 (i는 0, 1, 2, 3...)
def update(frame):
    ax1.clear()
    ax2.clear()
    ax3.clear()
    init()
    
    # 현재 프레임까지의 부트스트래핑 수행
    current_forwards = []
    for i in range(frame + 1):
        target_tenor = market_tenors[i]
        swap_rate = market_rates[i]
        
        def objective(f_current):
            temp_forwards = current_forwards + [f_current]
            temp_nodes = nodes[:i+1]
            payment_times = np.arange(dt, target_tenor + 1e-9, dt)
            fixed_leg = sum([swap_rate * dt * get_df(t, temp_nodes, temp_forwards) for t in payment_times])
            df_end = get_df(target_tenor, temp_nodes, temp_forwards)
            return fixed_leg - (1 - df_end)
        
        f_sol = fsolve(objective, market_rates[i])[0]
        current_forwards.append(f_sol)
    
    # 1. Forward Rate Step 업데이트 (Top)
    fwd_x = [0] + nodes[:frame+1]
    fwd_y = current_forwards + [current_forwards[-1]]
    ax1.step(fwd_x, fwd_y, where='post', color='green', lw=2, label='Step-by-step Fwd')
    ax1.legend(loc='upper left')

    # 2. DF Curve 업데이트 (Middle)
    plot_times = np.linspace(0, market_tenors[frame], 100)
    df_values = [get_df(t, nodes[:frame+1], current_forwards) for t in plot_times]
    ax2.plot(plot_times, df_values, color='blue', lw=2, label='Current DF Curve')
    ax2.scatter(market_tenors[:frame+1], [get_df(t, nodes[:frame+1], current_forwards) for t in market_tenors[:frame+1]], 
               color='red', zorder=5, label='Solved Nodes')
    ax2.legend(loc='lower left')

    # 3. IRS Cash Flow 업데이트 (Bottom)
    current_swap_rate = market_rates[frame]
    current_tenor = market_tenors[frame]
    payment_times = np.arange(dt, current_tenor + 1e-9, dt)
    
    # 고정금리 이자 (Coupon)
    cash_flows = [current_swap_rate * dt] * len(payment_times)
    # 마지막에 원금(1.0) 추가
    cash_flows[-1] += 1.0
    
    # 이자 부분과 원금 부분을 색상으로 구분하여 표시하기 위해 bar를 두 번 그릴 수도 있지만, 
    # 요청하신 대로 "붙여서" 나타내기 위해 하나의 bar로 그리되 색상을 강조하겠습니다.
    ax3.bar(payment_times, cash_flows, width=0.1, color='orange', alpha=0.7, label=f'{current_tenor}Y Swap CF')
    
    # 원금 부분만 별도로 강조 (마지막 막대 위에 텍스트 표시 등 가능)
    ax3.text(payment_times[-1], cash_flows[-1], ' Principal', va='bottom', ha='center', fontweight='bold')
    
    ax3.set_title(f'IRS Cash Flows for {current_tenor}Y Swap (Rate: {current_swap_rate*100:.1f}%)')
    ax3.legend(loc='upper right')
    
    return []

# 애니메이션 생성
ani = FuncAnimation(fig, update, frames=len(market_tenors), init_func=init, blit=False, repeat=True, interval=1500)

# 일시정지 기능을 위한 상태 변수 및 함수
is_paused = False

def on_click(event):
    global is_paused
    if is_paused:
        ani.event_source.start()
        is_paused = False
        print("Animation Resumed")
    else:
        ani.event_source.stop()
        is_paused = True
        print("Animation Paused (Click again to resume)")

# 마우스 클릭 이벤트 연결
fig.canvas.mpl_connect('button_press_event', on_click)

plt.tight_layout()
plt.show()

