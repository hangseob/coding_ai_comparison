import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
from matplotlib.animation import FuncAnimation
from scipy.optimize import newton # fsolve 대신 newton 사용

# ffmpeg 경로 수동 설정
plt.rcParams['animation.ffmpeg_path'] = r'C:\ffmpeg\bin\ffmpeg.exe'

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
fig = plt.figure(figsize=(8, 9)) # 높이를 조금 더 키워서 공간 확보
fig.subplots_adjust(hspace=0.7, right=0.85, top=0.92, bottom=0.08) # 간격을 더 넓게 조정
gs = fig.add_gridspec(3, 1)
ax1 = fig.add_subplot(gs[0, 0]) # Instantaneous Forward Rate
ax2 = fig.add_subplot(gs[1, 0]) # Discount Factor Curve
ax3 = fig.add_subplot(gs[2, 0]) # IRS Cash Flow (Left: Coupon)
ax3_twin = ax3.twinx()          # IRS Cash Flow (Right: Principal)
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
    ax3.set_ylim(0, 0.03) # 연 10% = 0.025 기준
    ax3.set_ylabel('Coupon (Left)', color='orange')
    ax3.set_title('IRS Fixed Leg Cash Flows (Dual Axis)')
    ax3.grid(True)

    ax3_twin.set_ylim(0, 1.2)
    ax3_twin.set_ylabel('Principal (Right)', color='red', labelpad=15)
    ax3_twin.yaxis.set_label_position("right")
    ax3_twin.yaxis.tick_right()
    return []

# 부트스트래핑 과정을 기록할 리스트
bootstrapping_log = []
final_forwards_for_animation = [] # 애니메이션에서 사용할 미리 계산된 결과

# xlwings 설정: 새로운 엑셀 북을 열거나 기존 북 연결
try:
    wb = xw.Book() # 새 통합 문서
    sheet = wb.sheets[0]
    sheet.name = "Bootstrapping Process"
    # 헤더 작성
    headers = ['Step', 'Iteration', 'Target_Tenor', 'Fwd_Attempted', 'Fixed_Bond_Value']
    sheet.range('A1').value = headers
except Exception as e:
    print(f"xlwings initialization error: {e}")
    wb = None

# --- [Phase 1] 부트스트래핑 및 DataFrame 데이터 축적 ---
print("Phase 1: Bootstrapping and accumulating data into DataFrame...")

# 데이터를 먼저 리스트에 모은 후 마지막에 DataFrame으로 변환 (경고 방지 및 속도 향상)
rows_list = []
temp_forwards = []

for step_idx in range(len(market_tenors)):
    target_tenor = market_tenors[step_idx]
    swap_rate = market_rates[step_idx]
    iter_context = {'count': 0}
    
    def objective(f_current):
        # newton은 f_current를 항상 스칼라로 전달함
        f_val = float(f_current)
        current_fwd_set = temp_forwards + [f_val]
        current_nodes = market_tenors[:len(current_fwd_set)].tolist()
        
        payment_times = np.arange(dt, target_tenor + 1e-9, dt)
        fixed_leg = sum([swap_rate * dt * get_df(t, current_nodes, current_fwd_set) for t in payment_times])
        df_end = get_df(target_tenor, current_nodes, current_fwd_set)
        fixed_bond = fixed_leg + (1.0 * df_end)
        
        iter_context['count'] += 1
        
        # 리스트에 데이터 추가
        rows_list.append({
            'Step': step_idx + 1,
            'Iteration': iter_context['count'],
            'Target_Tenor': target_tenor,
            'Fwd_Attempted': f_val,
            'Fixed_Bond_Value': fixed_bond,
            'All_Fwds': list(current_fwd_set)
        })
        
        return fixed_bond - 1.0

    # fsolve 대신 newton 사용 (1D 최적화에 더 적합하며 중복 호출이 적음)
    f_sol = newton(objective, market_rates[step_idx], tol=1e-7)
    temp_forwards.append(f_sol)

# 최종 DataFrame 생성
bootstrapping_df = pd.DataFrame(rows_list)
print("\n--- Iteration Log Preview (First 15 rows) ---")
print(bootstrapping_df[['Step', 'Iteration', 'Fwd_Attempted', 'Fixed_Bond_Value']].head(15))
print("-------------------------------------------\n")

# --- [Phase 2] 엑셀 일괄 기록 ---
if wb is not None:
    try:
        print("Phase 2: Writing to Excel...")
        # 엑셀 기록용 컬럼만 필터링하여 출력
        excel_columns = ['Step', 'Iteration', 'Target_Tenor', 'Fwd_Attempted', 'Fixed_Bond_Value']
        sheet.range('A2').value = bootstrapping_df[excel_columns].values
        print("Excel recording completed.")
    except Exception as e:
        print(f"Excel recording error: {e}")

# --- [Phase 3] DataFrame 데이터를 이용한 수동 애니메이션 ---
current_frame = 0
total_frames = len(bootstrapping_df)

def update_display():
    global current_frame
    if current_frame >= total_frames:
        print("End of process. Resetting to start.")
        current_frame = 0
        
    # 이전 프레임의 Fit Success 문구 제거
    fig.texts.clear()
        
    row = bootstrapping_df.iloc[current_frame]
    current_fwds = row['All_Fwds']
    current_nodes = market_tenors[:len(current_fwds)].tolist()
    
    ax1.clear()
    ax2.clear()
    ax3.clear()
    ax3_twin.clear()
    init()
    
    # 1. Forward Rate Step (Top)
    fwd_x = [0] + current_nodes
    fwd_y = current_fwds + [current_fwds[-1]]
    ax1.step(fwd_x, fwd_y, where='post', color='green', lw=2)
    ax1.set_title(f"Step {int(row['Step'])} | Iteration {int(row['Iteration'])}")
    ax1.annotate(f"Trying Fwd: {row['Fwd_Attempted']:.4%}", (0.5, 0.05), xycoords='axes fraction', ha='center', color='darkgreen', fontweight='bold')

    # 2. Discount Factor Curve (Middle)
    plot_times = np.linspace(0, row['Target_Tenor'], 100)
    df_values = [get_df(t, current_nodes, current_fwds) for t in plot_times]
    ax2.plot(plot_times, df_values, color='blue', lw=2)
    ax2.scatter(current_nodes, [get_df(t, current_nodes, current_fwds) for t in current_nodes], color='red')
    
    # 3. IRS Cash Flow (Bottom)
    current_swap_rate = market_rates[int(row['Step'])-1]
    payment_times = np.arange(dt, row['Target_Tenor'] + 1e-9, dt)
    ax3.bar(payment_times, [current_swap_rate * dt] * len(payment_times), width=0.1, color='orange', alpha=0.6)
    ax3_twin.bar(payment_times[-1], 1.0, width=0.15, color='red', alpha=0.4)
    
    ax3.set_title(f"Target: Bond Price = 1.0 | Current: {row['Fixed_Bond_Value']:.6f}")

    # Fit Success 체크
    is_last_iter = False
    if current_frame == total_frames - 1:
        is_last_iter = True
    elif current_frame < total_frames - 1 and bootstrapping_df.iloc[current_frame+1]['Step'] != row['Step']:
        is_last_iter = True
    
    if is_last_iter and abs(row['Fixed_Bond_Value'] - 1.0) < 1e-6:
        fig.text(0.5, 0.5, 'FIT SUCCESS!', fontsize=40, color='red', 
                 ha='center', va='center', fontweight='bold', alpha=0.7,
                 bbox=dict(facecolor='white', alpha=0.8, edgecolor='red', boxstyle='round,pad=0.5'))
    
    fig.canvas.draw()

def on_click(event):
    global current_frame
    current_frame += 1
    update_display()

# 마우스 클릭 이벤트 연결
fig.canvas.mpl_connect('button_press_event', on_click)

# 최초 화면 표시
update_display()

# --- [Phase 4] 애니메이션 저장 (필요 시에만 별도로 실행) ---
# 저장이 필요한 경우, 아래 코드를 별도 스크립트로 분리하여 실행하세요.
# from matplotlib.animation import FuncAnimation
# def update_for_save(frame):
#     ... (위의 update_display 로직을 frame 기반으로 수정)
# ani = FuncAnimation(fig, update_for_save, frames=total_frames, ...)
# ani.save('output.mp4', writer='ffmpeg', fps=2)

print("\n>>> Click anywhere in the graph window to advance to the next iteration. <<<")
plt.show()

