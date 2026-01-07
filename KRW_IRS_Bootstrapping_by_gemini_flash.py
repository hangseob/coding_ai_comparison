import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import fsolve

# 1. 입력 데이터 설정
market_tenors = np.array([1, 2, 3, 5])
market_rates = np.array([0.02, 0.03, 0.04, 0.05])
freq = 4  # Quarterly (분기 지급)
dt = 0.25  # Act/365 기준 1년을 365일로 보고 분기별 0.25년 가정


def get_df(t, nodes, forwards):
    """지정된 노드와 piecewise constant 선도금리를 이용해 t시점의 DF 계산"""
    if t <= 0: return 1.0
    total_integral = 0
    prev_node = 0
    for i in range(len(nodes)):
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


# 2. 부트스트래핑 수행
nodes = market_tenors.tolist()
forwards = []

for i in range(len(market_tenors)):
    target_tenor = market_tenors[i]
    swap_rate = market_rates[i]


    def objective(f_current):
        temp_forwards = forwards + [f_current]
        temp_nodes = nodes[:i + 1]

        # Fixed Leg: S * dt * sum(DF_i)
        payment_times = np.arange(dt, target_tenor + 1e-9, dt)
        fixed_leg = sum([swap_rate * dt * get_df(t, temp_nodes, temp_forwards) for t in payment_times])

        # Floating Leg: 1 - DF_end (No-arbitrage assumption)
        df_end = get_df(target_tenor, temp_nodes, temp_forwards)
        floating_leg = 1 - df_end

        return fixed_leg - floating_leg


    f_sol = fsolve(objective, swap_rate)[0]
    forwards.append(f_sol)

# 3. 결과 데이터 정리
results_df = pd.DataFrame({
    'Tenor (Year)': market_tenors,
    'Inst. Forward Rate': forwards,
    'Discount Factor': [get_df(t, nodes, forwards) for t in market_tenors]
})

print(results_df)

# 4. 차트 생성
plt.figure(figsize=(12, 5))

# Discount Factor Chart
plt.subplot(1, 2, 1)
plot_times = np.linspace(0, 5, 100)
plt.plot(plot_times, [get_df(t, nodes, forwards) for t in plot_times], color='blue', label='DF Curve')
plt.scatter(market_tenors, results_df['Discount Factor'], color='red', label='Market Nodes')
plt.title('Discount Factor Curve')
plt.xlabel('Years');
plt.ylabel('DF');
plt.grid(True);
plt.legend()

# Forward Rate Chart (Step)
plt.subplot(1, 2, 2)
fwd_steps_x = [0] + nodes
fwd_steps_y = forwards + [forwards[-1]]
plt.step(fwd_steps_x, fwd_steps_y, where='post', color='green', label='Piecewise Const Fwd')
plt.title('Instantaneous Forward Rate')
plt.xlabel('Years');
plt.ylabel('Rate');
plt.grid(True);
plt.legend()

plt.tight_layout()
plt.show()