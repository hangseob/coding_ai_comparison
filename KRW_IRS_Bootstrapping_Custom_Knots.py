import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import fsolve

# 1. 환경 설정 및 시장 데이터 (예시 금리 사용)
# 만기 (Year 단위)
maturities = {
    '1D Call': 1/365,
    '3M Depo': 3/12,
    '6M IRS': 6/12,
    '9M IRS': 9/12,
    '1Y IRS': 12/12
}

# 시장 금리 (예시: 3.5% ~ 3.9%)
market_rates = {
    '1D Call': 0.0350,
    '3M Depo': 0.0360,
    '6M IRS': 0.0370,
    '9M IRS': 0.0380,
    '1Y IRS': 0.0390
}

# Jumps (Knots) 위치 (Month -> Year 변환)
# 사용자의 요청: 2개월, 5개월, 7개월, 10개월, 13개월
nodes = np.array([2, 5, 7, 10, 13]) / 12.0

def get_df(t, nodes, forwards):
    """
    Piecewise flat instantaneous forward rate를 이용해 t시점의 Discount Factor 계산
    P(0, T) = exp(- integral_0^T f(s) ds)
    """
    if t <= 0:
        return 1.0
    
    integral = 0
    prev_node = 0
    for i in range(len(nodes)):
        node = nodes[i]
        f = forwards[i]
        if t <= node:
            integral += f * (t - prev_node)
            return np.exp(-integral)
        else:
            integral += f * (node - prev_node)
            prev_node = node
    
    # 마지막 노드 이후에도 마지막 forward rate 적용
    integral += forwards[-1] * (t - prev_node)
    return np.exp(-integral)

# 2. 부트스트래핑 수행
# 각 인스트루먼트별로 순차적으로 forward rate를 구함
solved_forwards = []

# 1) 1D Call (T=1/365) -> 1st Forward (0~2M 구간)
def obj_1d(f):
    df = get_df(maturities['1D Call'], nodes[:1], [f])
    # 단리 가정: DF = 1 / (1 + r * dt)
    target_df = 1 / (1 + market_rates['1D Call'] * maturities['1D Call'])
    return df - target_df

f1 = fsolve(obj_1d, market_rates['1D Call'])[0]
solved_forwards.append(f1)

# 2) 3M Depo (T=3/12) -> 2nd Forward (2M~5M 구간)
def obj_3m(f):
    temp_fwds = solved_forwards + [f]
    df = get_df(maturities['3M Depo'], nodes[:2], temp_fwds)
    target_df = 1 / (1 + market_rates['3M Depo'] * maturities['3M Depo'])
    return df - target_df

f2 = fsolve(obj_3m, market_rates['3M Depo'])[0]
solved_forwards.append(f2)

# 3) 6M IRS (T=6/12) -> 3rd Forward (5M~7M 구간)
def obj_6m(f):
    temp_fwds = solved_forwards + [f]
    dt = 3/12 # 분기 지급 가정
    pay_times = np.array([3/12, 6/12])
    dfs = [get_df(t, nodes[:3], temp_fwds) for t in pay_times]
    fixed_leg = market_rates['6M IRS'] * dt * sum(dfs)
    floating_leg = 1 - dfs[-1]
    return fixed_leg - floating_leg

f3 = fsolve(obj_6m, market_rates['6M IRS'])[0]
solved_forwards.append(f3)

# 4) 9M IRS (T=9/12) -> 4th Forward (7M~10M 구간)
def obj_9m(f):
    temp_fwds = solved_forwards + [f]
    dt = 3/12
    pay_times = np.array([3/12, 6/12, 9/12])
    dfs = [get_df(t, nodes[:4], temp_fwds) for t in pay_times]
    fixed_leg = market_rates['9M IRS'] * dt * sum(dfs)
    floating_leg = 1 - dfs[-1]
    return fixed_leg - floating_leg

f4 = fsolve(obj_9m, market_rates['9M IRS'])[0]
solved_forwards.append(f4)

# 5) 1Y IRS (T=12/12) -> 5th Forward (10M~13M 구간)
def obj_1y(f):
    temp_fwds = solved_forwards + [f]
    dt = 3/12
    pay_times = np.array([3/12, 6/12, 9/12, 12/12])
    dfs = [get_df(t, nodes[:5], temp_fwds) for t in pay_times]
    fixed_leg = market_rates['1Y IRS'] * dt * sum(dfs)
    floating_leg = 1 - dfs[-1]
    return fixed_leg - floating_leg

f5 = fsolve(obj_1y, market_rates['1Y IRS'])[0]
solved_forwards.append(f5)

# 3. 결과 출력
results = pd.DataFrame({
    'Knot (Month)': [2, 5, 7, 10, 13],
    'Knot (Year)': nodes,
    'Inst. Forward Rate': solved_forwards
})

print("### Bootstrapping Results (Instantaneous Forwards) ###")
print(results)

# 4. 시각화
plt.figure(figsize=(12, 5))

# Forward Rate Step Plot
plt.subplot(1, 2, 1)
plot_nodes = [0] + list(nodes)
plot_fwds = [solved_forwards[0]] + solved_forwards
plt.step(plot_nodes, plot_fwds, where='post', color='red', label='Inst. Forward Rate')
plt.title('Piecewise Flat Instantaneous Forward Rate')
plt.xlabel('Year')
plt.ylabel('Rate')
plt.grid(True)
plt.legend()

# Discount Factor Plot
plt.subplot(1, 2, 2)
t_range = np.linspace(0, 1.2, 100)
df_vals = [get_df(t, nodes, solved_forwards) for t in t_range]
plt.plot(t_range, df_vals, label='Discount Factor Curve', color='blue')
plt.scatter(list(maturities.values()), 
            [get_df(t, nodes, solved_forwards) for t in maturities.values()], 
            color='black', label='Instrument Maturities')
plt.title('Discount Factor Curve')
plt.xlabel('Year')
plt.ylabel('DF')
plt.grid(True)
plt.legend()

plt.tight_layout()
plt.show()
