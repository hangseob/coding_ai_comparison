import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import brentq

# ---------------------------------------------------------
# 1. 데이터 설정
# ---------------------------------------------------------
# Tenor (T)와 Market Rate (S)
market_data = {
    1: 0.02,
    2: 0.03,
    3: 0.04,
    5: 0.05
}

# ---------------------------------------------------------
# 2. Bootstrapping 함수 정의
# ---------------------------------------------------------
def bootstrap_irs(market_data):
    # 결과를 저장할 딕셔너리
    # dfs: {Tenor: Discount Factor}
    # forwards: {Tenor_End: Constant Forward Rate from Prev Tenor}
    dfs = {0: 1.0} 
    forwards = {} 
    
    sorted_tenors = sorted(market_data.keys())
    
    prev_tenor = 0
    
    for T in sorted_tenors:
        swap_rate = market_data[T]
        
        # 이전까지 확정된 현금흐름의 가치 (Known Fixed Leg PV)
        # 1년부터 prev_tenor까지의 PV 계산
        pv_existing = 0
        for t in range(1, prev_tenor + 1):
            if t in dfs:
                pv_existing += swap_rate * 1.0 * dfs[t] # delta_time = 1.0 가정
        
        # 목표 함수: 미지수 f(구간 선도금리)를 찾기 위한 함수
        # 3년->5년의 경우 4년차 현금흐름도 고려해야 함
        def objective_function(f):
            # 구간 내의 모든 현금흐름 시점(Payment Dates) 생성
            # 예: prev=3, T=5 이면 -> pay_dates = [4, 5]
            pay_dates = list(range(prev_tenor + 1, T + 1))
            
            temp_dfs = {}
            pv_new = 0
            
            # 각 시점별 DF 계산: DF(t) = DF(prev) * exp(-f * (t - prev))
            for t in pay_dates:
                df_val = dfs[prev_tenor] * np.exp(-f * (t - prev_tenor))
                temp_dfs[t] = df_val
                pv_new += swap_rate * 1.0 * df_val # Fixed Leg Payment
            
            # Floating Leg PV = 1 - DF(Last)
            # Fixed Leg PV = pv_existing + pv_new
            # Swap Equation: Fixed PV - Floating PV = 0
            df_final = temp_dfs[T]
            floating_pv = 1.0 - df_final
            fixed_pv = pv_existing + pv_new
            
            return fixed_pv - floating_pv

        # Root Finding (f 찾기)
        # 금리 범위를 -10% ~ 20% 사이로 가정하고 해를 찾음
        implied_f = brentq(objective_function, -0.1, 0.2)
        
        # 결과 저장
        for t in range(prev_tenor + 1, T + 1):
            dfs[t] = dfs[prev_tenor] * np.exp(-implied_f * (t - prev_tenor))
            # f는 구간 전체에 적용되므로 해당 구간의 대표 forward로 저장
            # 시각화 편의를 위해 구간별로 저장
            forwards[t] = implied_f 
            
        prev_tenor = T

    return dfs, forwards

# ---------------------------------------------------------
# 3. 실행 및 데이터 프레임 생성
# ---------------------------------------------------------
dfs, forwards = bootstrap_irs(market_data)

# 결과 정리
results = []
for t in sorted(dfs.keys()):
    if t == 0: continue
    results.append({
        "Tenor (Year)": t,
        "Discount Factor": dfs[t],
        "Inst. Forward Rate": forwards.get(t, np.nan)
    })

df_result = pd.DataFrame(results)
print("=== Bootstrapping Result ===")
print(df_result.to_string(index=False, formatters={
    'Discount Factor': '{:.6f}'.format,
    'Inst. Forward Rate': '{:.4%}'.format
}))

# ---------------------------------------------------------
# 4. 차트 그리기
# ---------------------------------------------------------
# 데이터 준비
t_values = df_result["Tenor (Year)"]
df_values = df_result["Discount Factor"]
fwd_values = df_result["Inst. Forward Rate"]

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

# Chart 1: Discount Factors
ax1.plot(t_values, df_values, marker='o', linestyle='-', color='b')
ax1.set_title("Discount Factors (Bootstrapped)")
ax1.set_xlabel("Tenor (Years)")
ax1.set_ylabel("Discount Factor")
ax1.grid(True)
for x, y in zip(t_values, df_values):
    ax1.text(x, y, f"{y:.4f}", ha='left', va='bottom')

# Chart 2: Instantaneous Forward Rates (Step Plot)
# Step plot logic: Forward rate is constant FROM prev TO current
# We need to construct step data explicitly for plotting
step_x = [0] + list(t_values)
step_y = [fwd_values.iloc[0]] + list(fwd_values) # Start with first fwd

ax2.step(step_x, step_y, where='pre', color='r', linewidth=2)
ax2.set_title("Instantaneous Forward Rates (Piecewise Constant)")
ax2.set_xlabel("Time (Years)")
ax2.set_ylabel("Rate")
ax2.grid(True)
ax2.set_ylim(0, max(fwd_values)*1.2)

# Annotate values
for i, txt in enumerate(fwd_values):
    ax2.text(t_values[i]-0.5, txt, f"{txt:.2%}", ha='center', va='bottom', color='darkred', fontweight='bold')

plt.tight_layout()
plt.show()