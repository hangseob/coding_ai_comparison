import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import fsolve

# 1. 입력 데이터 설정 (KRW IRS Rates)
tenors = np.array([1, 2, 3, 5])  # 년 단위
swap_rates = np.array([0.02, 0.03, 0.04, 0.05])  # 2%, 3%, 4%, 5%


def bootstrap_irs(tenors, rates):
    # 결과 저장용 리스트 (t=0일 때 DF=1)
    df_dict = {0: 1.0}
    forward_rates = []

    # 순차적으로 구간별 f(t) 계산
    # 구간 [t_i-1, t_i] 에서 instantaneous forward f_i는 일정함
    for i in range(len(tenors)):
        t_curr = tenors[i]
        t_prev = tenors[i - 1] if i > 0 else 0
        r_swap = rates[i]

        # 이전까지의 DF 합 계산 (Swap 고정금리 지급액 계산용)
        # 단순화를 위해 연간 1회 교환(Annual frequency) 가정
        prev_times = np.arange(1, int(t_curr))

        def objective(f_i):
            # 현재 구간의 DF 계산: DF(t) = DF(t_prev) * exp(-f_i * (t - t_prev))
            df_step = lambda t: df_dict[t_prev] * np.exp(-f_i * (t - t_prev))

            # Swap 가치 평가: sum(Swap Rate * DF(t)) = 1 - DF(T)
            # 고정금리부 합계 (Fixed Leg)
            fixed_leg = 0
            for t in range(1, int(t_curr) + 1):
                if t <= t_prev:
                    fixed_leg += r_swap * df_dict[t]
                else:
                    fixed_leg += r_swap * df_step(t)

            # Floating Leg의 합은 1 - DF(T)와 같음 (No-arbitrage)
            return fixed_leg - (1 - df_step(t_curr))

        # 구간별 선도금리 f_i 도출
        f_i_sol = fsolve(objective, 0.03)[0]
        forward_rates.append(f_i_sol)

        # 구간 내 정수 시점의 DF 저장
        for t in range(int(t_prev) + 1, int(t_curr) + 1):
            df_dict[t] = df_dict[t_prev] * np.exp(-f_i_sol * (t - t_prev))

    return df_dict, forward_rates


# 계산 실행
df_results, fwd_results = bootstrap_irs(tenors, swap_rates)

# 시각화를 위한 데이터 정리
plot_times = np.sort(list(df_results.keys()))
plot_dfs = [df_results[t] for t in plot_times]

# Forward Rate 계단형 데이터 생성
fwd_times = [0] + list(tenors)
fwd_plot_rates = [fwd_results[0]] + fwd_results

# --- 결과 출력 ---
print("Time (Year) | Discount Factor | Inst. Forward Rate")
print("-" * 50)
fwd_idx = 0
for t in plot_times:
    f_val = fwd_results[fwd_idx] if t > 0 and t <= tenors[fwd_idx] else 0  # 간략화된 출력
    # 실제 구간 확인
    for i, stop_t in enumerate(tenors):
        if t <= stop_t and t > (tenors[i - 1] if i > 0 else 0):
            f_val = fwd_results[i]
            break
    print(f"{t:10} | {df_results[t]:15.6f} | {f_val:18.2%}")

# --- 차트 시각화 ---
plt.figure(figsize=(12, 5))

# 1. Discount Factor Chart
plt.subplot(1, 2, 1)
plt.plot(plot_times, plot_dfs, 'bo-', label='Discount Factor')
plt.title('Discount Factor Curve')
plt.xlabel('Tenor (Years)')
plt.ylabel('DF')
plt.grid(True, alpha=0.3)
plt.legend()

# 2. Instantaneous Forward Rate Chart
plt.subplot(1, 2, 2)
# 계단식 차트를 위해 step 함수 사용
plt.step(fwd_times[1:], fwd_results, where='pre', color='red', marker='s', label='Inst. Forward')
plt.title('Piecewise Constant Forward Rates')
plt.xlabel('Tenor (Years)')
plt.ylabel('Rate')
plt.grid(True, alpha=0.3)
plt.legend()

plt.tight_layout()
plt.show()