"""
KRW IRS Bootstrapping Program by Grok

Original Prompt:
1년 2년 3년 5년 krw irs rates :

2% 3% 4% 5%

일때에



piecewise constant instantaneous forward 가정으로 bootstrapping 해서 주요 노드별 discount factor 와 구간별 instantaneous fowrad 값을 출력해주는 프로그램을 작성해주고



각각 챠트를 그려줘



file name 에 "by_Grok" 포함해주고



이 프롬프트를 문자 그대로 문서화 해서 적어줘
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def bootstrap_krw_irs_rates():
    """
    Bootstrap discount factors and instantaneous forward rates from KRW IRS rates
    assuming piecewise constant instantaneous forward rates.

    IRS rates are assumed to be continuously compounded.
    """

    # Input data: maturity (years) and IRS rates (%)
    maturities = np.array([1, 2, 3, 5])
    irs_rates = np.array([0.02, 0.03, 0.04, 0.05])  # 2%, 3%, 4%, 5%

    # Calculate discount factors from IRS rates (continuous compounding)
    discount_factors = np.exp(-irs_rates * maturities)

    # Calculate instantaneous forward rates using bootstrapping
    # Forward rate r_i for interval [t_{i-1}, t_i] satisfies:
    # DF(t_i) = DF(t_{i-1}) * exp(-r_i * (t_i - t_{i-1}))

    times = np.concatenate([[0], maturities])
    dfs = np.concatenate([[1.0], discount_factors])

    forward_rates = []
    interval_starts = []
    interval_ends = []

    for i in range(1, len(times)):
        t_prev = times[i-1]
        t_curr = times[i]
        df_prev = dfs[i-1]
        df_curr = dfs[i]

        # Calculate forward rate: r_i = -ln(DF(t_i)/DF(t_{i-1})) / (t_i - t_{i-1})
        r_i = -np.log(df_curr / df_prev) / (t_curr - t_prev)
        forward_rates.append(r_i)
        interval_starts.append(t_prev)
        interval_ends.append(t_curr)

    return times, dfs, interval_starts, interval_ends, forward_rates

def plot_results(times, dfs, interval_starts, interval_ends, forward_rates):
    """Plot discount factors and forward rates"""

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Plot discount factors
    ax1.plot(times, dfs, 'bo-', linewidth=2, markersize=8)
    ax1.set_xlabel('Time (years)')
    ax1.set_ylabel('Discount Factor')
    ax1.set_title('Discount Factors')
    ax1.grid(True, alpha=0.3)

    # Plot forward rates
    for i, (start, end, rate) in enumerate(zip(interval_starts, interval_ends, forward_rates)):
        ax2.hlines(rate, start, end, colors='red', linewidth=3, label=f'Interval {i+1}' if i == 0 else "")
        # Add markers at interval boundaries
        ax2.plot([start, end], [rate, rate], 'ro', markersize=6)

    ax2.set_xlabel('Time (years)')
    ax2.set_ylabel('Instantaneous Forward Rate')
    ax2.set_title('Piecewise Constant Instantaneous Forward Rates')
    ax2.grid(True, alpha=0.3)
    ax2.legend()

    plt.tight_layout()
    plt.show()

def main():
    """Main function to run the bootstrapping analysis"""

    print("KRW IRS Bootstrapping Analysis")
    print("=" * 40)

    # Perform bootstrapping
    times, dfs, interval_starts, interval_ends, forward_rates = bootstrap_krw_irs_rates()

    # Display results
    print("\nDiscount Factors at Major Nodes:")
    print("-" * 30)
    for t, df in zip(times, dfs):
        print(".4f")

    print("\nInstantaneous Forward Rates by Interval:")
    print("-" * 40)
    for i, (start, end, rate) in enumerate(zip(interval_starts, interval_ends, forward_rates)):
        print(".1f")

    # Plot results
    print("\nGenerating plots...")
    plot_results(times, dfs, interval_starts, interval_ends, forward_rates)

    print("\nAnalysis complete!")

if __name__ == "__main__":
    main()
