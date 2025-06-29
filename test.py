import numpy_financial as npf

cashflows = [-1000, 500, 400, 300]
irr = npf.irr(cashflows)

print(f"IRR = {irr * 100:.2f}%")

from scipy.optimize import brentq

cashflows = [-1000, 500, 400, 300]

def npv(rate):
    return sum(cf / (1 + rate) ** i for i, cf in enumerate(cashflows))

irr = brentq(npv, -0.99, 10)  # 保证在合理范围内搜索
print(f"IRR = {irr * 100:.2f}%")

import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import brentq

# 定义现金流
cashflows = [-1000, 500, 400, 300]

# 定义 NPV 函数
def npv(rate):
    return sum(cf / (1 + rate) ** i for i, cf in enumerate(cashflows))

# 求解IRR
irr = brentq(npv, -0.9, 2)  # 根据之前分析，搜索区间在 -90% 到 200%

print(f"IRR = {irr * 100:.2f}%")

# 构建收益率区间
rates = np.linspace(-0.9, 2, 500)

# 计算NPV
npvs = [npv(r) for r in rates]

# 画图
plt.figure(figsize=(8, 5))
plt.plot(rates, npvs, label="NPV curve")
plt.axhline(0, color="red", linestyle="--", label="NPV=0")
plt.axvline(irr, color="green", linestyle="--", label=f"IRR={irr * 100:.2f}%")
plt.scatter(irr, 0, color="green", zorder=5)

plt.title("NPV vs IRR")
plt.xlabel("Rate (r)")
plt.ylabel("NPV")
plt.grid(True)
plt.legend()
plt.show()
