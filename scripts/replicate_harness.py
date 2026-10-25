#!/usr/bin/env python3
"""独立复现与分析的通用骨架。把目标模型的计算逻辑写进 run()，其余工具直接复用。

使用步骤：
  1. 复制本文件为 engine_<ticker>.py，按 Excel 公式实现 run(params, glob) -> dict（至少含 'ps' 每股价值）。
  2. 在 SCENARIOS 里填入各情景参数，在 EXCEL_CHECK 里填 Excel 缓存值（dump_xlsx.py 可导出）。
  3. python3 engine_<ticker>.py check      # 逐项对账，误差应 < 1e-6
     python3 engine_<ticker>.py solve nrev 55.6   # 反推：基础情景下使每股价值=55.6 的 nrev
     python3 engine_<ticker>.py elastic    # 弹性表
     python3 engine_<ticker>.py curve nrev 10 18 1  # 输出一条曲线（可再按不同增长规模各跑一次）

示例实现见 examples/crwv/engine_crwv.py。
"""
import sys

# ---- 1. 全局输入（不随情景变化）----
GLOB = {}

# ---- 2. 情景参数 ----
SCENARIOS = {
    # 'base': dict(...),
}

# ---- 3. Excel 对账值：{(情景, 输出键): 缓存值} ----
EXCEL_CHECK = {}

# ---- 4. 弹性表：(参数, 变动量, 说明) ----
ELASTIC = []


def run(p, g=GLOB):
    """实现模型计算，返回 dict，至少包含 'ps'（每股价值）。"""
    raise NotImplementedError


# ================= 通用工具（一般无需修改）=================
def solve(key, target, lo, hi, base, out="ps", n=100):
    """二分法：求 base 中参数 key 取何值时 run()[out] == target。"""
    f = lambda x: run(dict(base, **{key: x}))[out] - target
    flo = f(lo)
    if flo * f(hi) > 0:
        return None  # 区间内无解
    for _ in range(n):
        mid = (lo + hi) / 2
        fm = f(mid)
        if flo * fm <= 0:
            hi = mid
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2


def check():
    worst = 0.0
    for (sc, key), xl in EXCEL_CHECK.items():
        v = run(SCENARIOS[sc])[key]
        d = abs(v - xl)
        worst = max(worst, d)
        flag = "" if d < 1e-6 else "  ⚠"
        print(f"{sc:>10} {key:>8}: py={v:.6f} xl={xl:.6f} diff={d:.2e}{flag}")
    print(f"最大误差 {worst:.2e}" + ("  —— 复现通过" if worst < 1e-6 else "  —— 未通过，请继续核对公式"))


def elastic(base_key="base"):
    b = SCENARIOS[base_key]
    b0 = run(b)["ps"]
    print(f"基础每股价值 {b0:.2f}")
    for k, d, label in ELASTIC:
        print(f"{label:<24} {run(dict(b, **{k: b[k] + d}))['ps'] - b0:+.2f}")


def curve(key, lo, hi, step, base_key="base"):
    b = SCENARIOS[base_key]
    x = lo
    while x <= hi + 1e-9:
        print(f"{x:g}\t{run(dict(b, **{key: x}))['ps']:.2f}")
        x += step


def cli():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "check"
    if cmd == "check":
        check()
    elif cmd == "solve":
        key, target = sys.argv[2], float(sys.argv[3])
        lo = float(sys.argv[4]) if len(sys.argv) > 4 else -1e3
        hi = float(sys.argv[5]) if len(sys.argv) > 5 else 1e3
        print(solve(key, target, lo, hi, SCENARIOS["base"]))
    elif cmd == "elastic":
        elastic()
    elif cmd == "curve":
        curve(sys.argv[2], float(sys.argv[3]), float(sys.argv[4]), float(sys.argv[5]))
    else:
        print(__doc__)


if __name__ == "__main__":
    cli()
