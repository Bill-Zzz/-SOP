#!/usr/bin/env python3
"""CoreWeave（CRWV）分层估值模型的 Python 复现（对应 CRWV_分层估值模型_修订版.xlsx）。

按季度计算 30 年（120 个季度），四层现金流：
  L1a 存量产能合同 EBITDA（折现率 d1）
  L1b 积压订单合同 EBITDA − 为其新建产能的资本开支（d1）
  L2  合同到期后按折价续租的残值 EBITDA（d2）
  L3  未来新签产能的合同 EBITDA + 残值 − 资本开支（d3）
  税在公司层面计算：资本开支当期全额抵税，亏损结转（d3 折现）
股权价值 = EV − 净债务 − 已收客户预付款；单位：十亿美元，每股为美元。

用法：python3 engine_crwv.py check | solve nrev 55.6 10 30 | elastic | curve nrev 10 18 1
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "scripts"))
import replicate_harness as H  # noqa: E402

GLOB = dict(
    tax=0.21,        # 联邦税率
    fixed=0.5,       # 续租价低于运营成本时仍需承担的运营成本比例（本文假设）
    q=0.25,          # 季度长度
    shares=0.608,    # 十亿股（高盛 10/21 市值反推；伯恩斯坦口径约 0.559，待核实）
    netdebt=35 - 5.5,  # 债务 350 亿（伯恩斯坦 10/15）− 非受限现金 55 亿（高盛 10/21 附表 1）
    prepay=7.5,      # 递延收入 75.23 亿（汇丰 8/6）
    nol0=5,          # 期初税务亏损（假设）
    backlog=129,     # 在手订单（高盛 10/21 附表 1）
    cov=50,          # 存量产能覆盖订单（高盛 10/21 情景一；情景二为 75）
    gw0=1.5, rev0=11, age0=1.5,  # 存量产能、每 GW 年收入、GPU 加权年龄（高盛 10/21）
    blife=4.5,       # 积压订单合同年限
    nlife=5,         # 新合同年限（高盛 10/19 附表 3）
    blag=0.75,       # 积压订单平均上线时点（年）
)
SCENARIOS = {
    "opt": dict(m=.71, brev=11.5, bmul=1.75, nrev=13.5, ncap=31, life=7, ren=.8,
                d1=.085, d2=.12, d3=.11, vol=[.5, .6] + [1.5] * 13),
    "base": dict(m=.67, brev=11, bmul=1.6, nrev=12, ncap=32, life=6, ren=.5,
                 d1=.095, d2=.14, d3=.12, vol=[.5, .6] + [.75] * 13),
    "bear": dict(m=.62, brev=10.5, bmul=1.35, nrev=10.5, ncap=35, life=5, ren=.25,
                 d1=.11, d2=.16, d3=.135, vol=[0, .5, .5] + [0] * 12),
    "gs": dict(m=.69, brev=11, bmul=1.7, nrev=13, ncap=30, life=6, ren=1.0,
               d1=.095, d2=.118, d3=.118, vol=[.5, .6] + [.75] * 13),
}
# 修订版 Excel（2026-10-25）估值汇总页第 14 行
EXCEL_CHECK = {
    ("opt", "ps"): 155.110149971726, ("base", "ps"): -40.6357146642081,
    ("bear", "ps"): -69.398194492577, ("gs", "ps"): 48.4266189646346,
}
ELASTIC = [
    ("nrev", 1, "新合同每 GW 年收入 +10 亿"),
    ("life", 1, "GPU 寿命 +1 年"),
    ("ren", .1, "续租价格 +10 个百分点"),
    ("ncap", -1, "每 GW 资本开支 −10 亿"),
    ("m", .01, "EBITDA 利润率 +1 个百分点"),
    ("bmul", .1, "积压订单收入/资本开支 +0.1 倍"),
]


def run(p, g=GLOB, detail=False):
    q = g["q"]
    A = np.arange(1, 121)
    C = (A - .5) * q          # 季中时点（年）
    D = (A - 1) * q           # 季初时点
    m, life, ren = p["m"], p["life"], p["ren"]
    rem0 = g["cov"] / (g["gw0"] * g["rev0"])         # 存量剩余合同年数
    life0 = max(life - g["age0"], rem0)              # 存量剩余寿命
    bl = g["backlog"] - g["cov"]
    bgw = bl / (p["brev"] * g["blife"])              # 积压所需 GW
    bcap = bl / p["bmul"]                            # 积压所需资本开支

    def resid(rev):  # 每 GW 每季度残值 EBITDA
        return max(rev * ren * q - rev * (1 - m) * q, -g["fixed"] * rev * (1 - m) * q)

    E = np.where(C < rem0, g["gw0"] * g["rev0"] * m * q, 0)
    F = np.where((C >= rem0) & (C < life0), g["gw0"] * resid(g["rev0"]), 0)
    x = C - g["blag"]
    G = np.where((x > 0) & (x < g["blife"]), bgw * p["brev"] * m * q, 0)
    Hh = np.where((x >= g["blife"]) & (x < life), bgw * resid(p["brev"]), 0)
    I = np.where((np.abs(D) < 1e-6) | (np.abs(D - .5) < 1e-6), bcap / 2, 0)
    J = np.zeros(120); K = np.zeros(120); L = np.zeros(120)
    for w, v in zip(range(1, 16), p["vol"]):
        J += p["nrev"] * m * q * v * (((C - w) > 0) & ((C - w) < g["nlife"]))
        K += resid(p["nrev"]) * v * (((C - w) >= g["nlife"]) & ((C - w) < life))
        L += p["ncap"] * v * (np.abs(D - (w - .5)) < 1e-6)
    M = E + F + G + Hh + J + K
    N = I + L
    O = M - N
    R = np.zeros(120)
    nol = g["nol0"]
    for i in range(120):
        if O[i] < 0:
            nol -= O[i]
        else:
            use = min(nol, O[i]); R[i] = g["tax"] * (O[i] - use); nol -= use
    S = (1 + p["d1"]) ** -C; T = (1 + p["d2"]) ** -C; U = (1 + p["d3"]) ** -C
    out = dict(L1a=E @ S, L1b=G @ S - I @ S, L2=F @ T + Hh @ T,
               L3=J @ U + K @ U - L @ U, tax=-(R @ U))
    out["EV"] = sum(out.values())
    out["eq"] = out["EV"] - g["netdebt"] - g["prepay"]
    out["ps"] = out["eq"] / g["shares"]
    if detail:
        out["fcf"] = M - N - R
    return out


H.GLOB, H.SCENARIOS, H.EXCEL_CHECK, H.ELASTIC, H.run = GLOB, SCENARIOS, EXCEL_CHECK, ELASTIC, run

if __name__ == "__main__":
    H.cli()
