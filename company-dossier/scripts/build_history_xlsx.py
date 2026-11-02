#!/usr/bin/env python3
"""从富途导出生成一家公司的"历史数据 Excel"（层级 1），规格见 ../templates/financial_history_spec.md。

用法：
  python3 build_history_xlsx.py <代码> --root <futu_export 根目录> --out <输出目录> [--ai 名字] [--check]
  python3 build_history_xlsx.py <代码> --root <根目录> --check-only <已生成的 xlsx>
  加 --png <目录>：把"图表总览"的每张图用同一份数据区（"图表计算"）画成 PNG 快照，文件名 图01_营收与增速.png、
  图02_分部结构.png……编号与图表总览顺序一致；第二层没填、数据为空的图跳过并打印"跳过：图 X 无数据"。
  依赖 matplotlib（缺时跳过并提示，不报错）；中文字体用 Noto Sans CJK / Noto Sans SC / WenQuanYi，找不到退回 DejaVu 并警告。
  同时加 --check 时多一项"9 PNG 快照数量 = 有数据的图数量"。每家公司用单独的 PNG 目录（文件名不含代码）。

输入（<根目录>/<代码>/ 下）：
  <代码>_富途三表_长表.csv（优先；没有则用 <代码>_富途财报_长表.csv）
  <代码>_日线_库内拆股调整.csv（优先）+ <代码>_日线_不复权.csv（补到最新）
  底稿CSV/<代码>_科目映射.csv、_缺口.csv、_检查.csv（有则用）
可选：--ref-mcap <CSV>（列：年度,市值，单位亿美元），作为"数据提供方市值"参与验收的 5% 对照。

只依赖 openpyxl 与标准库（--png 另需 matplotlib）。工作簿用 openpyxl 新建，图表为原生图表，数据区在"图表计算"表用公式引用。
"""
import argparse
import csv
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import OrderedDict, defaultdict
from decimal import Decimal, InvalidOperation
from xml.etree import ElementTree as ET

try:
    import openpyxl
    from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:  # pragma: no cover
    sys.exit("缺少 openpyxl：python3 -m pip install --user openpyxl")

SCRIPT_VERSION = "0.8.0"
CUR_NAME = {"USD": "美元", "TWD": "新台币", "EUR": "欧元", "HKD": "港元", "CNY": "人民币",
            "JPY": "日元", "GBP": "英镑", "KRW": "韩元"}
STMTS = ["利润表", "资产负债表", "现金流量表"]
# 个别公司的已知口径问题（来自任务 05 回执）
SPECIAL = {
    "DELL": {"price_start": "2018-12-28",
             "note": "2018-12-28 之前的日线是 DVMT 跟踪股，不是现在的 DELL C 类股，已舍弃。"},
    "SNDK": {"from_year": 2024,
             "note": "富途把旧 SanDisk（1994–2016）与新 Sandisk（2024 起）接在同一代码下，只用 FY2024 起的数据。"},
}

# 核心科目：键、中文名、报表、候选科目（按顺序逐期取第一个有数的）、单位类型
CORE = [
    ("rev", "营收", "利润表", ["营业总收入", "总收入", "营业收入"], "amt"),
    ("cogs", "营业成本", "利润表", ["营业总成本", "营业成本"], "amt"),
    ("gp", "毛利", "利润表", ["毛利"], "amt"),
    ("rd", "研发费用", "利润表", ["研发费用"], "amt"),
    ("sga", "销售及管理费用", "利润表", ["销售和管理费用"], "amt"),
    ("op", "营业利润", "利润表", ["营业利润"], "amt"),
    ("intexp", "利息费用", "利润表", ["营业外利息费用"], "amt"),
    ("pti", "税前利润", "利润表", ["税前利润"], "amt"),
    ("tax", "所得税", "利润表", ["所得税"], "amt"),
    ("ni", "净利润（归母）", "利润表", ["归属于母公司股东净利润", "净利润"], "amt"),
    ("nic", "归属普通股股东净利润", "利润表", ["归属于普通股股东净利润", "归属于母公司股东净利润", "净利润"], "amt"),
    ("eps_b", "基本每股收益", "利润表", ["基本每股收益"], "eps"),
    ("eps_d", "稀释每股收益", "利润表", ["稀释每股收益"], "eps"),
    ("cash", "现金及等价物", "资产负债表", ["-现金和现金等价物", "现金和现金等价物"], "amt"),
    ("cash_sti", "现金及短期投资", "资产负债表", ["现金及现金等价物和短期投资", "-现金和现金等价物"], "amt"),
    ("ar", "应收账款", "资产负债表", ["-应收账款净额", "应收款项"], "amt"),
    ("inv", "存货", "资产负债表", ["存货"], "amt"),
    ("ppe", "固定资产净额", "资产负债表", ["固定资产净额"], "amt"),
    ("gw", "商誉", "资产负债表", ["-商誉", "商誉"], "amt"),
    ("ta", "总资产", "资产负债表", ["资产合计"], "amt"),
    ("ap", "应付账款", "资产负债表", ["-应付账款", "应付款项"], "amt"),
    ("std", "短期债务", "资产负债表", ["-短期借款", "短期借款"], "amt"),
    ("ltd", "长期债务", "资产负债表", ["-长期借款", "长期借款"], "amt"),
    ("tl", "总负债", "资产负债表", ["负债合计"], "amt"),
    ("eq", "归母股东权益", "资产负债表", ["归属于母公司股东权益合计", "股东权益合计"], "amt"),
    ("eq_t", "股东权益合计（含少数股东）", "资产负债表", ["股东权益合计", "归属于母公司股东权益合计"], "amt"),
    ("ocf", "经营现金流", "现金流量表", ["经营活动现金流量净额"], "amt"),
    ("capex", "资本开支", "现金流量表", ["固定资产交易净额", "资本性支出", "购买固定资产"], "amt"),
    ("da", "折旧与摊销", "现金流量表", ["折旧摊销及损耗", "折旧与摊销"], "amt"),
    ("acq", "收购（业务交易净额）", "现金流量表", ["业务交易净额"], "amt"),
    ("div", "分红", "现金流量表", ["现金股利支付"], "amt"),
    ("buyback", "普通股发行/回购净额", "现金流量表", ["普通股发行/回购的净额"], "amt"),
    ("dtax", "递延所得税（现金流量表调整项）", "现金流量表", ["递延所得税"], "amt"),
    ("fcf_futu", "自由现金流（富途）", "现金流量表", ["自由现金流"], "amt"),
]
CORE_BY_KEY = {c[0]: c for c in CORE}
# 期末股数候选科目（单位：股；在资产负债表、关键指标里按顺序找）。富途美股模板的"股本 / -普通股股本"是
# 面值金额（发行股数 × 每股面值），不是股数，不在候选里；候选值还须与当年基本加权股数相差 0.5–2 倍才采用。
END_SHARE_STMTS = ["资产负债表", "关键指标"]
END_SHARE_ITEMS = ["期末普通股股数", "普通股股数", "期末股数", "期末流通股数", "流通股数", "已发行普通股股数",
                   "期末总股本（股）", "总股本（股）", "流通股本（股）"]
# 底稿"科目映射.csv"里的中文名 → 本脚本核心键
MAP_NAME = {"营收": "rev", "营业成本": "cogs", "毛利": "gp", "研发费用": "rd", "销售及管理费用": "sga",
            "营业利润": "op", "净利润": "ni", "基本每股收益": "eps_b", "稀释每股收益": "eps_d",
            "经营现金流": "ocf", "资本开支": "capex", "折旧与摊销": "da", "分红": "div",
            "现金及等价物": "cash", "应收账款": "ar", "存货": "inv", "商誉": "gw", "总资产": "ta",
            "总负债": "tl", "股东权益": "eq", "短期债务": "std", "长期债务": "ltd"}

HDR_FILL = PatternFill("solid", fgColor="DCE6F1")
SEC_FILL = PatternFill("solid", fgColor="F2F2F2")
L2_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
NOTE = Font(italic=True, color="666666", size=9)
WRAP = Alignment(wrap_text=True, vertical="top")
F_AMT, F_PCT, F_X, F_SH, F_DATE, F_PX = "#,##0.00", "0.0%", "0.00\"x\"", "#,##0.0", "yyyy-mm-dd", "#,##0.00"


# ---------------------------------------------------------------- 读数据
def fnum(s):
    if s is None:
        return None
    s = str(s).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def decimals_of(s):
    try:
        d = Decimal(str(s).strip()).normalize()
    except InvalidOperation:
        return None
    e = d.as_tuple().exponent
    return max(0, -e) if isinstance(e, int) else None


class Export:
    """富途长表 → {(报表, 期间): {科目: 值}} 以及期间元数据。"""

    def __init__(self, code, root, from_year=None):
        self.code = code
        self.dir = os.path.join(root, code)
        cands = [f"{code}_富途三表_长表.csv", f"{code}_富途财报_长表.csv"]
        self.long_file = next((c for c in cands if os.path.exists(os.path.join(self.dir, c))), None)
        if not self.long_file:
            sys.exit(f"找不到长表：{self.dir}/{cands[0]}（或 {cands[1]}）")
        path = os.path.join(self.dir, self.long_file)
        self.long_mtime = dt.date.fromtimestamp(os.path.getmtime(path)).isoformat()
        self.data = defaultdict(OrderedDict)
        self.raw = defaultdict(dict)
        self.meta = {}
        self.order = defaultdict(list)
        self.currencies = set()
        self.dup = 0
        self.date_fixed_col = False
        with open(path, encoding="utf-8-sig") as f:
            rd = csv.DictReader(f)
            self.date_fixed_col = "日期已校正" in (rd.fieldnames or [])
            for r in rd:
                st, per, item = r["报表"], r["期间"], r["科目"]
                try:
                    fy = int(float(r["财年"]))
                except (TypeError, ValueError):
                    continue
                if from_year and fy < from_year:
                    continue
                typ = str(r["报告类型"]).strip()
                if per not in self.meta:
                    self.meta[per] = {"end": r["期末日"][:10], "fy": fy, "type": typ}
                if r.get("币种"):
                    self.currencies.add(r["币种"].strip())
                if item not in self.order[st]:
                    self.order[st].append(item)
                if item in self.data[(st, per)]:
                    self.dup += 1
                    continue
                self.data[(st, per)][item] = fnum(r["数值"])
                self.raw[(st, per)][item] = r["数值"]
        self.currency = sorted(self.currencies)[0] if len(self.currencies) == 1 else (
            "USD" if "USD" in self.currencies else (sorted(self.currencies) or ["USD"])[0])
        self.annual = sorted((m["fy"], p) for p, m in self.meta.items() if m["type"] == "7")
        self.years = [y for y, _ in self.annual]
        self.aper = {y: p for y, p in self.annual}
        qs = [(m["end"], p) for p, m in self.meta.items() if m["type"] in ("1", "2", "3", "4")]
        self.quarters = [p for _, p in sorted(qs)]
        self.mapping = {}
        mp = os.path.join(self.dir, "底稿CSV", f"{code}_科目映射.csv")
        if os.path.exists(mp):
            with open(mp, encoding="utf-8-sig") as f:
                for r in csv.DictReader(f):
                    k = MAP_NAME.get(r.get("中文名称", "").strip())
                    nm = (r.get("公司原文名称") or "").strip()
                    if k and nm and "÷" not in nm:
                        self.mapping[k] = nm
        self.gaps = self._read_csv(os.path.join(self.dir, "底稿CSV", f"{code}_缺口.csv"))
        self.checks = self._read_csv(os.path.join(self.dir, "底稿CSV", f"{code}_检查.csv"))

    @staticmethod
    def _read_csv(p):
        if not os.path.exists(p):
            return []
        with open(p, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))

    def label(self, per):
        m = self.meta[per]
        return f"FY{m['fy']}" if m["type"] == "7" else f"FY{m['fy']}Q{m['type']}"

    def cands(self, key):
        c = list(CORE_BY_KEY[key][3])
        if key in self.mapping and self.mapping[key] not in c:
            c.insert(0, self.mapping[key])
        return c

    def core(self, key, per):
        """返回 (原币原值, 用到的科目, 原始字符串)。"""
        st = CORE_BY_KEY[key][2]
        d = self.data.get((st, per), {})
        for nm in self.cands(key):
            v = d.get(nm)
            if v is not None:
                if key == "capex" and nm == "资本性支出" and v > 0:
                    v = -v
                return v, nm, self.raw[(st, per)].get(nm)
        return None, None, None

    def shares(self, per, kind="b"):
        """加权股数（百万股）= 归属普通股净利润 ÷ 每股收益；返回 (值, 相对舍入误差)。"""
        nic = self.core("nic", per)[0]
        eps, _, eraw = self.core("eps_" + kind, per)
        if nic is None or eps in (None, 0) or nic == 0:
            return None, None
        d = decimals_of(eraw)
        # 导出会丢掉末尾的 0（3.30 → 3.3），按至少 2 位小数估计
        err = (0.5 * 10 ** (-max(d, 2))) / abs(eps) if d is not None else None
        return nic / eps / 1e6, err

    def end_shares(self, per):
        """期末股数（百万股）；返回 (值, 科目名, 未采用的原因)。量级与基本加权股数不符（0.5–2 倍以外）的不采用。"""
        for st in END_SHARE_STMTS:
            d = self.data.get((st, per), {})
            for nm in END_SHARE_ITEMS:
                v = d.get(nm)
                if v is None or v <= 0:
                    continue
                v /= 1e6
                wb = self.shares(per, "b")[0]
                if not wb or wb <= 0:
                    return None, nm, f"{st}·{nm} 有数，但当年没有加权股数可比对量级，未采用"
                if not 0.5 <= v / wb <= 2:
                    return None, nm, f"{st}·{nm} ÷ 基本加权股数 = {v / wb:.3g}，不在 0.5–2，疑似单位不同，未采用"
                return v, f"{st}·{nm}", None
        return None, None, None

    def par_amount(self, per):
        """资产负债表'股本'（面值金额，原币）——只用于在说明里解释它不是股数。"""
        d = self.data.get(("资产负债表", per), {})
        for nm in ("股本", "-普通股股本"):
            if d.get(nm) is not None:
                return d[nm], nm
        return None, None


def load_daily(code, root, price_start=None):
    """拆股调整日线：库内文件为主，用不复权文件补到最新（按重叠段比值换算）。"""
    d = os.path.join(root, code)
    adj = os.path.join(d, f"{code}_日线_库内拆股调整.csv")
    raw = os.path.join(d, f"{code}_日线_不复权.csv")

    def rd(p):
        out = {}
        if not os.path.exists(p):
            return out
        with open(p, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                c = fnum(r.get("close"))
                if c and c > 0:
                    out[r["time_key"][:10]] = c
        return out

    a, b = rd(adj), rd(raw)
    info = {"adj_file": os.path.basename(adj) if a else None, "raw_file": os.path.basename(raw) if b else None}
    if a:
        series = dict(a)
        last = max(a)
        common = sorted(set(a) & set(b))[-20:]
        if common and any(k > last for k in b):
            ratios = sorted(a[k] / b[k] for k in common)
            fac = ratios[len(ratios) // 2]
            for k, v in b.items():
                if k > last:
                    series[k] = v * fac
            info["ext"] = f"{last} 之后用不复权日线 × {fac:.6f}（近 {len(common)} 个重叠日中位比值）补到 {max(b)}"
        info["basis"] = "拆股调整价（库内文件）"
    else:
        series = dict(b)
        info["basis"] = "未复权价（无库内拆股调整文件；有拆股的年份股价与股数口径不一致）"
    if price_start:
        series = {k: v for k, v in series.items() if k >= price_start}
    days = sorted(series)
    info["start"], info["end"] = (days[0], days[-1]) if days else (None, None)
    return [(dt.date.fromisoformat(k), series[k]) for k in days], info


def price_on_or_before(px, date, max_gap=10):
    lo, hi = 0, len(px) - 1
    best = None
    while lo <= hi:
        mid = (lo + hi) // 2
        if px[mid][0] <= date:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    if best is None or (date - px[best][0]).days > max_gap:
        return None
    return px[best]


def find_events(px, one_day=0.08, win=21, win_th=0.20):
    ev = []
    for i in range(1, len(px)):
        r = px[i][1] / px[i - 1][1] - 1
        if abs(r) >= one_day:
            ev.append({"date": px[i][0], "type": "单日", "move": r, "close": px[i][1], "start": px[i - 1][0]})
    run = []

    def flush():
        if run:
            best = max(run, key=lambda x: abs(x[1]))
            j = best[0]
            ev.append({"date": px[j][0], "type": f"{win}日", "move": best[1], "close": px[j][1],
                       "start": px[j - win][0], "span": (px[run[0][0]][0], px[run[-1][0]][0])})

    for i in range(win, len(px)):
        r = px[i][1] / px[i - win][1] - 1
        if abs(r) >= win_th:
            if run and (i - run[-1][0] > 5 or (r > 0) != (run[-1][1] > 0)):
                flush()
                run = []
            run.append((i, r))
    flush()
    ev.sort(key=lambda e: (e["date"], e["type"]))
    return ev


# ---------------------------------------------------------------- 写工作簿
def hdr(ws, row, values, col=1, fill=HDR_FILL):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = BOLD
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = NOTE


def guard(expr, refs):
    """所有输入都有数才计算，否则留空；计算出错也留空（空白 ≠ 0）。"""
    return f'=IF(COUNT({",".join(refs)})<{len(refs)},"",IFERROR({expr},""))'


def q(sheet):
    return f"'{sheet}'!"


class Builder:
    def __init__(self, ex, px, pxinfo, ai, ref_mcap=None, special=None):
        self.ex, self.px, self.pxinfo, self.ai = ex, px, pxinfo, ai
        self.ref_mcap = ref_mcap or {}
        self.special = special or {}
        self.cur = ex.currency
        self.usd = self.cur == "USD"
        self.U = f"亿{CUR_NAME.get(self.cur, self.cur)}"
        self.today = dt.date.today()
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        names = ["图表总览", "说明", "历年营收与业务拆分", "年度三表", "年度衍生指标", "近12季度明细",
                 "市值与营收对比", "市值与估值", "股价事件", "IPO指标", "所得税与现金税", "重大融资时间线",
                 "重要收购", "指引与实际", "缺口", "来源", "图表计算"]
        self.ws = {n: self.wb.create_sheet(n) for n in names}
        self.years = ex.years
        self.gaps = []
        self.expected_charts = 0
        self.share_basis = {}

    # ---- 年度三表
    def sheet_statements(self):
        ws, ex = self.ws["年度三表"], self.ex
        title(ws, f"{ex.code} 年度三表（{self.years[0]}–{self.years[-1]}）",
              f"金额：{self.U}（富途原值 ÷ 1e8）；每股收益：{CUR_NAME.get(self.cur, self.cur)}/股；"
              "富途关键指标为原值（%、次、天）；空白 = 富途无数，不等于 0；富途是最新口径，不是首次披露。")
        self.ycol = {y: get_column_letter(5 + i) for i, y in enumerate(self.years)}
        self.src_col = get_column_letter(5 + len(self.years))
        hdr(ws, 4, ["区块", "原始科目名", "标准化中文名", "单位"] + self.years + ["来源ID"])
        ws.freeze_panes = "E5"
        r = 5
        self.trow = {}
        ws.cell(row=r, column=1, value="核心科目（标准化，按候选科目逐年取第一个有数的）").font = BOLD
        ws.cell(row=r, column=1).fill = SEC_FILL
        r += 1
        for key, name, st, _, kind in CORE:
            used = OrderedDict()
            for y in self.years:
                v, nm, _ = ex.core(key, ex.aper[y])
                if v is None:
                    continue
                used.setdefault(nm, []).append(y)
                c = ws[f"{self.ycol[y]}{r}"]
                c.value = v / 1e8 if kind == "amt" else v
                c.number_format = F_AMT if kind == "amt" else "0.0000"
            src = "；".join(f"{st}·{nm}（{ys[0]}–{ys[-1]}）" if len(used) > 1 else f"{st}·{nm}"
                           for nm, ys in used.items()) or "富途无此科目"
            ws.cell(row=r, column=1, value="核心")
            ws.cell(row=r, column=2, value=src)
            ws.cell(row=r, column=3, value=name)
            ws.cell(row=r, column=4, value=self.U if kind == "amt" else f"{CUR_NAME.get(self.cur, self.cur)}/股")
            ws[f"{self.src_col}{r}"] = "F1"
            self.trow[key] = r
            r += 1
        for key, name, kind in [("sh_b", "基本加权股数", "b"), ("sh_d", "稀释加权股数", "d")]:
            ws.cell(row=r, column=1, value="核心（推算）")
            ws.cell(row=r, column=2, value=f"归属普通股股东净利润 × 100 ÷ {'基本' if kind == 'b' else '稀释'}每股收益")
            ws.cell(row=r, column=3, value=name)
            ws.cell(row=r, column=4, value="百万股")
            for y in self.years:
                n = f"{self.ycol[y]}{self.trow['nic']}"
                e = f"{self.ycol[y]}{self.trow['eps_' + kind]}"
                ws[f"{self.ycol[y]}{r}"] = guard(f'IF(OR({e}=0,{n}=0),"",{n}*100/{e})', [n, e])
                ws[f"{self.ycol[y]}{r}"].number_format = F_SH
            ws[f"{self.src_col}{r}"] = "F1"
            self.trow[key] = r
            r += 1
        r += 1
        # 全部原始科目
        std_name = {}
        for key, name, st, _, _ in CORE:
            for nm in ex.cands(key):
                std_name.setdefault((st, nm), name)
        for st in STMTS + ["关键指标"]:
            items = ex.order.get(st, [])
            if not items:
                continue
            ws.cell(row=r, column=1, value=f"{st}（富途全部科目）" if st != "关键指标" else "富途关键指标（原值）").font = BOLD
            ws.cell(row=r, column=1).fill = SEC_FILL
            r += 1
            for it in items:
                ws.cell(row=r, column=1, value=st)
                ws.cell(row=r, column=2, value=it)
                ws.cell(row=r, column=3, value=std_name.get((st, it), ""))
                is_eps = "每股" in it
                unit = "原值" if st == "关键指标" else (f"{CUR_NAME.get(self.cur, self.cur)}/股" if is_eps else self.U)
                ws.cell(row=r, column=4, value=unit)
                for y in self.years:
                    v = ex.data.get((st, ex.aper[y]), {}).get(it)
                    if v is None:
                        continue
                    c = ws[f"{self.ycol[y]}{r}"]
                    c.value = v if (st == "关键指标" or is_eps) else v / 1e8
                    c.number_format = "0.0000" if is_eps else F_AMT
                ws[f"{self.src_col}{r}"] = "F1"
                r += 1
            r += 1
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 34
        ws.column_dimensions["C"].width = 22
        for y in self.years:
            ws.column_dimensions[self.ycol[y]].width = 11

    def T(self, key, y):
        return f"{q('年度三表')}{self.ycol[y]}{self.trow[key]}"

    # ---- 年度衍生指标
    def sheet_derived(self):
        ws, ex = self.ws["年度衍生指标"], self.ex
        title(ws, f"{ex.code} 年度衍生指标",
              f"公式写在单元格里，全部引用'年度三表'核心科目；金额 {self.U}；比率 %；空白 = 输入缺数。"
              "FCF = 经营现金流 + 资本开支（资本开支为负数，富途'固定资产交易净额'为购建减处置的净额）；"
              "净现金 = 现金及短期投资 − 短期债务 − 长期债务（债务缺数按 0）；"
              "ROIC = 营业利润 ×（1 − 税率）÷ 期末（股东权益合计 + 短期债务 + 长期债务），税率 = 所得税 ÷ 税前利润，限 0–35%，税前利润 ≤ 0 时取 21%；"
              "ROE = 净利润（归母）÷ 期末归母股东权益；天数按 365 天、期末余额。")
        cols = [
            ("年度", None), ("财年期末日", None), (f"营收（{self.U}）", F_AMT), ("营收同比", F_PCT),
            ("毛利率", F_PCT), ("营业利润率", F_PCT), ("净利率", F_PCT), ("研发 / 收入", F_PCT),
            ("资本开支 / 收入", F_PCT), ("折旧 / 收入", F_PCT), (f"经营现金流（{self.U}）", F_AMT),
            (f"资本开支（{self.U}）", F_AMT), (f"FCF（{self.U}）", F_AMT), ("FCF / 净利润", F_PCT),
            (f"净现金（{self.U}）", F_AMT), ("商誉 / 总资产", F_PCT), ("基本加权股数（百万股）", F_SH),
            ("股数年变化（倍）", "0.000"), ("稀释加权股数（百万股）", F_SH), ("每股营收", "0.000"),
            ("每股 FCF", "0.000"), ("有效税率", F_PCT), ("ROIC", F_PCT), ("ROE", F_PCT),
            (f"EBITDA（{self.U}）", F_AMT), ("OCF / 净利润", F_PCT), ("应收天数", "0.0"), ("存货天数", "0.0"),
            ("应付天数", "0.0"), ("GAAP − 非GAAP 差额（第二层）", F_AMT), ("备注", None), ("来源ID", None)]
        hdr(ws, 4, [c[0] for c in cols])
        self.dcol = {c[0]: get_column_letter(i + 1) for i, c in enumerate(cols)}
        ws.freeze_panes = "C5"
        self.drow = {}
        T = self.T
        for i, y in enumerate(self.years):
            r = 5 + i
            self.drow[y] = r
            p = self.years[i - 1] if i > 0 and self.years[i - 1] == y - 1 else None
            ws[f"A{r}"] = y
            ws[f"B{r}"] = dt.date.fromisoformat(ex.meta[ex.aper[y]]["end"])
            ws[f"B{r}"].number_format = F_DATE
            D = self.dcol
            cell = lambda name: f"{D[name]}{r}"
            f = {}
            f[f"营收（{self.U}）"] = f"={T('rev', y)}" if ex.core("rev", ex.aper[y])[0] is not None else None
            if p is not None:
                f["营收同比"] = guard(f"{T('rev', y)}/{T('rev', p)}-1", [T("rev", y), T("rev", p)])
                f["股数年变化（倍）"] = guard(f"{T('sh_b', y)}/{T('sh_b', p)}", [T("sh_b", y), T("sh_b", p)])
            for name, a, b in [("毛利率", "gp", "rev"), ("营业利润率", "op", "rev"), ("净利率", "ni", "rev"),
                               ("研发 / 收入", "rd", "rev"), ("折旧 / 收入", "da", "rev"),
                               ("商誉 / 总资产", "gw", "ta"), ("ROE", "ni", "eq")]:
                f[name] = guard(f"{T(a, y)}/{T(b, y)}", [T(a, y), T(b, y)])
            f["资本开支 / 收入"] = guard(f"-{T('capex', y)}/{T('rev', y)}", [T("capex", y), T("rev", y)])
            f[f"经营现金流（{self.U}）"] = guard(T("ocf", y), [T("ocf", y)])
            f[f"资本开支（{self.U}）"] = guard(T("capex", y), [T("capex", y)])
            f[f"FCF（{self.U}）"] = guard(f"{T('ocf', y)}+{T('capex', y)}", [T("ocf", y), T("capex", y)])
            fcf = cell(f"FCF（{self.U}）")
            f["FCF / 净利润"] = guard(f'IF({T("ni", y)}<=0,"",{fcf}/{T("ni", y)})', [fcf, T("ni", y)])
            f[f"净现金（{self.U}）"] = guard(f"{T('cash_sti', y)}-N({T('std', y)})-N({T('ltd', y)})", [T("cash_sti", y)])
            f["基本加权股数（百万股）"] = guard(T("sh_b", y), [T("sh_b", y)])
            f["稀释加权股数（百万股）"] = guard(T("sh_d", y), [T("sh_d", y)])
            sh = cell("稀释加权股数（百万股）")
            f["每股营收"] = guard(f"{T('rev', y)}*100/{sh}", [T("rev", y), sh])
            f["每股 FCF"] = guard(f"{fcf}*100/{sh}", [fcf, sh])
            f["有效税率"] = guard(f'IF({T("pti", y)}<=0,"",{T("tax", y)}/{T("pti", y)})', [T("tax", y), T("pti", y)])
            tr = f'IF(N({T("pti", y)})>0,MIN(MAX(N({T("tax", y)})/{T("pti", y)},0),0.35),0.21)'
            f["ROIC"] = guard(f"{T('op', y)}*(1-{tr})/({T('eq_t', y)}+N({T('std', y)})+N({T('ltd', y)}))",
                              [T("op", y), T("eq_t", y)])
            f[f"EBITDA（{self.U}）"] = guard(f"{T('op', y)}+{T('da', y)}", [T("op", y), T("da", y)])
            f["OCF / 净利润"] = guard(f'IF({T("ni", y)}<=0,"",{T("ocf", y)}/{T("ni", y)})', [T("ocf", y), T("ni", y)])
            f["应收天数"] = guard(f"{T('ar', y)}/{T('rev', y)}*365", [T("ar", y), T("rev", y)])
            f["存货天数"] = guard(f"{T('inv', y)}/{T('cogs', y)}*365", [T("inv", y), T("cogs", y)])
            f["应付天数"] = guard(f"{T('ap', y)}/{T('cogs', y)}*365", [T("ap", y), T("cogs", y)])
            for name, fmt in cols:
                if name in f and f[name]:
                    ws[cell(name)] = f[name]
                if fmt:
                    ws[cell(name)].number_format = fmt
            notes = []
            sb, err = ex.shares(ex.aper[y], "b")
            if err is not None and err > 0.01:
                notes.append(f"股数由每股收益推算，每股收益 {ex.core('eps_b', ex.aper[y])[2]} 的小数位少，股数（及市值）最大舍入误差约 ±{err:.1%}")
            if p is not None:
                sp = ex.shares(ex.aper[p], "b")[0]
                if sb and sp and not (0.5 <= sb / sp <= 2):
                    notes.append(f"股数比值 {sb / sp:.2f} 超出 0.5–2：疑似拆股 / 合股或推算失真，"
                                 "日线为拆股调整价，富途每股收益也按拆股调整，需人工核对")
            ws[cell("备注")] = "；".join(notes) or None
            ws[cell("来源ID")] = "F1"
        for name, _ in cols:
            ws.column_dimensions[self.dcol[name]].width = 12
        ws.column_dimensions[self.dcol["备注"]].width = 50

    def Dv(self, name, y):
        return f"{q('年度衍生指标')}{self.dcol[name]}{self.drow[y]}"

    # ---- 历年营收与业务拆分
    def sheet_revenue(self, nseg=6):
        ws, ex = self.ws["历年营收与业务拆分"], self.ex
        U = self.U
        title(ws, "历年营收与业务拆分")
        ws["B2"] = (f"金额：{U}；占比 / 同比：%；空白 ≠ 0；新旧分类不可相加。合并营收由脚本从'年度三表'引用（层级 1）；"
                    "分部列由第二层按公司原披露分类填写（分类变更时新旧分类并列、各自成组，把占位名改成真实分部名）；"
                    "未拆分收入 = 没有任何分部数时的合并营收；现代分类勾稽差额 = 合并营收 − Σ分部营收 − 未拆分收入"
                    "（新旧分类并列的年份，第二层填写时把旧分类组从勾稽公式中去掉）。")
        ws["B2"].font = NOTE
        heads = ["期间", "期间口径", f"合并营收（{U}）", "合并同比（%）"]
        for k in range(1, nseg + 1):
            n = f"分部{k}（待填）"
            heads += [f"{n}营收（{U}）", f"{n}占比（%）", f"{n}同比（%）"]
        heads += [f"未拆分收入营收（{U}）", "未拆分收入占比（%）", "未拆分收入同比（%）", "来源ID",
                  f"现代分类勾稽差额（{U}）", "分类与同比说明"]
        hdr(ws, 3, heads)
        for k in range(nseg):
            for j in range(3):
                ws.cell(row=3, column=5 + 3 * k + j).fill = L2_FILL
        ws.freeze_panes = "C4"
        C = lambda i: get_column_letter(i)
        seg_cols = [C(5 + 3 * k) for k in range(nseg)]
        un = C(5 + 3 * nseg)
        src, diff, note = C(8 + 3 * nseg), C(9 + 3 * nseg), C(10 + 3 * nseg)
        self.rev_row = {}
        self.rev_seg_cols, self.rev_un_col = seg_cols, un
        rows = [(y, "全年") for y in self.years]
        ytd = self.ytd()
        if ytd:
            rows.append((ytd["label"], ytd["kind"]))
        for i, (y, kind) in enumerate(rows):
            r = 4 + i
            ws[f"A{r}"], ws[f"B{r}"] = y, kind
            if isinstance(y, int):
                self.rev_row[y] = r
                ws[f"C{r}"] = guard(self.T("rev", y), [self.T("rev", y)])
                prev = self.rev_row.get(y - 1)
                if prev:
                    ws[f"D{r}"] = guard(f"C{r}/C{prev}-1", [f"C{r}", f"C{prev}"])
                    for sc in seg_cols + [un]:
                        c2 = C(column_index_from_string(sc) + 2)
                        ws[f"{c2}{r}"] = guard(f"{sc}{r}/{sc}{prev}-1", [f"{sc}{r}", f"{sc}{prev}"])
                ws[f"{note}{r}"] = "业务未拆分（分部待第二层填写）；未披露值留空，不按 0 计算。"
            else:
                ws[f"C{r}"] = ytd["value"] / 1e8
                if ytd.get("prior"):
                    ws[f"D{r}"] = ytd["value"] / ytd["prior"] - 1
                ws[f"{note}{r}"] = (f"{ytd['desc']}；同比基数为上年同期 {ytd.get('prior', 0) / 1e8:,.2f} {U}"
                                    if ytd.get("prior") else ytd["desc"])
            segs = [f"{sc}{r}" for sc in seg_cols]
            for sc in seg_cols:
                c1 = C(column_index_from_string(sc) + 1)
                ws[f"{c1}{r}"] = guard(f"{sc}{r}/C{r}", [f"{sc}{r}", f"C{r}"])
            ws[f"{un}{r}"] = f'=IF(COUNT(C{r})=0,"",IF(COUNT({",".join(segs)})=0,C{r},""))'
            ws[f"{C(column_index_from_string(un) + 1)}{r}"] = guard(f"{un}{r}/C{r}", [f"{un}{r}", f"C{r}"])
            ws[f"{src}{r}"] = "F1"
            ws[f"{diff}{r}"] = f'=IF(COUNT(C{r})=0,"",C{r}-SUM({",".join(segs)})-N({un}{r}))'
            for col in range(3, 10 + 3 * nseg):
                h = heads[col - 1]
                ws.cell(row=r, column=col).number_format = F_PCT if "%" in h else F_AMT
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions[note].width = 50
        self.rev_last_row = 3 + len(rows)

    def ytd(self):
        ex = self.ex
        if not ex.quarters or not self.years:
            return None
        last_fy = self.years[-1]
        qs = [p for p in ex.quarters if ex.meta[p]["fy"] > last_fy]
        if not qs:
            return None
        fy = ex.meta[qs[-1]]["fy"]
        qs = [p for p in qs if ex.meta[p]["fy"] == fy]
        vals = [ex.core("rev", p)[0] for p in qs]
        if any(v is None for v in vals):
            return None
        nums = sorted(int(ex.meta[p]["type"]) for p in qs)
        prior = []
        for n in nums:
            pp = next((p for p, m in ex.meta.items() if m["fy"] == fy - 1 and m["type"] == str(n)), None)
            prior.append(ex.core("rev", pp)[0] if pp else None)
        lab = f"FY{fy} 年初至今（Q{nums[0]}–Q{nums[-1]}）" if len(nums) > 1 else f"FY{fy} Q1"
        return {"label": lab, "kind": f"年初至今（{len(nums)} 个季度）", "value": sum(vals),
                "prior": sum(prior) if all(v is not None for v in prior) else None,
                "desc": f"最新年初至今口径：FY{fy} 第 {nums[0]}–{nums[-1]} 季单季营收合计，截至 {ex.meta[qs[-1]]['end']}"}

    # ---- 近12季度明细
    def sheet_quarters(self):
        ws, ex = self.ws["近12季度明细"], self.ex
        U = self.U
        qs = ex.quarters[-12:]
        self.q12 = qs
        title(ws, f"{ex.code} 近 12 个季度明细（{ex.label(qs[0])}–{ex.label(qs[-1])}）" if qs else "近 12 个季度明细",
              f"每季一列，单季数（富途报告类型 1–4 均为单季）；金额 {U}；比率 %；环比、同比公式写在单元格；"
              "上年同期为富途同一财季的数（同比基数）；第一季没有上一季，环比留空；分部逐季由第二层填写。")
        hdr(ws, 3, ["指标", "单位"] + [ex.label(p) for p in qs] + ["来源ID"])
        ws.freeze_panes = "C4"
        qc = {p: get_column_letter(3 + i) for i, p in enumerate(qs)}
        srcc = get_column_letter(3 + len(qs))
        self.qcol = qc
        ws["A4"], ws["B4"] = "期末日", "日期"
        for p in qs:
            ws[f"{qc[p]}4"] = dt.date.fromisoformat(ex.meta[p]["end"])
            ws[f"{qc[p]}4"].number_format = F_DATE
        self.qrow = {}
        r = 5

        def section(text):
            nonlocal r
            ws.cell(row=r, column=1, value=text).font = BOLD
            ws.cell(row=r, column=1).fill = SEC_FILL
            r += 1

        def put_vals(key, name, getter, unit, fmt, tag=""):
            nonlocal r
            ws[f"A{r}"], ws[f"B{r}"] = name, unit
            for p in qs:
                v = getter(p)
                if v is not None:
                    ws[f"{qc[p]}{r}"] = v
                ws[f"{qc[p]}{r}"].number_format = fmt
            ws[f"{srcc}{r}"] = "F1"
            self.qrow[tag + key] = r
            r += 1

        def put_f(key, name, fn, unit, fmt):
            nonlocal r
            ws[f"A{r}"], ws[f"B{r}"] = name, unit
            for i, p in enumerate(qs):
                v = fn(p, i)
                if v:
                    ws[f"{qc[p]}{r}"] = v
                ws[f"{qc[p]}{r}"].number_format = fmt
            ws[f"{srcc}{r}"] = "F1"
            self.qrow[key] = r
            r += 1

        def prior_q(p):
            m = ex.meta[p]
            return next((pp for pp, mm in ex.meta.items() if mm["fy"] == m["fy"] - 1 and mm["type"] == m["type"]), None)

        amt = lambda key, pp: (ex.core(key, pp)[0] / 1e8 if pp and ex.core(key, pp)[0] is not None else None)
        A = lambda key, p, i: f"{qc[p]}{self.qrow[key]}"
        section("本季")
        base = [("rev", "营收"), ("gp", "毛利"), ("op", "营业利润"), ("ni", "净利润（归母）"),
                ("ocf", "经营现金流"), ("capex", "资本开支"), ("nic", "归属普通股股东净利润")]
        for key, name in base:
            put_vals(key, name, lambda p, k=key: amt(k, p), U, F_AMT)
        put_vals("eps_d", "稀释每股收益", lambda p: ex.core("eps_d", p)[0], f"{CUR_NAME.get(self.cur, self.cur)}/股", "0.0000")
        put_f("fcf", "FCF（经营现金流 + 资本开支）",
              lambda p, i: guard(f"{A('ocf', p, i)}+{A('capex', p, i)}", [A("ocf", p, i), A("capex", p, i)]), U, F_AMT)
        put_f("sh_d", "稀释加权股数（推算）",
              lambda p, i: guard(f'IF({A("eps_d", p, i)}=0,"",{A("nic", p, i)}*100/{A("eps_d", p, i)})',
                                 [A("nic", p, i), A("eps_d", p, i)]), "百万股", F_SH)
        for key, name, a in [("gm", "毛利率", "gp"), ("opm", "营业利润率", "op"), ("npm", "净利率", "ni")]:
            put_f(key, name, lambda p, i, a=a: guard(f"{A(a, p, i)}/{A('rev', p, i)}", [A(a, p, i), A("rev", p, i)]),
                  "%", F_PCT)
        section("环比（本季 ÷ 上一季 − 1）")
        for key, name in [("rev", "营收"), ("gp", "毛利"), ("op", "营业利润"), ("ni", "净利润（归母）"),
                          ("ocf", "经营现金流"), ("fcf", "FCF"), ("eps_d", "稀释每股收益")]:
            def fn(p, i, key=key):
                if i == 0:
                    return None
                a, b = A(key, p, i), f"{qc[qs[i - 1]]}{self.qrow[key]}"
                return guard(f'IF({b}<=0,"",{a}/{b}-1)', [a, b])
            put_f("qoq_" + key, f"{name}环比", fn, "%", F_PCT)
        section("上年同期（同比基数）")
        for key, name in base:
            put_vals(key, name + "（上年同期）", lambda p, k=key: amt(k, prior_q(p)), U, F_AMT, tag="py_")
        put_vals("eps_d", "稀释每股收益（上年同期）",
                 lambda p: ex.core("eps_d", prior_q(p))[0] if prior_q(p) else None,
                 f"{CUR_NAME.get(self.cur, self.cur)}/股", "0.0000", tag="py_")
        put_f("py_fcf", "FCF（上年同期）",
              lambda p, i: guard(f"{A('py_ocf', p, i)}+{A('py_capex', p, i)}", [A("py_ocf", p, i), A("py_capex", p, i)]),
              U, F_AMT)
        section("同比（本季 ÷ 上年同期 − 1）")
        for key, name in [("rev", "营收"), ("gp", "毛利"), ("op", "营业利润"), ("ni", "净利润（归母）"),
                          ("ocf", "经营现金流"), ("fcf", "FCF"), ("eps_d", "稀释每股收益")]:
            put_f("yoy_" + key, f"{name}同比",
                  lambda p, i, key=key: guard(f'IF({A("py_" + key, p, i)}<=0,"",{A(key, p, i)}/{A("py_" + key, p, i)}-1)',
                                              [A(key, p, i), A("py_" + key, p, i)]), "%", F_PCT)
        section("分部逐季（第二层填写；占比 = 分部营收 ÷ 营收）")
        for k in range(1, 4):
            put_vals(f"seg{k}", f"分部{k}（待填）营收", lambda p: None, U, F_AMT)
            ws.cell(row=r - 1, column=1).fill = L2_FILL
            put_f(f"segp{k}", f"分部{k}（待填）占比",
                  lambda p, i, k=k: guard(f"{A(f'seg{k}', p, i)}/{A('rev', p, i)}", [A(f"seg{k}", p, i), A("rev", p, i)]),
                  "%", F_PCT)
        ws.column_dimensions["A"].width = 30
        for p in qs:
            ws.column_dimensions[qc[p]].width = 12

    # ---- 市值与估值
    def sheet_valuation(self):
        ws, ex = self.ws["市值与估值"], self.ex
        U = self.U
        usd = self.usd
        title(ws, f"{ex.code} 市值与估值（逐财年年末）",
              "收盘价 = 财年期末日当天或之前最后一个交易日的收盘价（拆股调整价，美元）；股数 = 长表有期末股数的年份用期末股数，"
              "没有的年份退回当年基本加权股数（'年度三表'推算，富途每股收益已按拆股调整，与拆股调整价同口径），'股数口径'列逐年注明；"
              "市值 = 收盘价 × 股数 ÷ 100（亿美元）；"
              "总债务 = 短期债务 + 长期债务（不含租赁，缺数按 0）；现金 = 现金及短期投资；EV = 市值 + 总债务 − 现金；"
              "P/E 用当年稀释每股收益，亏损年份留空；年内最高 / 最低市值 = 财年内每日收盘价 × 当年股数。"
              + ("" if usd else f"财报为 {self.cur} 原币、股价为美元，未换算：EV 与各倍数留空。"))
        cols = ["年度", "财年期末日", "年末交易日", "收盘价（美元）", "股数（百万股）", "股数口径", "市值（亿美元）",
                f"总债务（{U}）", f"现金及短期投资（{U}）", "企业价值（亿美元）", f"稀释每股收益（{CUR_NAME.get(self.cur, self.cur)}）",
                "P/E", f"营收（{U}）", "P/S", f"EBITDA（{U}）", "EV/EBITDA", f"FCF（{U}）", "FCF 收益率",
                "年内最高市值（亿美元）", "最高日", "年内最低市值（亿美元）", "最低日", "数据提供方市值（亿美元）",
                "差异", "备注", "来源ID"]
        hdr(ws, 4, cols)
        V = {c: get_column_letter(i + 1) for i, c in enumerate(cols)}
        self.vcol = V
        self.vrow = {}
        ws.freeze_panes = "C5"
        self.mcap_py = {}
        for i, y in enumerate(self.years):
            r = 5 + i
            self.vrow[y] = r
            c = lambda name: f"{V[name]}{r}"
            end = dt.date.fromisoformat(ex.meta[ex.aper[y]]["end"])
            ws[c("年度")] = y
            ws[c("财年期末日")] = end
            ws[c("财年期末日")].number_format = F_DATE
            pt = price_on_or_before(self.px, end)
            notes = []
            if pt:
                ws[c("年末交易日")] = pt[0]
                ws[c("年末交易日")].number_format = F_DATE
                ws[c("收盘价（美元）")] = pt[1]
            else:
                notes.append("日线未覆盖该财年末")
            es, es_item, es_why = ex.end_shares(ex.aper[y])
            if es:
                ws[c("股数（百万股）")] = es
                ws[c("股数口径")] = f"期末股数（{es_item}）"
                self.share_basis[y] = "期末"
            else:
                ws[c("股数（百万股）")] = guard(self.T("sh_b", y), [self.T("sh_b", y)])
                ws[c("股数口径")] = "基本加权股数（推算）"
                self.share_basis[y] = "加权"
                if es_why:
                    notes.append(es_why)
            P, S = c("收盘价（美元）"), c("股数（百万股）")
            ws[c("市值（亿美元）")] = guard(f"{P}*{S}/100", [P, S])
            ws[c(f"总债务（{U}）")] = f"=N({self.T('std', y)})+N({self.T('ltd', y)})"
            ws[c(f"现金及短期投资（{U}）")] = guard(self.T("cash_sti", y), [self.T("cash_sti", y)])
            M, Dbt, Cs = c("市值（亿美元）"), c(f"总债务（{U}）"), c(f"现金及短期投资（{U}）")
            E = c(f"稀释每股收益（{CUR_NAME.get(self.cur, self.cur)}）")
            ws[E] = guard(self.T("eps_d", y), [self.T("eps_d", y)])
            ws[c(f"营收（{U}）")] = guard(self.T("rev", y), [self.T("rev", y)])
            ws[c(f"EBITDA（{U}）")] = guard(self.Dv(f"EBITDA（{U}）", y), [self.Dv(f"EBITDA（{U}）", y)])
            ws[c(f"FCF（{U}）")] = guard(self.Dv(f"FCF（{U}）", y), [self.Dv(f"FCF（{U}）", y)])
            if usd:
                ws[c("企业价值（亿美元）")] = guard(f"{M}+{Dbt}-{Cs}", [M, Cs])
                ws[c("P/E")] = guard(f'IF({E}<=0,"",{P}/{E})', [P, E])
                R = c(f"营收（{U}）")
                ws[c("P/S")] = guard(f"{M}/{R}", [M, R])
                EV, EB = c("企业价值（亿美元）"), c(f"EBITDA（{U}）")
                ws[c("EV/EBITDA")] = guard(f'IF({EB}<=0,"",{EV}/{EB})', [EV, EB])
                FC = c(f"FCF（{U}）")
                ws[c("FCF 收益率")] = guard(f"{FC}/{M}", [FC, M])
            sb = es or ex.shares(ex.aper[y], "b")[0]
            if pt and sb:
                self.mcap_py[y] = pt[1] * sb / 100
                start = dt.date.fromisoformat(ex.meta[ex.aper[y - 1]]["end"]) if (y - 1) in ex.aper else end - dt.timedelta(days=365)
                win = [(d, v) for d, v in self.px if start < d <= end]
                if win:
                    hi, lo = max(win, key=lambda x: x[1]), min(win, key=lambda x: x[1])
                    ws[c("年内最高市值（亿美元）")] = hi[1] * sb / 100
                    ws[c("最高日")] = hi[0]
                    ws[c("年内最低市值（亿美元）")] = lo[1] * sb / 100
                    ws[c("最低日")] = lo[0]
                    for k in ("最高日", "最低日"):
                        ws[c(k)].number_format = F_DATE
                    if win[0][0] > start + dt.timedelta(days=10):
                        notes.append(f"日线从 {win[0][0]} 起，年内高低点只覆盖部分财年")
            elif not sb:
                notes.append("股数缺（净利润或每股收益为 0 / 缺数）")
            serr = None if es else ex.shares(ex.aper[y], "b")[1]
            if serr is not None and serr > 0.01:
                notes.append(f"股数由每股收益推算，最大舍入误差约 ±{serr:.1%}")
            if y in self.ref_mcap:
                ws[c("数据提供方市值（亿美元）")] = self.ref_mcap[y]
                ref = c("数据提供方市值（亿美元）")
                ws[c("差异")] = guard(f"{M}/{ref}-1", [M, ref])
            ws[c("备注")] = "；".join(notes) or None
            ws[c("来源ID")] = "F1；F2" + ("；M1" if y in self.ref_mcap else "")
            for name in cols:
                if name in ("年度", "财年期末日", "年末交易日", "最高日", "最低日", "股数口径", "备注", "来源ID"):
                    continue
                fmt = F_X if name in ("P/E", "P/S", "EV/EBITDA") else (F_PCT if name in ("FCF 收益率", "差异") else
                                                                     ("0.0000" if "每股" in name else F_AMT))
                ws[c(name)].number_format = fmt
        for name in cols:
            ws.column_dimensions[V[name]].width = 12
        ws.column_dimensions[V["股数口径"]].width = 22
        ws.column_dimensions[V["备注"]].width = 40

    # ---- 市值与营收对比
    def sheet_mcap_vs_rev(self):
        ws, ex = self.ws["市值与营收对比"], self.ex
        code = ex.code
        title(ws, f"市值与营收对比（{self.years[0]}—{self.years[-1]}）",
              f"金额：亿美元（营收为 {self.U}）。本公司列由脚本引用'市值与估值'与'历年营收与业务拆分'（层级 1）；"
              "参考公司列由第二层填写：默认取同期行业龙头或市值最接近的标杆，可多家（多家时复制 C–F 列成组）；参考公司营收期间可能不同，须注明。")
        cols = ["年度", f"{code}市值", "参考公司", "参考公司市值", f"{code}营收", "参考公司营收", f"{code}同比",
                "参考公司同比", f"{code}市值营收倍数", "参考公司市值营收倍数", "市值比例", "营收比例", "来源ID"]
        hdr(ws, 3, cols)
        for col in (3, 4, 6, 8, 10, 11, 12):
            ws.cell(row=3, column=col).fill = L2_FILL
        self.mrow = {}
        for i, y in enumerate(self.years):
            r = 4 + i
            self.mrow[y] = r
            ws[f"A{r}"] = y
            mv = f"{q('市值与估值')}{self.vcol['市值（亿美元）']}{self.vrow[y]}"
            ws[f"B{r}"] = guard(mv, [mv])
            rv = f"{q('历年营收与业务拆分')}C{self.rev_row[y]}"
            ws[f"E{r}"] = guard(rv, [rv])
            yv = f"{q('历年营收与业务拆分')}D{self.rev_row[y]}"
            ws[f"G{r}"] = guard(yv, [yv])
            if i > 0:
                ws[f"H{r}"] = guard(f"F{r}/F{r - 1}-1", [f"F{r}", f"F{r - 1}"])
            ws[f"I{r}"] = guard(f"B{r}/E{r}", [f"B{r}", f"E{r}"])
            ws[f"J{r}"] = guard(f"D{r}/F{r}", [f"D{r}", f"F{r}"])
            ws[f"K{r}"] = guard(f"B{r}/D{r}", [f"B{r}", f"D{r}"])
            ws[f"L{r}"] = guard(f"E{r}/F{r}", [f"E{r}", f"F{r}"])
            ws[f"M{r}"] = "F1；F2"
            for col, fmt in zip("BDEFGHIJKL", [F_AMT, F_AMT, F_AMT, F_AMT, F_PCT, F_PCT, F_X, F_X, F_PCT, F_PCT]):
                ws[f"{col}{r}"].number_format = fmt
        # 同行横向比较（第二层，从行业报告填）
        ws["O3"] = "同行横向比较（第二层：从行业报告填；第一行为本公司最新财年，自动）"
        ws["O3"].font = BOLD
        hdr(ws, 4, ["公司", f"营收（最新财年，{self.U}）", "营业利润率", "市值（亿美元）", "来源ID"], col=15)
        y = self.years[-1]
        ws["O5"] = code
        ws["P5"] = guard(self.Dv(f"营收（{self.U}）", y), [self.Dv(f"营收（{self.U}）", y)])
        ws["Q5"] = guard(self.Dv("营业利润率", y), [self.Dv("营业利润率", y)])
        mv = f"{q('市值与估值')}{self.vcol['市值（亿美元）']}{self.vrow[y]}"
        ws["R5"] = guard(mv, [mv])
        ws["S5"] = "F1；F2"
        for r in range(5, 15):
            ws[f"P{r}"].number_format = F_AMT
            ws[f"Q{r}"].number_format = F_PCT
            ws[f"R{r}"].number_format = F_AMT
            if r > 5:
                for col in "OPQRS":
                    ws[f"{col}{r}"].fill = L2_FILL
        ws.column_dimensions["O"].width = 14
        for col in "BCDEFGHIJKLPQR":
            ws.column_dimensions[col].width = 12

    # ---- 股价事件
    def sheet_events(self):
        ws, ex = self.ws["股价事件"], self.ex
        ev = find_events(self.px)
        self.events = ev
        title(ws, f"{ex.code} 股价事件（单日 ±8% 或 21 个交易日 ±20%）",
              f"日线：{self.pxinfo['basis']}，{self.pxinfo['start']} 至 {self.pxinfo['end']}。21 日事件：连续触发的交易日（同方向、间隔 ≤ 5 天）"
              "合并为一次，取幅度绝对值最大的那天，窗口起始日 = 该日往前 21 个交易日。同日大盘幅度：没有指数数据，写'待补'。"
              "事件类型留空，由公司报告填写。右侧 K–L 列为月末收盘价，供图表总览的股价事件标注图使用。")
        hdr(ws, 4, ["日期", "触发规则", "幅度", "收盘价（美元）", "窗口起始日", "同日大盘幅度", "事件类型", "备注", "来源ID"])
        for i, e in enumerate(ev):
            r = 5 + i
            ws[f"A{r}"] = e["date"]
            ws[f"A{r}"].number_format = F_DATE
            ws[f"B{r}"] = e["type"]
            ws[f"C{r}"] = e["move"]
            ws[f"C{r}"].number_format = F_PCT
            ws[f"D{r}"] = e["close"]
            ws[f"D{r}"].number_format = "0.0000" if e["close"] < 1 else F_PX
            ws[f"E{r}"] = e["start"]
            ws[f"E{r}"].number_format = F_DATE
            ws[f"F{r}"] = "待补"
            if e.get("span") and e["span"][0] != e["span"][1]:
                ws[f"H{r}"] = f"连续触发 {e['span'][0]} 至 {e['span'][1]}"
            ws[f"I{r}"] = "F2"
        self.ev_last = 4 + max(1, len(ev))
        hdr(ws, 4, ["月末日期", "月末收盘（美元）"], col=11)
        months = OrderedDict()
        for d, v in self.px:
            months[(d.year, d.month)] = (d, v)
        for i, (d, v) in enumerate(months.values()):
            r = 5 + i
            ws[f"K{r}"] = d
            ws[f"K{r}"].number_format = F_DATE
            ws[f"L{r}"] = v
            ws[f"L{r}"].number_format = "0.0000" if v < 1 else F_PX
        self.month_last = 4 + max(1, len(months))
        for col, w in zip("ABCDEFGHIKL", [12, 9, 9, 12, 12, 12, 14, 36, 8, 12, 14]):
            ws.column_dimensions[col].width = w

    # ---- 所得税与现金税
    def sheet_tax(self):
        ws, ex = self.ws["所得税与现金税"], self.ex
        U = self.U
        title(ws, f"所得税与现金税｜{self.years[0]}—{self.years[-1]} 全年",
              f"金额：{U}；比例 / 同比：%；空白代表未取得独立披露。层级 1 自动填：所得税费用、税前利润（富途利润表）、"
              "递延所得税费用（取富途现金流量表'递延所得税'调整项，符号与费用相同，近似）、当期所得税费用（= 所得税费用 − 递延，推算）、"
              "有效税率；其余列（现金所得税、美国 / 海外拆分、递延税资产与估值备抵等）由第二层从 10-K 税务附注补。")
        base = ["所得税费用", "现金所得税", "税前利润", "当期所得税费用", "递延所得税费用", "美国税前利润", "海外税前利润",
                "美国联邦及州当期税", "海外当期税", "美国联邦及州递延税", "海外递延税", "美国联邦当期税", "美国州当期税",
                "美国联邦递延税", "美国州递延税", "美国联邦现金税", "美国州现金税", "海外现金税", "递延税资产总额",
                "递延税资产估值备抵", "递延税净资产／负债", "税务不确定性准备", "旧口径杂项州及海外税"]
        cols = ["期间", "期间口径", "年份"]
        for b in base:
            cols += [f"{b}（{U}）", f"{b}同比（%）"]
        cols += ["有效税率（%）", "现金税／税前利润（%）", "联邦现金税占比（%）", "州现金税占比（%）", "海外现金税占比（%）",
                 "来源", "补充来源", "披露口径"]
        hdr(ws, 3, cols)
        X = {c: get_column_letter(i + 1) for i, c in enumerate(cols)}
        auto = {"所得税费用", "税前利润", "当期所得税费用", "递延所得税费用"}
        for b in base:
            if b not in auto:
                ws[f"{X[f'{b}（{U}）']}3"].fill = L2_FILL
        rows = {}
        for i, y in enumerate(self.years):
            r = 4 + i
            rows[y] = r
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = y, "全年", y
            c = lambda name: f"{X[name]}{r}"
            ws[c(f"所得税费用（{U}）")] = guard(self.T("tax", y), [self.T("tax", y)])
            ws[c(f"税前利润（{U}）")] = guard(self.T("pti", y), [self.T("pti", y)])
            ws[c(f"递延所得税费用（{U}）")] = guard(self.T("dtax", y), [self.T("dtax", y)])
            a, d = c(f"所得税费用（{U}）"), c(f"递延所得税费用（{U}）")
            ws[c(f"当期所得税费用（{U}）")] = guard(f"{a}-{d}", [a, d])
            pti, cash = c(f"税前利润（{U}）"), c(f"现金所得税（{U}）")
            ws[c("有效税率（%）")] = guard(f'IF({pti}<=0,"",{a}/{pti})', [a, pti])
            ws[c("现金税／税前利润（%）")] = guard(f'IF({pti}<=0,"",{cash}/{pti})', [cash, pti])
            for b in base:
                col = X[f"{b}（{U}）"]
                ws[f"{col}{r}"].number_format = F_AMT
                yc = X[f"{b}同比（%）"]
                if i > 0 and self.years[i - 1] == y - 1:
                    ws[f"{yc}{r}"] = guard(f'IF({col}{r - 1}<=0,"",{col}{r}/{col}{r - 1}-1)', [f"{col}{r}", f"{col}{r - 1}"])
                ws[f"{yc}{r}"].number_format = F_PCT
            for k in ("有效税率（%）", "现金税／税前利润（%）", "联邦现金税占比（%）", "州现金税占比（%）", "海外现金税占比（%）"):
                ws[c(k)].number_format = F_PCT
            ws[c("来源")] = "F1"
            ws[c("披露口径")] = "富途最新口径；当期所得税费用为推算（所得税费用 − 现金流量表递延所得税）；现金税等待第二层。"

    # ---- 第二层空表
    def sheets_layer2(self):
        L2 = "第二层：从券商模型、公司 IR 数据包、公告补；补不到的留空并记入'缺口'。"
        ws = self.ws["IPO指标"]
        title(ws, "IPO指标", L2 + "价格与股数使用上市当时口径，每项附来源。")
        hdr(ws, 3, ["上市节点指标", "数值与含义", "来源ID"])
        for i, k in enumerate(["上市日期", "发行价与发行股数", "募资额（总额 / 净额）", "发行后总股本", "发行价对应市值",
                               "首日收盘与首日涨幅", "首日收盘对应市值", "上市时最近财年营收", "上市时市值营收倍数"]):
            ws[f"A{4 + i}"] = k
            ws[f"B{4 + i}"].fill = L2_FILL
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 60
        for name, heads, note in [
            ("重大融资时间线", ["日期", "融资类型", "金额与币种", "票息、期限与口径", "来源"],
             "金额保留原币与本金 / 净款口径。"),
            ("重要收购", ["时间与状态", "收购对象", "金额口径（交易价值 / 净现金 / 终止费分列）", "后续处置", "原始依据"],
             "金额区分交易价值、净现金与终止费。"),
            ("指引与实际", ["期间", "指标", "指引下限", "指引上限", "指引发布日期", "实际", "偏差方向", "来源ID"],
             "逐年或逐季；偏差方向 = 实际高于上限 / 区间内 / 低于下限。")]:
            ws = self.ws[name]
            title(ws, name, L2 + note)
            hdr(ws, 3, heads)
            for col in range(1, len(heads) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 22

    # ---- 图表计算 + 图表总览
    def sheet_calc_and_charts(self):
        calc, ov, ex = self.ws["图表计算"], self.ws["图表总览"], self.ex
        U = self.U
        title(calc, "图表计算（所有图表的数据区）",
              "全部是公式，引用前面各表；年度块按'图表总览'B2:D2 的起止年份筛选；无数显示 #N/A（图表中为断点）。不要在这里手填数字。")
        ov["A1"] = f"{ex.code} 图表总览"
        ov["A1"].font = TITLE
        ov["A2"], ov["C2"] = "起始年份", "结束年份"
        ov["B2"], ov["D2"] = max(self.years[0], self.years[-1] - 19), self.years[-1]
        for k in ("B2", "D2"):
            ov[k].fill = L2_FILL
            ov[k].font = BOLD
        ov["A3"] = (f"使用方法：改 B2、D2 的年份（数据范围 {self.years[0]}–{self.years[-1]}，默认显示最近 20 年，一次最多 {len(self.years)} 年），"
                    "年度图（第 1–8、11、12 张）随之更新；近 12 季图固定为最新 12 个季度；"
                    "分部结构图、参考公司市值图、同行比较图在第二层填好后自动出现。估值倍数图不画 > 100 倍的点。数据区在'图表计算'表。")
        ov["A3"].font = NOTE
        S, E = f"{q('图表总览')}$B$2", f"{q('图表总览')}$D$2"
        n = len(self.years)
        r0 = 5
        rN = r0 + n - 1
        NA = lambda ref: f'IF({ref}="",NA(),{ref})'

        def lookup(sheet, col, yr_col, first, last, a, neg=False):
            rng = f"{q(sheet)}${col}${first}:${col}${last}"
            yrs = f"{q(sheet)}${yr_col}${first}:${yr_col}${last}"
            idx = f"INDEX({rng},MATCH({a},{yrs},0))"
            return f'=IF({a}="",NA(),IFERROR(IF({idx}="",NA(),{"-" if neg else ""}{idx}),NA()))'

        def lookup_row(key, a, neg=False):
            r = self.trow[key]
            c1, c2 = self.ycol[self.years[0]], self.ycol[self.years[-1]]
            idx = f"INDEX({q('年度三表')}${c1}${r}:${c2}${r},MATCH({a},{q('年度三表')}${c1}$4:${c2}$4,0))"
            return f'=IF({a}="",NA(),IFERROR(IF({idx}="",NA(),{"-" if neg else ""}{idx}),NA()))'

        def cap(f, hi=100):
            body = f[1:]
            return f"=IFERROR(IF({body}>{hi},NA(),{body}),NA())"

        d_first, d_last = 5, 4 + len(self.years)
        v_first, v_last = 5, 4 + len(self.years)
        m_first, m_last = 4, 3 + len(self.years)
        rv_first, rv_last = 4, self.rev_last_row
        D, V = self.dcol, self.vcol
        annual = [
            ("营收", lambda a: lookup("年度衍生指标", D[f"营收（{U}）"], "A", d_first, d_last, a)),
            ("营收同比", lambda a: lookup("年度衍生指标", D["营收同比"], "A", d_first, d_last, a)),
            ("毛利率", lambda a: lookup("年度衍生指标", D["毛利率"], "A", d_first, d_last, a)),
            ("营业利润率", lambda a: lookup("年度衍生指标", D["营业利润率"], "A", d_first, d_last, a)),
            ("净利率", lambda a: lookup("年度衍生指标", D["净利率"], "A", d_first, d_last, a)),
            ("经营现金流", lambda a: lookup("年度衍生指标", D[f"经营现金流（{U}）"], "A", d_first, d_last, a)),
            ("资本开支（正数）", lambda a: lookup("年度衍生指标", D[f"资本开支（{U}）"], "A", d_first, d_last, a, neg=True)),
            ("资本开支", lambda a: lookup("年度衍生指标", D[f"资本开支（{U}）"], "A", d_first, d_last, a)),
            ("收购（业务交易净额）", lambda a: lookup_row("acq", a)),
            ("分红", lambda a: lookup_row("div", a)),
            ("普通股发行/回购净额", lambda a: lookup_row("buyback", a)),
            ("稀释加权股数（百万股）", lambda a: lookup("年度衍生指标", D["稀释加权股数（百万股）"], "A", d_first, d_last, a)),
            ("市值（亿美元）", lambda a: lookup("市值与估值", V["市值（亿美元）"], "A", v_first, v_last, a)),
            ("P/E", lambda a: cap(lookup("市值与估值", V["P/E"], "A", v_first, v_last, a))),
            ("P/S", lambda a: cap(lookup("市值与估值", V["P/S"], "A", v_first, v_last, a))),
            ("EV/EBITDA", lambda a: cap(lookup("市值与估值", V["EV/EBITDA"], "A", v_first, v_last, a))),
            ("参考公司市值", lambda a: lookup("市值与营收对比", "D", "A", m_first, m_last, a)),
        ]
        for k, sc in enumerate(self.rev_seg_cols):
            annual.append((f"分部{k + 1}", lambda a, sc=sc: lookup("历年营收与业务拆分", sc, "A", rv_first, rv_last, a)))
        annual.append(("未拆分收入", lambda a: lookup("历年营收与业务拆分", self.rev_un_col, "A", rv_first, rv_last, a)))
        calc["A3"] = "年度块（按起止年份筛选）"
        calc["A3"].font = BOLD
        hdr(calc, 4, ["年度"] + [a[0] for a in annual])
        ac = {name: get_column_letter(2 + i) for i, (name, _) in enumerate(annual)}
        for i in range(n):
            r = r0 + i
            calc[f"A{r}"] = f'=IF({S}+{i}>{E},"",{S}+{i})'
            for name, fn in annual:
                calc[f"{ac[name]}{r}"] = fn(f"$A{r}")
        # 季度块
        qc0 = 2 + len(annual) + 1
        Q = lambda j: get_column_letter(qc0 + j)
        calc[f"{Q(0)}3"] = "近 12 季块"
        calc[f"{Q(0)}3"].font = BOLD
        qnames = ["季度", "营收", "营收同比", "毛利率", "营业利润率", "净利率"]
        hdr(calc, 4, qnames, col=qc0)
        qkeys = [None, "rev", "yoy_rev", "gm", "opm", "npm"]
        for i, p in enumerate(self.q12):
            r = r0 + i
            calc[f"{Q(0)}{r}"] = f"={q('近12季度明细')}{self.qcol[p]}3"
            for j, key in enumerate(qkeys[1:], start=1):
                ref = f"{q('近12季度明细')}{self.qcol[p]}{self.qrow[key]}"
                calc[f"{Q(j)}{r}"] = f"={NA(ref)}"
        qN = r0 + max(1, len(self.q12)) - 1
        # 月末收盘 + 事件
        mc0 = qc0 + len(qnames) + 1
        M = lambda j: get_column_letter(mc0 + j)
        calc[f"{M(0)}3"] = "股价块（月末收盘与事件，按起止年份筛选）"
        calc[f"{M(0)}3"].font = BOLD
        hdr(calc, 4, ["月末日期", "月末收盘", "事件日期", "事件日收盘", "事件幅度"], col=mc0)
        nm = self.month_last - 4
        for i in range(nm):
            r, sr = r0 + i, 5 + i
            dref = f"{q('股价事件')}K{sr}"
            calc[f"{M(0)}{r}"] = f"={dref}"
            calc[f"{M(0)}{r}"].number_format = F_DATE
            calc[f"{M(1)}{r}"] = f"=IF(OR(YEAR({dref})<{S},YEAR({dref})>{E}),NA(),{q('股价事件')}L{sr})"
        ne = max(1, len(self.events))
        for i in range(ne):
            r, sr = r0 + i, 5 + i
            dref = f"{q('股价事件')}A{sr}"
            calc[f"{M(2)}{r}"] = f"={dref}"
            calc[f"{M(2)}{r}"].number_format = F_DATE
            calc[f"{M(3)}{r}"] = f'=IF({dref}="",NA(),IF(OR(YEAR({dref})<{S},YEAR({dref})>{E}),NA(),{q("股价事件")}D{sr}))'
            calc[f"{M(4)}{r}"] = f"={NA(q('股价事件') + 'C' + str(sr))}"
            calc[f"{M(4)}{r}"].number_format = F_PCT
        # 同行块
        pc0 = mc0 + 6
        P = lambda j: get_column_letter(pc0 + j)
        calc[f"{P(0)}3"] = "同行横向比较块"
        calc[f"{P(0)}3"].font = BOLD
        hdr(calc, 4, ["公司", "营收", "营业利润率", "市值"], col=pc0)
        for i in range(10):
            r, sr = r0 + i, 5 + i
            calc[f"{P(0)}{r}"] = f'=IF({q("市值与营收对比")}O{sr}="","",{q("市值与营收对比")}O{sr})'
            for j, col in enumerate("PQR", start=1):
                calc[f"{P(j)}{r}"] = f"={NA(q('市值与营收对比') + col + str(sr))}"
        for r in range(r0, rN + 1):
            for name in ("营收同比", "毛利率", "营业利润率", "净利率"):
                calc[f"{ac[name]}{r}"].number_format = F_PCT
        for r in range(r0, qN + 1):
            for j in range(2, 6):
                calc[f"{Q(j)}{r}"].number_format = F_PCT

        # ---- 图表
        cats = Reference(calc, min_col=1, min_row=r0, max_row=rN)
        col_idx = lambda name: column_index_from_string(ac[name])

        def data(names, ws=calc, row0=r0, rowN=rN, idx=col_idx):
            return [Reference(ws, min_col=idx(nm), min_row=4, max_row=rowN) for nm in names]

        def style(ch, t, ytitle=None, yfmt=None):
            ch.title = t
            ch.display_blanks = "gap"
            for s in ch.series:
                s.smooth = False
            ch.width, ch.height = 17, 8.5
            ch.legend.position = "b"
            ch.x_axis.delete = False
            ch.y_axis.delete = False
            if ytitle:
                ch.y_axis.title = ytitle
            if yfmt:
                ch.y_axis.number_format = yfmt
            return ch

        def bar(t, names, ytitle=None, stacked=False, catref=cats, refs=None, yfmt=None):
            ch = BarChart()
            ch.type = "col"
            for ref in (refs or data(names)):
                ch.add_data(ref, titles_from_data=True)
            ch.set_categories(catref)
            if stacked:
                ch.grouping, ch.overlap = "stacked", 100
            return style(ch, t, ytitle, yfmt)

        def line(t, names, ytitle=None, catref=cats, refs=None, yfmt=None):
            ch = LineChart()
            for ref in (refs or data(names)):
                ch.add_data(ref, titles_from_data=True)
            ch.set_categories(catref)
            return style(ch, t, ytitle, yfmt)

        def combo(t, bnames, lnames, ytitle, y2fmt, catref=cats, brefs=None, lrefs=None):
            b = bar(t, bnames, ytitle, catref=catref, refs=brefs)
            ln = LineChart()
            for ref in (lrefs or data(lnames)):
                ln.add_data(ref, titles_from_data=True)
            for s in ln.series:
                s.smooth = False
            ln.y_axis.axId = 200
            ln.y_axis.crosses = "max"
            ln.y_axis.number_format = y2fmt
            ln.y_axis.delete = False
            ln.y_axis.majorGridlines = None
            b += ln
            return b

        charts = []
        charts.append(combo(f"1 营收（{U}）与同比", ["营收"], ["营收同比"], U, "0%"))
        charts.append(bar(f"2 分部结构（{U}，第二层填好后出现）", [f"分部{k + 1}" for k in range(len(self.rev_seg_cols))] + ["未拆分收入"], U, stacked=True))
        charts.append(line("3 利润率带", ["毛利率", "营业利润率", "净利率"], yfmt="0%"))
        charts.append(bar(f"4 现金流去向（{U}，流出为负）", ["经营现金流", "资本开支", "收购（业务交易净额）", "分红", "普通股发行/回购净额"], U))
        charts.append(line(f"5 资本开支 vs 经营现金流（{U}）", ["经营现金流", "资本开支（正数）"], U))
        charts.append(combo("6 稀释股数（百万股）与发行 / 回购净额", ["稀释加权股数（百万股）"], ["普通股发行/回购净额"], "百万股", "#,##0"))
        charts.append(line(f"7 市值（亿美元）vs 营收（{U}）", ["市值（亿美元）", "营收"]))
        charts.append(line("8 估值倍数带", ["P/E", "P/S", "EV/EBITDA"], yfmt="0.0"))
        qidx = lambda nm: qc0 + qnames.index(nm)
        qcats = Reference(calc, min_col=qc0, min_row=r0, max_row=qN)
        charts.append(combo(f"9 近 12 季营收（{U}）与同比", None, None, U, "0%", catref=qcats,
                            brefs=data(["营收"], rowN=qN, idx=qidx), lrefs=data(["营收同比"], rowN=qN, idx=qidx)))
        charts.append(line("10 近 12 季利润率", None, catref=qcats, refs=data(["毛利率", "营业利润率", "净利率"], rowN=qN, idx=qidx), yfmt="0%"))
        sc = ScatterChart()
        sc.style = 13
        xs = Reference(calc, min_col=mc0, min_row=r0, max_row=r0 + nm - 1)
        ys = Reference(calc, min_col=mc0 + 1, min_row=r0, max_row=r0 + nm - 1)
        s1 = Series(ys, xs, title="月末收盘（拆股调整，美元）")
        s1.marker.symbol = "none"
        sc.series.append(s1)
        xe = Reference(calc, min_col=mc0 + 2, min_row=r0, max_row=r0 + ne - 1)
        ye = Reference(calc, min_col=mc0 + 3, min_row=r0, max_row=r0 + ne - 1)
        s2 = Series(ye, xe, title="股价事件（单日 ±8% / 21 日 ±20%）")
        s2.marker.symbol = "circle"
        s2.marker.size = 5
        s2.marker.graphicalProperties = GraphicalProperties(solidFill="C00000")
        s2.marker.graphicalProperties.line.solidFill = "C00000"
        s2.graphicalProperties.line.noFill = True
        s1.graphicalProperties.line.solidFill = "1F4E79"
        s1.graphicalProperties.line.width = 15000
        sc.series.append(s2)
        sc.x_axis.number_format = "yyyy"
        sc.y_axis.scaling.logBase = 10
        style(sc, "11 股价与事件标注（对数坐标）", "美元")
        charts.append(sc)
        charts.append(line("12 市值对比：本公司 vs 参考公司（亿美元，参考公司第二层）", ["市值（亿美元）", "参考公司市值"]))
        pidx = lambda nm: pc0 + ["公司", "营收", "营业利润率", "市值"].index(nm)
        pcats = Reference(calc, min_col=pc0, min_row=r0, max_row=r0 + 9)
        charts.append(bar("13 同行横向比较：营收（第二层，从行业报告填）", None, catref=pcats,
                          refs=data(["营收"], rowN=r0 + 9, idx=pidx)))
        for i, ch in enumerate(charts):
            anchor = f"{'A' if i % 2 == 0 else 'K'}{5 + (i // 2) * 18}"
            ov.add_chart(ch, anchor)
        self.expected_charts = len(charts)
        ov.column_dimensions["A"].width = 10

    # ---- 说明 / 缺口 / 来源
    def sheet_notes(self):
        ws, ex = self.ws["说明"], self.ex
        sp = self.special
        ends = [ex.meta[ex.aper[y]]["end"] for y in self.years[-4:]]
        months = sorted(set(int(e[5:7]) for e in ends))
        lq = ex.quarters[-1] if ex.quarters else None
        title(ws, f"{ex.code} 历史数据 Excel · 说明")
        rows = [
            ("公司", ex.code),
            ("代码", ex.code),
            ("财报币种", f"{self.cur}（{CUR_NAME.get(self.cur, self.cur)}）" + ("" if self.usd else "：原币，未换算成美元；股价、市值为美元")),
            ("股价币种", "美元（美股交易价）"),
            ("财年口径", f"FY 标签 = 年报期末日所在年份（期末日落在 1 月 1–7 日的算上一年），与富途'财年'列一致；"
                     f"近 4 个年报期末日：{'、'.join(ends)}（年报期末月份：{'、'.join(f'{m} 月' for m in months)}）；季度的 FY 标签跟所属年报走。"),
            ("覆盖年限", f"年度 FY{self.years[0]}–FY{self.years[-1]}（{len(self.years)} 年）；季度 {ex.label(ex.quarters[0])}–{ex.label(lq)}（共 {len(ex.quarters)} 期）" if lq else f"年度 FY{self.years[0]}–FY{self.years[-1]}"),
            ("最新年报", f"{ex.label(ex.aper[self.years[-1]])}，期末日 {ex.meta[ex.aper[self.years[-1]]]['end']}"),
            ("最新季报（近12季度明细的最新一季）", f"{ex.label(lq)}，期末日 {ex.meta[lq]['end']}" if lq else "无季报"),
            ("报告类型编号", "7 = 全财年（年报）；1/2/3/4 = 财年第 1–4 季度的单季数（不是累计）；资产负债表为期末时点数。"),
            ("单位约定", f"金额 {self.U}（富途原值 ÷ 1e8）；股数百万股；占比与同比 %；空白 ≠ 0（富途无数或输入缺数）；新旧分类不可相加。"),
            ("数据口径", "富途导出是最新口径（含后续重列），不是首次披露；与原报对账时差异可能来自重列。"
                     + ("" if ex.date_fixed_col else "本长表没有'日期已校正'列，美股期末日可能比真实期末日早一天。")),
            ("日线", f"{self.pxinfo['basis']}；{self.pxinfo['start']} 至 {self.pxinfo['end']}"
                   + (f"；{self.pxinfo['ext']}" if self.pxinfo.get("ext") else "")),
            ("股数", "加权股数由'归属普通股股东净利润 ÷ 每股收益'推算（富途没有股数行）；每股收益小数位少的年份有舍入误差，见'年度衍生指标'备注。"),
            ("市值的股数口径", self.share_basis_text()),
            ("已知缺口", "见'缺口'表：底稿缺口原样并入；第二层各表未填；大盘同日幅度待补。"),
            ("生成", f"{self.today.isoformat()}；脚本 build_history_xlsx.py v{SCRIPT_VERSION}；AI：{self.ai}；输入：{ex.long_file}（文件日期 {ex.long_mtime}）"),
        ]
        if sp.get("note"):
            rows.insert(11, ("本公司特殊口径", sp["note"]))
        if ex.dup:
            rows.append(("长表重复行", f"同一报表、期间、科目出现多次的 {ex.dup} 行，只取第一次出现的值。"))
        hdr(ws, 3, ["项目", "内容"])
        r = 4
        for k, v in rows:
            ws[f"A{r}"], ws[f"B{r}"] = k, v
            ws[f"B{r}"].alignment = WRAP
            r += 1
        r += 1
        ws[f"A{r}"] = "各表口径与缺口"
        ws[f"A{r}"].font = BOLD
        r += 1
        hdr(ws, r, ["工作表", "口径", "缺口 / 层级"])
        r += 1
        sheets = [
            ("图表总览", "13 张原生图表，数据区全部在'图表计算'；B2、D2 起止年份可改。", "分部、参考公司、同行三张图待第二层数据"),
            ("说明", "本页。", "—"),
            ("历年营收与业务拆分", "合并营收引用'年度三表'；最后一行为最新年初至今口径。", "分部列为第二层（层级 2），占位名待改"),
            ("年度三表", "富途三表全部科目 + 富途关键指标原值；顶部为标准化核心科目（逐年取第一个有数的候选科目）。", "富途无股权薪酬行；回购只有'普通股发行/回购的净额'"),
            ("年度衍生指标", "公式写在单元格，口径写在表头说明行。", "GAAP − 非GAAP 差额为第二层"),
            ("近12季度明细", "最近 12 个单季；环比、同比公式；上年同期为同比基数。", "分部逐季为第二层"),
            ("市值与营收对比", "本公司列自动；参考公司列与同行横向比较表为第二层。", "参考公司全部待填"),
            ("市值与估值", "财年末收盘价 × 股数（有期末股数用期末股数，否则用当年基本加权股数，'股数口径'列逐年注明）；"
                        "EV = 市值 + 总债务 − 现金及短期投资。",
             "数据提供方市值" + ("已提供（M1）" if self.ref_mcap else "未提供：外部对照待补") + ("" if self.usd else "；非美元财报不算倍数")
             + ("；期末股数缺，市值用加权股数" if "加权" in self.share_basis.values() else "")),
            ("股价事件", "单日 ±8% 或 21 日 ±20%；拆股调整价。", "同日大盘幅度待补；事件类型由报告填"),
            ("IPO指标", "第二层空表。", "全部待填"),
            ("所得税与现金税", "所得税费用、税前利润、递延（现金流量表调整项）、当期（推算）、有效税率自动。", "现金税、地域拆分、递延税资产等为第二层"),
            ("重大融资时间线", "第二层空表。", "全部待填"),
            ("重要收购", "第二层空表。", "全部待填"),
            ("指引与实际", "第二层空表。", "全部待填"),
            ("缺口", "导出缺的科目与期间、第二层未填的表、可从哪补。", "—"),
            ("来源", "来源 ID → 文件、位置、定位。", "参考公司来源待第二层"),
            ("图表计算", "全部公式，引用前面各表；年度块随起止年份筛选。", "—"),
        ]
        for s, a, b in sheets:
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = s, a, b
            ws[f"B{r}"].alignment = WRAP
            ws[f"C{r}"].alignment = WRAP
            r += 1
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 100
        ws.column_dimensions["C"].width = 40

    def share_basis_text(self):
        ex = self.ex
        end_y = [y for y in self.years if self.share_basis.get(y) == "期末"]
        wt_y = [y for y in self.years if self.share_basis.get(y) == "加权"]
        span = lambda ys: f"FY{ys[0]}–FY{ys[-1]}（{len(ys)} 年）" if ys else "无"
        if not wt_y:
            return f"市值全部用期末股数：{span(end_y)}。"
        y = self.years[-1]
        par, par_nm = ex.par_amount(ex.aper[y])
        par_txt = (f"资产负债表的'{par_nm}'是面值金额（发行股数 × 每股面值，FY{y} 为 {par / 1e8:,.2f} {self.U}），不是股数；"
                   "导出里没有每股面值和库存股股数，不能由它换算出流通股数。") if par is not None else ""
        head = (f"期末股数：{span(end_y)}；其余年份退回基本加权股数：{span(wt_y)}。" if end_y else
                "长表（资产负债表、关键指标）没有期末股数类科目（候选：" + "、".join(END_SHARE_ITEMS[:6]) + " 等），"
                "市值全部用当年基本加权股数。")
        return (head + par_txt + "加权股数是全年平均，股数增长快的年份（上市当年、大量换股收购、期权行权多）会低于年末股数，"
                "市值因此偏低；与按年末流通股计算的数据提供方市值对照时，这些年份差异最大。期末股数可从 10-K 封面页"
                "（流通股数）或股东权益附注补，属第二层。")

    def sheet_gaps(self):
        ws, ex = self.ws["缺口"], self.ex
        title(ws, "缺口", "导出缺的科目与期间（底稿缺口原样并入）、第二层尚未填的表、可从哪补。")
        hdr(ws, 3, ["项目", "期间", "缺什么", "可从哪补", "层级", "来源"])
        rows = []
        covered = {"基本加权股数": "本工作簿已按'归属普通股股东净利润 ÷ 每股收益'推算，不受底稿规则限制（精度见年度衍生指标备注）",
                   "稀释加权股数": "本工作簿已按'归属普通股股东净利润 ÷ 每股收益'推算，不受底稿规则限制（精度见年度衍生指标备注）",
                   "期末收盘价": f"本工作簿用 {self.pxinfo['start']} 起的日线（F2），覆盖范围见'说明'"}
        for g in ex.gaps:
            item = (g.get("项目") or "").strip()
            how = g.get("需要谁处理") or g.get("试过哪些来源")
            if item in covered:
                how = f"{how}；{covered[item]}"
            rows.append((item, g.get("期间"), g.get("缺什么"), how, "1（导出）", "底稿缺口.csv"))
        for g in ex.checks:
            if str(g.get("通过", "")).strip() not in ("是", "通过", "True", "true", "1"):
                rows.append((g.get("检查项"), g.get("期间"), f"底稿检查未通过：{g.get('说明')}", "人工核对", "1（导出）", "底稿检查.csv"))
        for s in ["历年营收与业务拆分（分部列）", "近12季度明细（分部逐季）", "市值与营收对比（参考公司）", "IPO指标", "所得税与现金税（现金税等）",
                  "重大融资时间线", "重要收购", "指引与实际", "年度衍生指标（GAAP − 非GAAP 差额）", "图表总览（份额变迁、同行横向比较）"]:
            rows.append((s, "全部", "第二层尚未填写", "券商模型 / 公司 IR 数据包 / 10-K 附注 / 行业报告", "2", "—"))
        rows.append(("股价事件：同日大盘幅度", "全部", "没有指数日线", "yfinance 取 ^GSPC / ^IXIC 或 SOXX 日线", "1（待补）", "—"))
        if not self.ref_mcap:
            rows.append(("市值与估值：数据提供方市值", "全部", "没有外部市值对照", "companiesmarketcap 等；用 --ref-mcap 传入", "1（待补）", "—"))
        wt_y = [y for y in self.years if self.share_basis.get(y) == "加权"]
        if wt_y:
            per = "全部" if len(wt_y) == len(self.years) else "、".join(f"FY{y}" for y in wt_y)
            rows.append(("市值与估值：期末股数", per, "长表没有期末股数，市值用基本加权股数（股数增长快的年份偏低）",
                         "10-K 封面页流通股数 / 股东权益附注", "2", "F1"))
        if self.pxinfo["start"] and self.years:
            first_end = ex.meta[ex.aper[self.years[0]]]["end"]
            if self.pxinfo["start"] > first_end:
                rows.append(("日线", f"{first_end} 至 {self.pxinfo['start']}", "日线起点晚于第一个财年末，这些年份没有市值",
                             "yfinance / 券商行情", "1", "F2"))
        for i, rr in enumerate(rows):
            for j, v in enumerate(rr):
                ws.cell(row=4 + i, column=1 + j, value=v).alignment = WRAP
        for col, w in zip("ABCDEF", [28, 30, 50, 36, 10, 14]):
            ws.column_dimensions[col].width = w

    def sheet_sources(self):
        ws, ex = self.ws["来源"], self.ex
        title(ws, "来源", "来源 ID 对应各表最后一列。路径相对于 futu_export 根目录（Drive：计算机 › 我的 Mac › futu_export）。")
        ws["A3"] = "财务数据来源"
        ws["A3"].font = BOLD
        hdr(ws, 4, ["来源ID", "文件", "网址 / 位置", "表格或页码定位"])
        rows = [("F1", ex.long_file, f"{ex.code}/{ex.long_file}（富途 OpenD 导出，文件日期 {ex.long_mtime}）",
                 "报表 × 期间 × 科目；报告类型 7 = 年报，1–4 = 单季"),
                ("F2", " + ".join(x for x in [self.pxinfo.get("adj_file"), self.pxinfo.get("raw_file")] if x),
                 f"{ex.code}/", f"close 列；{self.pxinfo['basis']}" + (f"；{self.pxinfo['ext']}" if self.pxinfo.get("ext") else "")),
                ("F3", f"底稿CSV/{ex.code}_科目映射.csv、_缺口.csv、_检查.csv", f"{ex.code}/底稿CSV/", "to_digao.py 生成")]
        if self.ref_mcap:
            rows.append(("M1", "数据提供方市值（--ref-mcap）", self.ref_mcap.get("_src", ""), "年度 → 市值（亿美元）"))
        for i, rr in enumerate(rows):
            for j, v in enumerate(rr):
                ws.cell(row=5 + i, column=1 + j, value=v)
        r = 7 + len(rows)
        ws[f"A{r}"] = "市值与参考公司来源"
        ws[f"A{r}"].font = BOLD
        hdr(ws, r + 1, ["年度", f"{ex.code}营收来源", f"{ex.code}市值来源", "参考公司市值来源", "参考公司营收来源", "参考公司营收口径"])
        for i, y in enumerate(self.years):
            rr = r + 2 + i
            ws[f"A{rr}"], ws[f"B{rr}"] = y, "F1"
            ws[f"C{rr}"] = "F2 收盘价 × F1 期末股数" if self.share_basis.get(y) == "期末" else "F2 收盘价 × F1 推算基本加权股数"
        for col, w in zip("ABCDEF", [10, 40, 60, 40, 30, 30]):
            ws.column_dimensions[col].width = w

    def build(self, out_path):
        self.sheet_statements()
        self.sheet_derived()
        self.sheet_revenue()
        self.sheet_quarters()
        self.sheet_valuation()
        self.sheet_mcap_vs_rev()
        self.sheet_events()
        self.sheet_tax()
        self.sheets_layer2()
        self.sheet_calc_and_charts()
        self.sheet_notes()
        self.sheet_gaps()
        self.sheet_sources()
        self.wb.calculation.fullCalcOnLoad = True
        self.wb.save(out_path)
        return out_path


# ---------------------------------------------------------------- 公式求值（验收用，覆盖本脚本生成的公式子集）
class XErr(Exception):
    pass


TOK = re.compile(r"""\s*(?:
  (?P<str>"(?:[^"]|"")*")|
  (?P<ref>(?:'(?:[^']|'')+'!|[A-Za-z_][\w\.]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|
  (?P<num>\d+(?:\.\d*)?(?:[eE][-+]?\d+)?)|
  (?P<fn>[A-Z][A-Z0-9\.]*)\(|
  (?P<op><>|<=|>=|[-+*/^&=<>(),%])
)""", re.X)


class Evaluator:
    def __init__(self, wb):
        self.wb = wb
        self.memo = {}

    def tokens(self, s):
        out, pos = [], 0
        while pos < len(s):
            m = TOK.match(s, pos)
            if not m or m.end() == pos:
                if s[pos:].strip() == "":
                    break
                raise ValueError(f"无法解析：{s[pos:pos + 20]}")
            pos = m.end()
            k = m.lastgroup
            out.append((k, m.group(k)))
        return out

    def cell(self, sheet, coord):
        key = (sheet, coord)
        if key in self.memo:
            v = self.memo[key]
        else:
            self.memo[key] = XErr("#CIRC")
            raw = self.wb[sheet][coord].value
            if isinstance(raw, str) and raw.startswith("="):
                try:
                    v = self.formula(raw[1:], sheet)
                except XErr as e:
                    v = e
            else:
                v = raw
            self.memo[key] = v
        if isinstance(v, XErr):
            raise v
        return v

    def formula(self, s, sheet):
        self.toks, self.i, self.sheet = self.tokens(s), 0, sheet
        node = self.p_cmp()
        if self.i != len(self.toks):
            raise ValueError(f"多余的符号：{s}")
        return self.ev(node)

    # 语法分析 → 嵌套元组
    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def p_cmp(self):
        a = self.p_cat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            a = ("bin", op, a, self.p_cat())
        return a

    def p_cat(self):
        a = self.p_add()
        while self.peek()[1] == "&":
            self.take()
            a = ("bin", "&", a, self.p_add())
        return a

    def p_add(self):
        a = self.p_mul()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            a = ("bin", op, a, self.p_mul())
        return a

    def p_mul(self):
        a = self.p_pow()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            a = ("bin", op, a, self.p_pow())
        return a

    def p_pow(self):
        a = self.p_un()
        while self.peek()[1] == "^":
            self.take()
            a = ("bin", "^", a, self.p_un())
        return a

    def p_un(self):
        if self.peek()[1] in ("-", "+"):
            op = self.take()[1]
            return ("neg", self.p_un()) if op == "-" else self.p_un()
        return self.p_prim()

    def p_prim(self):
        k, v = self.take()
        if k == "num":
            return ("lit", float(v))
        if k == "str":
            return ("lit", v[1:-1].replace('""', '"'))
        if k == "ref":
            if "!" in v:
                sh, c = v.rsplit("!", 1)
                sh = sh.strip("'").replace("''", "'")
            else:
                sh, c = self.sheet, v
            return ("ref", sh, c.replace("$", ""))
        if k == "fn":
            args = []
            if self.peek()[1] != ")":
                args.append(self.p_cmp())
                while self.peek()[1] == ",":
                    self.take()
                    args.append(self.p_cmp())
            assert self.take()[1] == ")"
            return ("fn", v, args)
        if v == "(":
            a = self.p_cmp()
            assert self.take()[1] == ")"
            return a
        raise ValueError(f"意外的符号 {v}")

    # 求值
    def rng(self, node):
        _, sh, c = node
        if ":" not in c:
            return [self.cell(sh, c)]
        a, b = c.split(":")
        ws = self.wb[sh]
        out = []
        for row in ws[a:b]:
            for cc in row:
                try:
                    out.append(self.cell(sh, cc.coordinate))
                except XErr as e:
                    out.append(e)
        return out

    @staticmethod
    def num(v):
        if isinstance(v, XErr):
            raise v
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, dt.datetime):
            return (v - dt.datetime(1899, 12, 30)).days
        if isinstance(v, dt.date):
            return (v - dt.date(1899, 12, 30)).days
        raise XErr("#VALUE!")

    def ev(self, n):
        t = n[0]
        if t == "lit":
            return n[1]
        if t == "ref":
            if ":" in n[2]:
                raise XErr("#VALUE!")
            return self.cell(n[1], n[2])
        if t == "neg":
            return -self.num(self.ev(n[1]))
        if t == "bin":
            op = n[1]
            a, b = self.ev(n[2]), self.ev(n[3])
            if op in ("+", "-", "*", "/", "^"):
                x, y = self.num(a), self.num(b)
                if op == "+":
                    return x + y
                if op == "-":
                    return x - y
                if op == "*":
                    return x * y
                if op == "^":
                    return x ** y
                if y == 0:
                    raise XErr("#DIV/0!")
                return x / y
            if op == "&":
                return f"{'' if a is None else a}{'' if b is None else b}"
            for v in (a, b):
                if isinstance(v, XErr):
                    raise v
            if a is None:
                a = "" if isinstance(b, str) else 0.0
            if b is None:
                b = "" if isinstance(a, str) else 0.0
            if isinstance(a, str) != isinstance(b, str):
                return op in ("<>",)
            if not isinstance(a, str):
                a, b = self.num(a), self.num(b)
            return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]
        if t == "fn":
            f, args = n[1], n[2]
            if f == "IF":
                c = self.ev(args[0])
                if isinstance(c, str):
                    raise XErr("#VALUE!")
                return self.ev(args[1]) if self.num(c) else (self.ev(args[2]) if len(args) > 2 else False)
            if f == "IFERROR":
                try:
                    return self.ev(args[0])
                except XErr:
                    return self.ev(args[1])
            if f == "NA":
                raise XErr("#N/A")
            if f in ("COUNT", "SUM", "MIN", "MAX", "AVERAGE"):
                vals = []
                for a in args:
                    if a[0] == "ref":
                        for v in self.rng(a):
                            if isinstance(v, XErr):
                                if f != "COUNT":
                                    raise v
                                continue
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                vals.append(float(v))
                    else:
                        try:
                            vals.append(self.num(self.ev(a)))
                        except XErr:
                            if f != "COUNT":
                                raise
                if f == "COUNT":
                    return float(len(vals))
                if f == "SUM":
                    return sum(vals)
                if not vals:
                    if f == "AVERAGE":
                        raise XErr("#DIV/0!")
                    return 0.0
                return {"MIN": min, "MAX": max}.get(f, lambda x: sum(x) / len(x))(vals)
            if f == "N":
                try:
                    v = self.ev(args[0])
                except XErr:
                    raise
                return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            if f == "ABS":
                return abs(self.num(self.ev(args[0])))
            if f in ("AND", "OR"):
                vals = [bool(self.num(self.ev(a))) for a in args]
                return all(vals) if f == "AND" else any(vals)
            if f == "YEAR":
                v = self.ev(args[0])
                if isinstance(v, (dt.date, dt.datetime)):
                    return float(v.year)
                if v is None or v == "":
                    raise XErr("#VALUE!")
                return float((dt.date(1899, 12, 30) + dt.timedelta(days=int(self.num(v)))).year)
            if f == "MATCH":
                x = self.ev(args[0])
                lst = self.rng(args[1])
                for i, v in enumerate(lst):
                    if not isinstance(v, XErr) and v is not None and (
                            v == x or (isinstance(v, (int, float)) and isinstance(x, (int, float)) and float(v) == float(x))):
                        return float(i + 1)
                raise XErr("#N/A")
            if f == "INDEX":
                lst = self.rng(args[0])
                k = int(self.num(self.ev(args[1])))
                if not 1 <= k <= len(lst):
                    raise XErr("#REF!")
                v = lst[k - 1]
                if isinstance(v, XErr):
                    raise v
                return v
            raise ValueError(f"未支持的函数 {f}")
        raise ValueError(n)


# ---------------------------------------------------------------- 验收
def run_checks(xlsx, ex, ref_mcap=None, expected_charts=13):
    wb = openpyxl.load_workbook(xlsx)
    ev = Evaluator(wb)
    res = []

    def add(name, ok, detail):
        res.append((name, ok, detail))

    def val(sheet, coord):
        try:
            return ev.cell(sheet, coord)
        except XErr as e:
            return e

    # 定位行列
    t3 = wb["年度三表"]
    years, ycol = [], {}
    for c in t3[4][4:]:
        if isinstance(c.value, int):
            years.append(c.value)
            ycol[c.value] = c.column_letter
    trow = {}
    for r in range(5, t3.max_row + 1):
        if str(t3[f"A{r}"].value or "").startswith("核心"):
            trow[t3[f"C{r}"].value] = r
    get = lambda name, y: (lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None)(
        val("年度三表", f"{ycol[y]}{trow[name]}")) if name in trow else None

    # 1 三表勾稽
    bad_gp, bad_bs, n_gp, n_bs = [], [], 0, 0
    for y in years:
        rev, cogs, gp = get("营收", y), get("营业成本", y), get("毛利", y)
        if None not in (rev, cogs, gp) and rev:
            n_gp += 1
            if abs(gp - (rev - cogs)) / abs(rev) > 0.005:
                bad_gp.append(y)
        ta, tl, eq = get("总资产", y), get("总负债", y), get("股东权益合计（含少数股东）", y)
        if None not in (ta, tl, eq) and ta:
            n_bs += 1
            if abs(ta - (tl + eq)) / abs(ta) > 0.02:
                bad_bs.append(y)
    add("1 年度三表勾稽（毛利 = 营收 − 成本 ±0.5%；资产 = 负债 + 权益 ±2%）", not bad_gp and not bad_bs,
        f"毛利检查 {n_gp} 年、不通过 {bad_gp or '无'}；资产平衡检查 {n_bs} 年、不通过 {bad_bs or '无'}")

    # 2 衍生指标与三表一致：逐格求公式，并与长表直接计算的独立值比较
    dv = wb["年度衍生指标"]
    dh = {c.value: c.column_letter for c in dv[4] if c.value}
    drow = {dv[f"A{r}"].value: r for r in range(5, dv.max_row + 1) if isinstance(dv[f"A{r}"].value, int)}
    errs, nform, ncmp, mism = [], 0, 0, []
    for r in drow.values():
        for h, col in dh.items():
            v = dv[f"{col}{r}"].value
            if isinstance(v, str) and v.startswith("="):
                nform += 1
                x = val("年度衍生指标", f"{col}{r}")
                if isinstance(x, XErr):
                    errs.append(f"{col}{r}{x}")
    U = next((h[3:-1] for h in dh if h.startswith("营收（")), "亿美元")

    def indep(y, h):
        per = ex.aper.get(y)
        c = lambda k: ex.core(k, per)[0]
        try:
            if h == "毛利率":
                return c("gp") / c("rev")
            if h == "营业利润率":
                return c("op") / c("rev")
            if h == "净利率":
                return c("ni") / c("rev")
            if h == f"FCF（{U}）":
                return (c("ocf") + c("capex")) / 1e8
            if h == "基本加权股数（百万股）":
                return ex.shares(per, "b")[0]
            if h == "ROE":
                return c("ni") / c("eq")
            if h == f"营收（{U}）":
                return c("rev") / 1e8
        except (TypeError, ZeroDivisionError):
            return None
        return None

    for y, r in drow.items():
        for h in ["毛利率", "营业利润率", "净利率", f"FCF（{U}）", "基本加权股数（百万股）", "ROE", f"营收（{U}）"]:
            a = indep(y, h)
            if h not in dh:
                continue
            b = val("年度衍生指标", f"{dh[h]}{r}")
            b = b if isinstance(b, (int, float)) and not isinstance(b, bool) else None
            if a is None and b is None:
                continue
            ncmp += 1
            if a is None or b is None or abs(a - b) > 1e-9 * max(1, abs(a)):
                mism.append(f"{y}{h}")
    add("2 年度衍生指标与年度三表一致", not errs and not mism and nform > 0,
        f"公式 {nform} 格全部可求值" + (f"，求值出错 {errs[:5]}" if errs else "") +
        f"；与长表独立计算比对 {ncmp} 个值，不一致 {mism[:8] or '无'}")

    # 3 近12季度最新一季 = 导出最新报告期，说明写明期末日
    qs = wb["近12季度明细"]
    labels = [c.value for c in qs[3][2:] if c.value and str(c.value).startswith("FY")]
    latest = ex.label(ex.quarters[-1]) if ex.quarters else None
    end = ex.meta[ex.quarters[-1]]["end"] if ex.quarters else None
    notes_txt = " ".join(str(c.value) for row in wb["说明"].iter_rows() for c in row if c.value)
    ok3 = bool(labels) and labels[-1] == latest and len(labels) == min(12, len(ex.quarters)) and end in notes_txt
    add("3 近12季度明细最新一季 = 导出最新报告期，说明写明期末日", ok3,
        f"工作簿最新一季 {labels[-1] if labels else '无'}，导出最新 {latest}（期末日 {end}）；共 {len(labels)} 季；说明{'已' if end and end in notes_txt else '未'}写期末日")

    # 4 现代分类勾稽差额全部填写
    rv = wb["历年营收与业务拆分"]
    hc = {c.value: c.column_letter for c in rv[3] if c.value}
    dcol = next((v for k, v in hc.items() if str(k).startswith("现代分类勾稽差额")), None)
    empty, nrow = [], 0
    for r in range(4, rv.max_row + 1):
        if rv[f"A{r}"].value is None:
            continue
        nrow += 1
        x = val("历年营收与业务拆分", f"{dcol}{r}") if dcol else None
        if not isinstance(x, (int, float)) or isinstance(x, bool):
            empty.append(rv[f"A{r}"].value)
    add("4 历年营收与业务拆分的现代分类勾稽差额全部填写", bool(dcol) and not empty and nrow > 0,
        f"{nrow} 行，未填 {empty or '无'}")

    # 5 股数年变化 0.5–2，超出的已标注
    col_ratio, col_note = dh.get("股数年变化（倍）"), dh.get("备注")
    out, unflag, nratio = [], [], 0
    for y, r in drow.items():
        x = val("年度衍生指标", f"{col_ratio}{r}")
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            nratio += 1
            if not 0.5 <= x <= 2:
                out.append(f"{y}:{x:.2f}")
                if "拆股" not in str(dv[f"{col_note}{r}"].value or ""):
                    unflag.append(y)
    add("5 股数年变化比值在 0.5–2，超出的标注拆股 / 合股", not unflag,
        f"检查 {nratio} 对相邻财年；超出 {out or '无'}" + (f"；未标注 {unflag}" if unflag else ("；已全部标注，待人工核对日线复权" if out else "")))

    # 6 市值
    vs = wb["市值与估值"]
    vh = {c.value: c.column_letter for c in vs[4] if c.value}
    bad_px, bad_m, far, nm, nref = [], [], [], 0, 0
    basis, no_basis = defaultdict(int), []
    for r in range(5, vs.max_row + 1):
        y = vs[f"A{r}"].value
        if not isinstance(y, int):
            continue
        p, d, e = vs[f"{vh['收盘价（美元）']}{r}"].value, vs[f"{vh['年末交易日']}{r}"].value, vs[f"{vh['财年期末日']}{r}"].value
        m = val("市值与估值", f"{vh['市值（亿美元）']}{r}")
        if not isinstance(m, (int, float)) or isinstance(m, bool):
            continue
        nm += 1
        if "股数口径" in vh:
            b = str(vs[f"{vh['股数口径']}{r}"].value or "")
            if b.startswith("期末股数"):
                basis["期末股数"] += 1
            elif b.startswith("基本加权股数"):
                basis["基本加权股数"] += 1
            else:
                no_basis.append(y)
        s = val("市值与估值", f"{vh.get('股数（百万股）') or vh['基本加权股数（百万股）']}{r}")
        if abs(m - p * s / 100) > 1e-9 * abs(m):
            bad_m.append(y)
        if (e - d).days > 10 or d > e:
            bad_px.append(y)
        if ref_mcap and y in ref_mcap and ref_mcap[y]:
            nref += 1
            diff = m / ref_mcap[y] - 1
            if abs(diff) >= 0.05:
                far.append(f"{y}:{diff:+.1%}")
    if ref_mcap:
        ok6 = not bad_px and not bad_m and not far
        ext = f"与数据提供方市值对照 {nref} 年，差异 ≥5% 的 {far or '无'}"
    else:
        ok6 = not bad_px and not bad_m
        ext = "没有数据提供方市值（未传 --ref-mcap），外部 5% 对照未做，已记入缺口"
    ok6 = ok6 and not no_basis
    btxt = ("股数口径：" + "、".join(f"{k} {v} 年" for k, v in basis.items()) +
            (f"，未注明 {no_basis}" if no_basis else "")) if "股数口径" in vh else "股数口径列缺（旧版工作簿）"
    add("6 市值与估值：年末价 × 股数（与数据提供方市值差 < 5%）", ok6,
        f"{nm} 年有市值；{btxt}；年末交易日距财年末 > 10 天的 {bad_px or '无'}；市值 ≠ 价 × 股数 的 {bad_m or '无'}；{ext}")

    # 7 图表
    z = zipfile.ZipFile(xlsx)
    parts = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    bad_ref, bad_cell, nref_c = [], [], 0
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    calc = wb["图表计算"]
    for p in parts:
        root = ET.fromstring(z.read(p))
        for f in root.iter("{%s}f" % ns["c"]):
            nref_c += 1
            txt = f.text or ""
            if not txt.startswith("'图表计算'!"):
                bad_ref.append(txt)
                continue
            rng = txt.split("!", 1)[1].replace("$", "")
            a, b = rng.split(":") if ":" in rng else (rng, rng)
            for row in calc[a:b]:
                for c in row:
                    if c.row >= 5 and not (isinstance(c.value, str) and c.value.startswith("=")):
                        bad_cell.append(c.coordinate)
    ov_charts = len(wb["图表总览"]._charts)
    calc_err = []
    for row in calc.iter_rows(min_row=5, max_row=min(calc.max_row, 70)):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                x = val("图表计算", c.coordinate)
                if isinstance(x, XErr) and str(x) != "#N/A":
                    calc_err.append(f"{c.coordinate}{x}")
    lo = lo_open_test(xlsx)
    ok7 = (len(parts) == expected_charts == ov_charts and not bad_ref and not bad_cell and not calc_err and lo[0] is not False)
    add("7 图表总览的全部图表能打开，数据区引用'图表计算'的公式", ok7,
        f"图表 {len(parts)} 张（应为 {expected_charts}，图表总览上 {ov_charts}）；数据引用 {nref_c} 处，"
        f"不指向图表计算的 {bad_ref[:3] or '无'}，指向非公式单元格的 {bad_cell[:3] or '无'}；"
        f"图表计算公式求值错误 {calc_err[:3] or '无'}；{lo[1]}")

    # 8 说明写清每张表
    sm = wb["说明"]
    listed = {sm[f"A{r}"].value: (sm[f"B{r}"].value, sm[f"C{r}"].value) for r in range(1, sm.max_row + 1)}
    missing = [s for s in wb.sheetnames if not (s in listed and listed[s][0] and listed[s][1])]
    add("8 说明写清每张表的口径与缺口", not missing, f"{len(wb.sheetnames)} 张表，缺说明 {missing or '无'}")
    return res


def lo_open_test(xlsx):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None, "未装 LibreOffice，打开测试改为 XML 解析（已通过）"
    with tempfile.TemporaryDirectory() as td:
        try:
            subprocess.run([exe, "--headless", f"-env:UserInstallation=file://{td}/u", "--convert-to", "pdf",
                            "--outdir", td, xlsx], capture_output=True, timeout=240)
        except (subprocess.TimeoutExpired, OSError) as e:
            return None, f"LibreOffice 打开超时 / 失败（{e.__class__.__name__}），改为 XML 解析（已通过）"
        pdf = os.path.join(td, os.path.splitext(os.path.basename(xlsx))[0] + ".pdf")
        if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
            return True, "LibreOffice 无界面打开并导出 PDF 成功"
        return False, "LibreOffice 打开失败"


def print_checks(res, code):
    print(f"== {code} 验收 ==")
    for name, ok, detail in res:
        print(f"[{'通过' if ok else '不通过'}] {name}\n        {detail}")
    n_ok = sum(1 for _, ok, _ in res if ok)
    print(f"结果：{n_ok}/{len(res)} 通过")
    return n_ok == len(res)


# ---------------------------------------------------------------- PNG 快照（--png）
# 编号与"图表总览"的顺序一致；文件名 图01_营收与增速.png …
PNG_NAMES = {1: "营收与增速", 2: "分部结构", 3: "利润率带", 4: "现金流去向", 5: "资本开支与经营现金流",
             6: "股数与回购", 7: "市值与营收", 8: "估值倍数带", 9: "近12季营收与增速", 10: "近12季利润率",
             11: "股价与事件", 12: "市值对比参考公司", 13: "同行横向比较"}
# 低饱和、色盲友好的固定配色（Tableau 10 前段），不用彩虹色
PNG_COLORS = ["#4E79A7", "#F28E2B", "#59A14F", "#E15759", "#76B7B2", "#B07AA1", "#9C755F", "#BAB0AC"]
LAYER2_PREFIX = ("分部", "参考公司")  # 第二层系列：这些全空时整张图视为无数据
PNG_RE = re.compile(r"^图\d{2}_.+\.png$")


def _title_text(t):
    if t is None:
        return ""
    if isinstance(t, str):
        return t
    try:
        return "".join(r.t for p in t.tx.rich.p for r in (p.r or []))
    except AttributeError:
        return ""


def _ref_values(ev, f):
    sh, rng = f.rsplit("!", 1)
    sh = sh.strip("'").replace("''", "'")
    rng = rng.replace("$", "")
    a, b = rng.split(":") if ":" in rng else (rng, rng)
    out = []
    for row in ev.wb[sh][a:b]:
        for c in row:
            try:
                out.append(ev.cell(sh, c.coordinate))
            except XErr:
                out.append(None)
    return out


def _num_or_none(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)


def chart_specs(wb):
    """读"图表总览"上的每张原生图表，按其数据引用从"图表计算"求值，返回可画图的结构。"""
    ev = Evaluator(wb)
    specs = []
    for k, ch in enumerate(wb["图表总览"]._charts, start=1):
        t = _title_text(ch.title)
        m = re.match(r"\s*(\d+)\s*(.*)", t)
        num, label = (int(m.group(1)), m.group(2)) if m else (k, t)
        groups = []
        for gi, sub in enumerate(getattr(ch, "_charts", None) or [ch]):
            kind = sub.tagname
            series = []
            for s in sub.series:
                name = ""
                if s.tx is not None and s.tx.strRef is not None:
                    v = _ref_values(ev, s.tx.strRef.f)
                    name = str(v[0]) if v and v[0] is not None else ""
                elif s.tx is not None and s.tx.v:
                    name = s.tx.v
                if kind == "scatterChart":
                    xs = _ref_values(ev, s.xVal.numRef.f)
                    ys = [_num_or_none(v) for v in _ref_values(ev, s.yVal.numRef.f)]
                    # 线条设为 noFill 的是只画点的系列（股价事件），其余画成线（月末收盘）
                    ln = s.spPr.ln if s.spPr is not None else None
                    as_line = not (ln is not None and ln.noFill)
                    series.append({"name": name, "x": xs, "y": ys, "line": as_line})
                else:
                    cref = (s.cat.numRef or s.cat.strRef) if s.cat is not None else None
                    cats = _ref_values(ev, cref.f) if cref is not None else []
                    ys = [_num_or_none(v) for v in _ref_values(ev, s.val.numRef.f)]
                    series.append({"name": name, "cat": cats, "y": ys})
            yfmt = getattr(sub.y_axis, "number_format", None) or ""
            log = getattr(getattr(sub.y_axis, "scaling", None), "logBase", None)
            groups.append({"kind": kind, "stacked": getattr(sub, "grouping", "") == "stacked",
                           "secondary": gi > 0, "pct": "%" in str(yfmt), "log": bool(log),
                           "ytitle": _title_text(getattr(sub.y_axis, "title", None)), "series": series})
        specs.append({"num": num, "title": label, "groups": groups})
    return specs


def spec_has_data(spec):
    ser = [s for g in spec["groups"] for s in g["series"]]
    filled = lambda s: any(v is not None for v in s["y"])
    l2 = [s for s in ser if s["name"].startswith(LAYER2_PREFIX)]
    if l2 and not any(filled(s) for s in l2):
        return False
    if "同行" in spec["title"]:  # 第一行是本公司（自动），至少还要有一家同行才算第二层已填
        return any(sum(v is not None for v in s["y"]) >= 2 for s in ser)
    return any(filled(s) for s in ser)


def png_name(num):
    return f"图{num:02d}_{PNG_NAMES.get(num, f'图表{num}')}.png"


def _pick_cjk_font(fm):
    names = {f.name for f in fm.fontManager.ttflist}
    for want in ("Noto Sans CJK SC", "Noto Sans SC"):
        if want in names:
            return want
    for n in sorted(names):
        if n.startswith("Noto Sans CJK") and not n.endswith(("Black", "Thin", "Light", "DemiLight", "Medium")):
            return n
    for n in sorted(names):
        if n.startswith("WenQuanYi"):
            return n
    return None


def export_png(xlsx, outdir, code=""):
    """把"图表总览"的每张图按同一份数据区画成 PNG。缺 matplotlib 时返回 None（不报错）。"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager as fm
        from matplotlib.ticker import FuncFormatter, PercentFormatter
    except ImportError:
        print("未装 matplotlib，跳过 PNG 导出（python3 -m pip install --user matplotlib）")
        return None
    font = _pick_cjk_font(fm)
    if font:
        plt.rcParams["font.family"] = [font, "DejaVu Sans"]
    else:
        plt.rcParams["font.family"] = ["DejaVu Sans"]
        print("警告：没找到 Noto Sans CJK / Noto Sans SC / WenQuanYi 中文字体，退回 DejaVu Sans，图中中文可能显示为方框")
    plt.rcParams["axes.unicode_minus"] = False
    os.makedirs(outdir, exist_ok=True)
    wb = openpyxl.load_workbook(xlsx)
    written = []
    for sp in chart_specs(wb):
        if not spec_has_data(sp):
            print(f"跳过：图 {sp['num']} 无数据")
            continue
        fig, ax = plt.subplots(figsize=(8, 4), dpi=200)
        ax2, ci, cats_main = None, 0, []
        handles = []
        for g in sp["groups"]:
            a = ax
            if g["secondary"]:
                ax2 = ax2 or ax.twinx()
                a = ax2
            if g["kind"] == "scatterChart":
                for s in g["series"]:
                    pts = [(x, y) for x, y in zip(s["x"], s["y"])
                           if y is not None and isinstance(x, (dt.date, dt.datetime))]
                    if not pts:
                        continue
                    xs, ys = zip(*pts)
                    c = PNG_COLORS[ci % len(PNG_COLORS)]
                    ci += 1
                    if s["line"]:
                        h, = a.plot(xs, ys, color=c, linewidth=1.2, label=s["name"])
                    else:
                        h = a.scatter(xs, ys, color=c, s=14, zorder=3, label=s["name"])
                    handles.append(h)
                if g["log"]:
                    a.set_yscale("log")
            else:
                # 组合图的副轴系列没有单独的分类轴，沿用主图的分类
                cats = (g["series"][0]["cat"] if g["series"] else []) or cats_main
                cats_main = cats_main or cats
                keep = [i for i, v in enumerate(cats) if v not in (None, "")]
                lab = [str(int(cats[i])) if isinstance(cats[i], float) and cats[i].is_integer() else str(cats[i])
                       for i in keep]
                xs = list(range(len(keep)))
                nb = sum(1 for s in g["series"] if any(v is not None for v in s["y"]))
                jb = 0
                pos_base, neg_base = [0.0] * len(keep), [0.0] * len(keep)
                for j, s in enumerate(g["series"]):
                    ys = [s["y"][i] if i < len(s["y"]) else None for i in keep]
                    c = PNG_COLORS[ci % len(PNG_COLORS)]
                    ci += 1
                    if not any(v is not None for v in ys):
                        continue
                    if g["kind"] == "barChart":
                        vals = [v if v is not None else 0.0 for v in ys]
                        if g["stacked"]:
                            bottom = [pb if v >= 0 else nb_ for v, pb, nb_ in zip(vals, pos_base, neg_base)]
                            h = a.bar(xs, vals, 0.7, bottom=bottom, color=c, label=s["name"])
                            for i, v in enumerate(vals):
                                if v >= 0:
                                    pos_base[i] += v
                                else:
                                    neg_base[i] += v
                        else:
                            w = 0.8 / max(nb, 1)
                            h = a.bar([x - 0.4 + w * (jb + 0.5) for x in xs], vals, w, color=c, label=s["name"])
                            jb += 1
                    else:
                        h, = a.plot(xs, [v if v is not None else float("nan") for v in ys], color=c,
                                    linewidth=1.6, marker="o", markersize=2.5, label=s["name"])
                    handles.append(h)
                if not g["secondary"]:
                    a.set_xticks(xs)
                    a.set_xticklabels(lab, rotation=45 if len(lab) > 12 else 0,
                                      ha="right" if len(lab) > 12 else "center", fontsize=7)
            if g["pct"]:
                a.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=None))
            else:
                plain = FuncFormatter(lambda v, _: f"{v:,.0f}" if abs(v - round(v)) < 1e-9 else f"{v:,.1f}")
                a.yaxis.set_major_formatter(plain)
                if g["log"]:
                    a.yaxis.set_minor_formatter(plain)
            if g["ytitle"]:
                a.set_ylabel(g["ytitle"], fontsize=8)
            a.tick_params(which="both", labelsize=7)
        if any(g["kind"] == "barChart" for g in sp["groups"]):
            ax.axhline(0, color="#999999", linewidth=0.6)
        ax.set_title(f"图 {sp['num']}  {sp['title']}" + (f"｜{code}" if code else ""), fontsize=10, loc="left")
        for a in (ax, ax2):
            if a is not None:
                for side in ("top",) + (() if a is ax2 else ("right",)):
                    a.spines[side].set_visible(False)
        ax.grid(axis="y", color="#E5E5E5", linewidth=0.6)
        ax.set_axisbelow(True)
        if handles:
            ax.legend(handles=handles, loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=min(len(handles), 4),
                      fontsize=7, frameon=False)
        fig.tight_layout()
        p = os.path.join(outdir, png_name(sp["num"]))
        fig.savefig(p)
        plt.close(fig)
        written.append(p)
    print(f"PNG 快照：{len(written)} 张，写入 {outdir}")
    return written


def check_png(xlsx, outdir):
    wb = openpyxl.load_workbook(xlsx)
    want = sorted(png_name(sp["num"]) for sp in chart_specs(wb) if spec_has_data(sp))
    have = sorted(f for f in (os.listdir(outdir) if os.path.isdir(outdir) else []) if PNG_RE.match(f))
    ok = len(have) == len(want) and set(have) == set(want)
    return ("9 PNG 快照数量 = 有数据的图数量", ok,
            f"有数据的图 {len(want)} 张，{outdir} 里 PNG {len(have)} 张；缺 {sorted(set(want) - set(have)) or '无'}，"
            f"多出 {sorted(set(have) - set(want)) or '无'}")


def read_ref_mcap(path):
    out = {"_src": path}
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            y, v = fnum(r.get("年度")), fnum(r.get("市值") or r.get("市值（亿美元）"))
            if y and v:
                out[int(y)] = v
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="从富途导出生成历史数据 Excel（层级 1）")
    ap.add_argument("code")
    ap.add_argument("--root", required=True, help="futu_export 根目录")
    ap.add_argument("--out", help="输出目录")
    ap.add_argument("--ai", default="Claude", help="文件名里的 AI 名")
    ap.add_argument("--ver", type=int, default=1, help="文件名里的版本序号")
    ap.add_argument("--from-year", type=int, help="只用这一年及以后的财年")
    ap.add_argument("--ref-mcap", help="数据提供方市值 CSV（列：年度,市值；亿美元）")
    ap.add_argument("--check", action="store_true", help="生成后按规格验收")
    ap.add_argument("--check-only", help="只验收已有的 xlsx，不重新生成")
    ap.add_argument("--checks-csv", help="把验收结果追加写入这个 CSV（汇总多家用）")
    ap.add_argument("--png", metavar="目录", help="把图表总览的每张图导出为 PNG 快照到这个目录（每家公司一个目录）")
    a = ap.parse_args(argv)
    code = a.code.upper()
    sp = SPECIAL.get(code, {})
    ex = Export(code, a.root, a.from_year or sp.get("from_year"))
    if not ex.years:
        sys.exit(f"{code}：长表里没有年报（报告类型 7）")
    ref = read_ref_mcap(a.ref_mcap) if a.ref_mcap else None
    if a.check_only:
        path, nchart = a.check_only, 13
    else:
        if not a.out:
            sys.exit("缺 --out")
        px, info = load_daily(code, a.root, sp.get("price_start"))
        os.makedirs(a.out, exist_ok=True)
        name = f"{dt.date.today():%Y%m%d}_{code}_历史数据_v{a.ver}_{a.ai}.xlsx"
        b = Builder(ex, px, info, a.ai, ref, sp)
        path = b.build(os.path.join(a.out, name))
        nchart = b.expected_charts
        print(f"已生成 {path}（{len(ex.years)} 个财年，{len(b.events)} 个股价事件，{nchart} 张图）")
    png_dir = None
    if a.png:
        if export_png(path, a.png, code) is not None:
            png_dir = a.png
        else:
            print("PNG 未导出，--check 不做第 9 项")
    if a.check or a.check_only:
        res = run_checks(path, ex, ref, nchart)
        if png_dir:
            res.append(check_png(path, png_dir))
        ok = print_checks(res, code)
        if a.checks_csv:
            new = not os.path.exists(a.checks_csv)
            with open(a.checks_csv, "a", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(["代码", "文件", "检查项", "结果", "说明"])
                for n, o, d in res:
                    w.writerow([code, os.path.basename(path), n, "通过" if o else "不通过", d])
        return 0 if ok else 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
