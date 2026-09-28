# -*- coding: utf-8 -*-
"""
退款单数据分析 · 计算模块（可复用）
================================
本模块负责：读取电商退款单明细(xlsx) → 按口径核算 → 输出结构化分析结果。
适用于淘宝/天猫「退款明细」导出的标准字段，也支持按 COLUMN_MAP 适配其它平台。

使用方式：
    import refund_analysis as A
    # 先修改本文件顶部 CONFIG，再：
    A.run()            # 执行核算并打印 CHECK 摘要
    # 或直接作为脚本运行： python refund_analysis.py

输出（模块级变量，供 refund_render.py 复用）：
    succ / closed        退款成功 / 退款关闭 行
    b_count / b_amt      四大板块 单数 / 金额
    ship_count/ ship_amt 未发货 / 发货后 两桶
    total_refund 等
"""
import openpyxl, math, datetime
from collections import defaultdict, Counter

# ============================ CONFIG ============================
# 修改这里即可适配不同文件 / 店铺 / 周期
PATH = r"D:\Chill\训练数据源\喵住萤石6.27-9.27退款单.xlsx"   # 输入文件
SHEET_NAME = "export"            # 数据所在 Sheet 名
TOTAL_SALES = 5_550_000.0        # 分母：店铺总销售额（元）
START_DATE = "2026-06-27"        # 周期起点（用于按周切片）；格式 YYYY-MM-DD
# 广义品退命中关键词（出现在"买家退款原因"即算）
PIN_KEYS = ["质量", "瑕疵", "破损", "污渍", "描述不符", "材质",
            "少件", "漏发", "发错货", "缺货", "做工"]

# 逻辑字段 → 实际表头（不同平台/导出版本可能不同，按需改右侧）
COLUMN_MAP = {
    "退款状态": "退款状态",
    "货物状态": "货物状态",
    "售后类型": "售后类型",
    "退款总额": "退款总额",
    "退给买家金额": "退给买家金额",
    "买家退款原因": "买家退款原因",
    "商品id": "商品id",
    "宝贝标题": "宝贝标题",
    "订单编号": "订单编号",
    "退款编号": "退款编号",
    "订单付款时间": "订单付款时间",
    "退款完结时间": "退款完结时间",
    "部分退款/全额退款": "部分退款/全额退款",
    "责任方": "责任方",
}
# ================================================================

def _load():
    wb = openpyxl.load_workbook(PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    # 兼容表头别名：用 COLUMN_MAP 把逻辑名映射到实际列索引
    idx = {}
    for logical, actual in COLUMN_MAP.items():
        if actual in header:
            idx[logical] = header.index(actual)
        else:
            # 退化：尝试直接按逻辑名找
            idx[logical] = header.index(logical) if logical in header else None
    return rows[1:], idx

DATA, IDX = _load()

def fnum(v):
    try:
        return float(v)
    except Exception:
        return 0.0

def _col(name):
    return IDX[name]

def bucket(r):
    """四大板块分类：货物状态 × 售后类型 交叉判定。"""
    gs = r[_col("货物状态")]; st = r[_col("售后类型")]
    if gs == "未发货":
        return "未发货退款"
    if gs == "未收到货" and st == "仅退款":
        return "在途仅退款"
    if gs == "已收到货" and st == "仅退款":
        return "已发仅退款"
    if st == "退货退款":
        return "退货退款"
    return "其他"

def ship_bucket(r):
    """首要对比：未发货 vs 发货后。换货/补寄/维修 → 排除。"""
    b = bucket(r)
    if b == "未发货退款":
        return "未发货"
    if b in ("在途仅退款", "已发仅退款", "退货退款"):
        return "发货后"
    return "排除"

def is_pintui(r):
    """广义品退判定。"""
    reason = r[_col("买家退款原因")] or ""
    return any(k in str(reason) for k in PIN_KEYS)

def pdate(v):
    if not v:
        return None
    s = str(v)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s[:19], fmt)
        except Exception:
            pass
    return None

START = datetime.datetime.strptime(START_DATE, "%Y-%m-%d")
def week_idx(d):
    return (d.date() - START.date()).days // 7

MAIN = ["未发货退款", "在途仅退款", "已发仅退款", "退货退款"]
COLORS = {
    "未发货退款": "#4C72B0", "在途仅退款": "#DD8452", "已发仅退款": "#55A868",
    "退货退款": "#C44E52", "其他": "#8172B3", "未发货": "#4C72B0", "发货后": "#C44E52",
}

def run():
    succ = [r for r in DATA if r[_col("退款状态")] == "退款成功"]
    closed = [r for r in DATA if r[_col("退款状态")] == "退款关闭"]

    b_count = Counter(); b_amt = defaultdict(float)
    ship_count = Counter(); ship_amt = defaultdict(float)
    ALLB = MAIN + ["其他"]
    reason_c = {b: Counter() for b in ALLB}
    fp_c = {b: Counter() for b in ALLB}
    big_n = {b: 0 for b in MAIN}; big_a = {b: 0.0 for b in MAIN}
    week_amt = {b: defaultdict(float) for b in ALLB}
    week_ship = {"未发货": defaultdict(float), "发货后": defaultdict(float)}
    month_amt = {b: defaultdict(float) for b in ALLB}
    prod = defaultdict(lambda: {"n": 0, "amt": 0.0, "pn": 0, "pa": 0.0, "reasons": Counter()})
    pintui_n = 0; pintui_a = 0.0
    pintui_prod = defaultdict(lambda: {"n": 0, "amt": 0.0})
    pintui_week = defaultdict(float)
    pintui_resp = Counter(); pintui_reason = Counter()
    big_rows = []; pintui_rows = []
    all_wi = []

    for r in succ:
        b = bucket(r); amt = fnum(r[_col("退款总额")]); d = pdate(r[_col("退款完结时间")])
        reason = r[_col("买家退款原因")]
        b_count[b] += 1; b_amt[b] += amt
        reason_c[b][reason] += 1
        fp_c[b][r[_col("部分退款/全额退款")]] += 1
        if amt > 1000:
            big_n[b] += 1; big_a[b] += amt; big_rows.append(r)
        sb = ship_bucket(r)
        if sb in ("未发货", "发货后"):
            ship_count[sb] += 1; ship_amt[sb] += amt
        if d:
            wi = week_idx(d); all_wi.append(wi)
            if wi >= 0:
                week_amt[b][wi] += amt
                if sb in week_ship:
                    week_ship[sb][wi] += amt
                month_amt[b][d.strftime("%Y-%m")] += amt
        pid = r[_col("商品id")]; p = prod[pid]
        p["n"] += 1; p["amt"] += amt; p["reasons"][reason] += 1
        if is_pintui(r):
            p["pn"] += 1; p["pa"] += amt
            pintui_n += 1; pintui_a += amt
            pintui_prod[pid]["n"] += 1; pintui_prod[pid]["amt"] += amt
            if d and d >= START:
                wi = week_idx(d)
                if wi >= 0:
                    pintui_week[wi] += amt
            pintui_resp[r[_col("责任方")] if r[_col("责任方")] else "未标记"] += 1
            pintui_reason[reason] += 1
            pintui_rows.append(r)

    WEEKS = (max(all_wi) + 1) if all_wi else 0
    week_labels = []
    for i in range(WEEKS):
        s = START + datetime.timedelta(days=7 * i)
        e = s + datetime.timedelta(days=6)
        week_labels.append("W" + str(i + 1) + " " + s.strftime('%m-%d') + "~" + e.strftime('%m-%d'))
    week_amt_arr = {b: [week_amt[b].get(i, 0.0) for i in range(WEEKS)] for b in MAIN}
    week_ship_arr = {k: [week_ship[k].get(i, 0.0) for i in range(WEEKS)] for k in week_ship}
    pintui_week_arr = [pintui_week.get(i, 0.0) for i in range(WEEKS)]
    months = sorted({d.strftime("%Y-%m") for d in [pdate(r[_col("退款完结时间")]) for r in succ] if d})
    month_amt_arr = {b: [month_amt[b].get(m, 0.0) for m in months] for b in MAIN}
    top_prod = sorted(prod.items(), key=lambda x: -x[1]["amt"])[:20]
    top_pintui_prod = sorted(pintui_prod.items(), key=lambda x: -x[1]["amt"])[:15]
    total_refund = sum(b_amt.values())
    unship_rate = ship_amt["未发货"] / TOTAL_SALES
    ship_rate = ship_amt["发货后"] / TOTAL_SALES
    overall_rate = total_refund / TOTAL_SALES

    print("=== CHECK ===")
    print("succ", len(succ), "closed", len(closed))
    print("total_refund", round(total_refund), "overall_rate", round(overall_rate * 100, 2))
    print("未发货", ship_count["未发货"], round(ship_amt["未发货"]), round(unship_rate * 100, 2))
    print("发货后", ship_count["发货后"], round(ship_amt["发货后"]), round(ship_rate * 100, 2))
    print("WEEKS", WEEKS, "months", months)
    print(">1000 n", len(big_rows), "amt", round(sum(fnum(r[_col('退款总额')]) for r in big_rows)))
    print("pintui n", pintui_n, "amt", round(pintui_a))
    return dict(
        succ=succ, closed=closed, b_count=b_count, b_amt=b_amt,
        ship_count=ship_count, ship_amt=ship_amt, reason_c=reason_c, fp_c=fp_c,
        big_n=big_n, big_a=big_a, week_labels=week_labels, week_amt_arr=week_amt_arr,
        week_ship_arr=week_ship_arr, pintui_week_arr=pintui_week_arr, months=months,
        month_amt_arr=month_amt_arr, top_prod=top_prod, top_pintui_prod=top_pintui_prod,
        total_refund=total_refund, unship_rate=unship_rate, ship_rate=ship_rate,
        overall_rate=overall_rate, pintui_n=pintui_n, pintui_a=pintui_a,
        pintui_resp=pintui_resp, pintui_reason=pintui_reason, big_rows=big_rows,
        pintui_rows=pintui_rows, prod=prod, pintui_prod=pintui_prod, WEEKS=WEEKS,
        DATA=DATA,
    )

if __name__ == "__main__":
    run()
