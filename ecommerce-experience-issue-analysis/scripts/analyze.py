#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
体验问题订单分析引擎（零依赖，仅标准库）

输入：平台导出的体验问题订单表（.csv / .json / .xlsx）
输出：metrics.json —— 供 build_report.py 渲染 HTML 报告

用法：
    python analyze.py --input orders.csv --output metrics.json
    python analyze.py --input orders.json --sku-keys 内存容量,颜色分类
    python analyze.py --input orders.xlsx --rules my_rules.json

设计要点：
1. 字段自动映射：兼容淘宝/天猫/京东/拼多多/抖音等平台的表头命名差异。
2. 旅程六阶段归因：按「消费者旅程」而非平台枚举分类，跨平台稳定。
3. 责任穿透：平台给的「问题部门」是受理部门，不等于根因部门，两者分开统计。
4. 置信度标注：只匹配到宽泛词（如"退款问题"）的判为 low，需人工回溯，不许虚高。
"""

import argparse
import csv
import json
import os
import re
import statistics
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------- 字段映射

FIELD_ALIASES = {
    "product": ["商品名称", "商品标题", "宝贝标题", "商品", "品名", "product", "title"],
    "code": ["编码", "商品编码", "商品ID", "商品id", "货号", "sku编码", "code"],
    "sku": ["商品SKU", "SKU", "sku", "规格属性", "商品规格", "sku属性", "规格"],
    "sales": ["销量", "销量(件)", "近30天销量", "累计销量", "销售数量", "sales"],
    "order": ["订单编号", "订单号", "订单id", "订单ID", "子订单编号", "order_id"],
    "amount": ["订单金额(元)", "订单金额", "实付金额", "支付金额", "成交额", "付款金额", "amount"],
    "pay_time": ["付款时间", "支付时间", "下单时间", "订单创建时间", "成交时间", "pay_time"],
    "issue_time": ["首次问题发生时间", "问题发生时间", "问题时间", "投诉时间",
                   "工单创建时间", "售后申请时间", "issue_time"],
    "issue_type": ["问题类型", "问题一级分类", "体验问题类型", "一级原因", "issue_type"],
    "issue_detail": ["问题详情", "问题二级分类", "问题原因", "二级原因", "问题描述", "issue_detail"],
    "dept": ["问题部门", "责任部门", "受理部门", "归属部门", "处理部门", "dept"],
}

# ------------------------------------------------- 旅程六阶段归因规则（默认）

# 顺序即优先级：越靠前 = 越上游 = 越接近根因。
# 命中即停止，所以"质量问题→退款"会判给 S5 收货体验，而不是 S6 售后维权。
DEFAULT_STAGE_RULES = [
    ("S1 售前呈现", ["描述不符", "图文不符", "规格不符", "尺寸不符", "大小不符", "重量不符",
                     "颜色不符", "色差", "夸大", "虚假宣传", "与描述", "拍错", "买错",
                     "参数标错", "页面", "详情", "宣传"]),
    ("S2 下单支付", ["价格", "差价", "优惠", "券", "活动", "拍下", "无法下单", "超卖",
                     "改价", "重复扣款"]),
    ("S3 履约发货", ["延迟发货", "虚假发货", "未按约定", "缺货", "无货", "漏发", "少发",
                     "错发", "发错", "揽收", "发货慢", "不发货", "空包"]),
    ("S4 物流配送", ["物流", "快递", "派送", "停滞", "丢件", "破损", "配送", "签收",
                     "运输", "送货", "驿站", "派件"]),
    ("S5 收货体验", ["质量", "损坏", "坏了", "故障", "不能用", "不好用", "瑕疵", "做工",
                     "异味", "安装", "不会用", "不清晰", "噪音", "掉线", "断连", "续航",
                     "失灵", "开裂", "生锈", "缺少配件", "功能"]),
    ("S6 售后维权", ["退款", "退货", "换货", "运费", "维修", "客服", "态度", "服务差",
                     "差评", "求助平台", "投诉", "纠纷", "介入", "响应慢", "不回复",
                     "旺旺", "催"]),
]

# 只命中这些宽泛词时，归因置信度判为 low —— 意味着"受理在售后，根因在上游，需回溯"
VAGUE_TERMS = ["售后", "退款问题", "退款", "退货", "换货", "维权", "求助平台", "投诉", "纠纷"]

STAGE_OWNER = {
    "S1 售前呈现": "运营部（商品/详情页）",
    "S2 下单支付": "运营部（活动/价格）",
    "S3 履约发货": "仓储部门",
    "S4 物流配送": "仓储部门 / 承运商",
    "S5 收货体验": "采购部 / 供应商（品质）",
    "S6 售后维权": "客服部（受理）",
    "S0 未分类": "待人工回溯",
}

STAGE_ORDER = ["S1 售前呈现", "S2 下单支付", "S3 履约发货",
               "S4 物流配送", "S5 收货体验", "S6 售后维权", "S0 未分类"]


# ---------------------------------------------------------------- 读取输入

def read_csv(path):
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                rows = list(csv.DictReader(f))
            if rows:
                return rows
        except (UnicodeDecodeError, LookupError):
            continue
    raise RuntimeError("无法解码 CSV，请转为 UTF-8 编码后重试：%s" % path)


def read_json(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
        raise RuntimeError("JSON 结构中未找到记录数组")
    return data


def _xlsx_shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    return ["".join(t.text or "" for t in si.iter(ns + "t")) for si in root.iter(ns + "si")]


def read_xlsx(path):
    """零依赖读第一个工作表的首个工作表（无 openpyxl 时的降级方案）"""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as z:
        shared = _xlsx_shared_strings(z)
        sheet = next(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
        root = ET.fromstring(z.read(sheet))
    rows = []
    for row in root.iter(ns + "row"):
        cells = {}
        for c in row.iter(ns + "c"):
            ref = c.get("r", "")
            col = re.match(r"([A-Z]+)", ref)
            if not col:
                continue
            t = c.get("t")
            v = c.find(ns + "v")
            if v is None:
                cells[col.group(1)] = ""
            elif t == "s":
                cells[col.group(1)] = shared[int(v.text)]
            else:
                cells[col.group(1)] = v.text or ""
        rows.append(cells)
    if not rows:
        return []
    # 列序 → 表头
    def col_idx(name):
        n = 0
        for ch in name:
            n = n * 26 + (ord(ch) - 64)
        return n
    header_keys = sorted(rows[0].keys(), key=col_idx)
    headers = [rows[0].get(k, "") for k in header_keys]
    return [dict(zip(headers, [r.get(k, "") for k in header_keys])) for r in rows[1:]]


def load_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return read_csv(path)
    if ext == ".json":
        return read_json(path)
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    raise RuntimeError("不支持的文件类型：%s（请用 .csv / .json / .xlsx）" % ext)


# ---------------------------------------------------------------- 归一化

def build_field_map(headers):
    """把实际表头映射到标准字段名"""
    norm = {h: re.sub(r"\s+", "", str(h or "")).lower() for h in headers}
    mapping = {}
    used = set()
    for std, aliases in FIELD_ALIASES.items():
        for h in headers:
            if h in used:
                continue
            nh = norm[h]
            if any(re.sub(r"\s+", "", a).lower() == nh for a in aliases):
                mapping[std] = h
                used.add(h)
                break
    return mapping


def parse_time(s):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S",
                "%Y/%m/%d %H:%M", "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_amount(s):
    s = re.sub(r"[^\d.\-]", "", str(s or ""))
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_sku_attrs(sku_text):
    """解析 '内存容量：64GB\\n颜色分类：【500万...】' 形式的规格串"""
    attrs = {}
    for line in re.split(r"[\n;；|，,]", str(sku_text or "")):
        if "：" in line or ":" in line:
            k, v = re.split(r"[：:]", line, maxsplit=1)
            k, v = k.strip(), v.strip()
            if k and v:
                attrs[k] = v
    return attrs


def strip_tags(v):
    """去掉 SKU 值里的营销符号，取核心规格名"""
    v = re.sub(r"[【】\[\]✅⭕★☆※→⚡️\s]+", "", str(v or ""))
    return v


# ---------------------------------------------------------------- 归因

def assign_stage(text, rules):
    """返回 (阶段, 命中词, 置信度)"""
    for stage, kws in rules:
        for kw in kws:
            if kw in text:
                conf = "low" if kw in VAGUE_TERMS else "high"
                return stage, kw, conf
    return "S0 未分类", "", "none"


def bucket_days(d):
    if d is None:
        return "未知"
    if d < 0:
        return "异常（问题早于付款）"
    if d <= 3:
        return "0-3 天"
    if d <= 7:
        return "4-7 天"
    if d <= 14:
        return "8-14 天"
    if d <= 30:
        return "15-30 天"
    return "30 天以上"


DAY_BUCKETS = ["0-3 天", "4-7 天", "8-14 天", "15-30 天", "30 天以上", "未知", "异常（问题早于付款）"]


def iso_week(dt):
    if dt is None:
        return "未知"
    y, w, _ = dt.isocalendar()
    return "%d-W%02d" % (y, w)


# ---------------------------------------------------------------- 主流程

def analyze(rows, sku_keys=None, rules=None, top_n=10):
    rules = rules or DEFAULT_STAGE_RULES
    fmap = build_field_map(rows[0].keys())
    missing = [f for f in ("pay_time", "issue_type") if f not in fmap]
    if missing:
        raise RuntimeError("未能识别关键字段：%s\n实际表头：%s" %
                           (missing, list(rows[0].keys())))

    recs = []
    for r in rows:
        g = lambda k: (r.get(fmap.get(k, ""), "") or "").strip()
        pay = parse_time(g("pay_time"))
        iss = parse_time(g("issue_time"))
        days = round((iss - pay).total_seconds() / 86400, 2) if (pay and iss) else None
        raw = (g("issue_type") + " " + g("issue_detail")).strip()
        stage, hit, conf = assign_stage(raw, rules)
        rec = {
            "product": g("product"),
            "code": g("code"),
            "sku": g("sku"),
            "attrs": parse_sku_attrs(g("sku")),
            "sales": parse_amount(g("sales")),
            "order": g("order"),
            "amount": parse_amount(g("amount")),
            "pay_time": pay.isoformat(sep=" ") if pay else "",
            "issue_time": iss.isoformat(sep=" ") if iss else "",
            "days": days,
            "day_bucket": bucket_days(days),
            "pay_week": iso_week(pay),
            "issue_week": iso_week(iss),
            "issue_type": g("issue_type"),
            "issue_detail": g("issue_detail"),
            "dept_accept": g("dept"),
            "stage": stage,
            "stage_hit": hit,
            "confidence": conf,
            "stage_owner": STAGE_OWNER[stage],
        }
        recs.append(rec)

    total = len(recs)
    total_amount = sum(x["amount"] for x in recs)
    days_list = [x["days"] for x in recs if x["days"] is not None and x["days"] >= 0]

    # 阶段分布
    stage_stat = []
    for s in STAGE_ORDER:
        sub = [x for x in recs if x["stage"] == s]
        if not sub:
            continue
        stage_stat.append({
            "stage": s,
            "owner": STAGE_OWNER[s],
            "count": len(sub),
            "pct": round(len(sub) / total * 100, 1),
            "amount": round(sum(x["amount"] for x in sub), 2),
            "avg_days": round(statistics.mean([x["days"] for x in sub if x["days"] is not None]), 1)
                        if any(x["days"] is not None for x in sub) else None,
        })

    # 二级归因
    detail_counter = Counter((x["issue_type"] + " · " + x["issue_detail"]).strip(" ·")
                             for x in recs)
    detail_stat = []
    for k, c in detail_counter.most_common(top_n):
        sub = [x for x in recs if (x["issue_type"] + " · " + x["issue_detail"]).strip(" ·") == k]
        detail_stat.append({
            "key": k, "count": c,
            "amount": round(sum(x["amount"] for x in sub), 2),
            "stage": Counter(x["stage"] for x in sub).most_common(1)[0][0],
        })

    # 责任穿透：受理部门 vs 根因归属
    accept_counter = Counter(x["dept_accept"] or "未标注" for x in recs)
    penetration = []
    for dept, c in accept_counter.most_common():
        sub = [x for x in recs if (x["dept_accept"] or "未标注") == dept]
        root = Counter(x["stage_owner"] for x in sub).most_common(1)[0][0]
        misrouted = sum(1 for x in sub if x["stage_owner"] != dept and "客服" in dept)
        penetration.append({
            "dept": dept, "count": c,
            "pct": round(c / total * 100, 1),
            "main_root": root,
            "misrouted": misrouted,
        })

    # 置信度（低置信 = 需人工回溯）
    conf_counter = Counter(x["confidence"] for x in recs)

    # SKU 属性归因
    attr_keys = sku_keys or sorted({k for x in recs for k in x["attrs"]})
    sku_matrix = []
    for k in attr_keys:
        vals = Counter(strip_tags(x["attrs"].get(k, "未标注")).split("适")[0][:24]
                       for x in recs if k in x["attrs"])
        if not vals:
            continue
        sku_matrix.append({
            "attr": k,
            "values": [{"value": v, "count": c,
                        "amount": round(sum(x["amount"] for x in recs
                                            if strip_tags(x["attrs"].get(k, "")).split("适")[0][:24] == v), 2)}
                       for v, c in vals.most_common(top_n)],
        })

    # 爆发节奏
    day_stat = [{"bucket": b, "count": sum(1 for x in recs if x["day_bucket"] == b)}
                for b in DAY_BUCKETS if any(x["day_bucket"] == b for x in recs)]
    days_summary = {}
    if days_list:
        srt = sorted(days_list)
        days_summary = {
            "avg": round(statistics.mean(srt), 1),
            "median": round(statistics.median(srt), 1),
            "p90": round(srt[max(0, int(len(srt) * 0.9) - 1)], 1),
            "min": round(srt[0], 1), "max": round(srt[-1], 1),
        }

    # 周度趋势 + 集中度判定
    week_counter = Counter(x["issue_week"] for x in recs if x["issue_week"] != "未知")
    week_stat = [{"week": w, "count": c} for w, c in sorted(week_counter.items())]
    top_week_pct = round(max(week_counter.values()) / total * 100, 1) if week_counter else 0
    if top_week_pct >= 40:
        nature, nature_note = "批次性 / 集中爆发", "问题高度集中在某一周，优先怀疑该批次的来货、活动或发货波次。"
    elif top_week_pct >= 25:
        nature, nature_note = "阶段性", "有一定集中，但需与销量波次对照后判断。"
    else:
        nature, nature_note = "结构性 / 长期存在", "问题在时间上均匀分布，属于流程或商品的固有缺陷，靠单次整改难解决。"

    # 优先级排序：金额 × 单量 × 上游可控度
    prio = []
    for s in stage_stat:
        weight = s["count"] * (s["amount"] / total_amount if total_amount else 0) + s["count"]
        prio.append({"stage": s["stage"], "owner": s["owner"], "count": s["count"],
                     "amount": s["amount"], "score": round(weight, 2)})
    prio.sort(key=lambda x: -x["score"])

    metrics = {
        "meta": {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_records": total,
            "field_map": fmap,
            "products": sorted({x["product"] for x in recs if x["product"]})[:5],
            "codes": sorted({x["code"] for x in recs if x["code"]})[:5],
        },
        "overview": {
            "total": total,
            "total_amount": round(total_amount, 2),
            "avg_amount": round(total_amount / total, 2) if total else 0,
            "total_sales": round(sum(x["sales"] for x in recs), 0),
            "nature": nature,
            "nature_note": nature_note,
            "top_week_pct": top_week_pct,
        },
        "stage_stat": stage_stat,
        "detail_stat": detail_stat,
        "penetration": penetration,
        "confidence": dict(conf_counter),
        "sku_matrix": sku_matrix,
        "day_stat": day_stat,
        "days_summary": days_summary,
        "week_stat": week_stat,
        "priority": prio,
        "records": recs,
    }
    return metrics


def main():
    ap = argparse.ArgumentParser(description="体验问题订单分析引擎")
    ap.add_argument("--input", required=True, help="csv / json / xlsx 订单表")
    ap.add_argument("--output", default="metrics.json", help="输出指标文件")
    ap.add_argument("--sku-keys", default="", help="要拆解的 SKU 属性名，逗号分隔")
    ap.add_argument("--rules", default="", help="自定义归因规则 JSON（阶段→关键词列表）")
    ap.add_argument("--top", type=int, default=10)
    args = ap.parse_args()

    rows = load_rows(args.input)
    if not rows:
        sys.exit("输入为空：%s" % args.input)

    rules = None
    if args.rules:
        with open(args.rules, "r", encoding="utf-8") as f:
            rules = [(k, v) for k, v in json.load(f).items()]

    sku_keys = [k.strip() for k in args.sku_keys.split(",") if k.strip()] or None
    m = analyze(rows, sku_keys=sku_keys, rules=rules, top_n=args.top)

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=1)

    o = m["overview"]
    print("已分析 %d 条，涉及金额 ¥%.2f，问题性质：%s" % (o["total"], o["total_amount"], o["nature"]))
    print("阶段分布：" + " | ".join("%s %d单(%.1f%%)" % (s["stage"], s["count"], s["pct"])
                                    for s in m["stage_stat"]))
    print("低置信待回溯：%d 条" % m["confidence"].get("low", 0))
    print("指标已写入 %s" % args.output)


if __name__ == "__main__":
    main()
