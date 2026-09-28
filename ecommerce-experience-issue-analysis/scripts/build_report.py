#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
把 analyze.py 产出的 metrics.json 渲染成单文件 HTML 报告（离线可看，无外部依赖）。

用法：
    python build_report.py --metrics metrics.json --output 体验问题分析报告.html
"""

import argparse
import html
import json
from datetime import datetime

CSS = """
*{box-sizing:border-box}
body{margin:0;padding:32px 20px 60px;background:#f5f7fa;color:#1f2937;
font-family:"Microsoft YaHei","PingFang SC",-apple-system,"Segoe UI",sans-serif;line-height:1.7}
.wrap{max-width:1080px;margin:0 auto}
h1{font-size:26px;margin:0 0 6px;color:#0f2744;letter-spacing:.5px}
.sub{color:#64748b;font-size:13px;margin-bottom:26px}
section{background:#fff;border:1px solid #e3e8ef;border-radius:14px;padding:24px 26px;margin-bottom:20px;
box-shadow:0 1px 3px rgba(15,39,68,.05)}
h2{font-size:18px;margin:0 0 4px;color:#0f2744;display:flex;align-items:center;gap:9px}
h2 .num{display:inline-flex;width:24px;height:24px;border-radius:7px;background:#0f2744;color:#fff;
font-size:13px;align-items:center;justify-content:center;flex:0 0 24px}
.note{color:#64748b;font-size:12.5px;margin:0 0 16px;padding-left:33px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px}
.kpi{background:#fff;border:1px solid #e3e8ef;border-radius:12px;padding:14px 16px}
.kpi .v{font-size:24px;font-weight:700;color:#0f2744}
.kpi .l{font-size:12px;color:#64748b;margin-top:2px}
.kpi.warn .v{color:#c0392b}
.kpi.ok .v{color:#1e7d5a}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th,td{padding:9px 10px;border-bottom:1px solid #eef2f7;text-align:left}
th{background:#f8fafc;color:#475569;font-weight:600;font-size:12.5px}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
.bar{position:relative;background:#eef2f7;border-radius:5px;height:22px;min-width:120px}
.bar i{position:absolute;left:0;top:0;bottom:0;border-radius:5px;background:#3b6ea5}
.bar.s1 i{background:#8b5cf6}.bar.s2 i{background:#0ea5e9}.bar.s3 i{background:#f59e0b}
.bar.s4 i{background:#6366f1}.bar.s5 i{background:#c0392b}.bar.s6 i{background:#475569}
.bar em{position:absolute;left:8px;top:0;font-style:normal;font-size:12px;color:#fff;line-height:22px}
.tag{display:inline-block;padding:1px 8px;border-radius:20px;font-size:11.5px;border:1px solid}
.tag.hi{background:#fef2f2;color:#b91c1c;border-color:#fecaca}
.tag.md{background:#fffbeb;color:#b45309;border-color:#fde68a}
.tag.lo{background:#f0fdf4;color:#15803d;border-color:#bbf7d0}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
ul{margin:8px 0 0;padding-left:20px}li{margin-bottom:7px;font-size:13.5px}
b{color:#0f2744}
.callout{background:#f8fafc;border-left:3px solid #0f2744;padding:12px 16px;border-radius:0 8px 8px 0;
font-size:13.5px;margin:12px 0 0;color:#334155}
.callout.red{border-left-color:#c0392b;background:#fef6f6}
.scroll{max-height:520px;overflow:auto;border:1px solid #eef2f7;border-radius:10px}
.scroll table{font-size:12.5px}.scroll th{position:sticky;top:0;z-index:1}
.small{font-size:12px;color:#94a3b8}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}
@media(max-width:760px){.grid2{grid-template-columns:1fr}}
"""

STAGE_COLOR = {"S1 售前呈现": "#8b5cf6", "S2 下单支付": "#0ea5e9", "S3 履约发货": "#f59e0b",
               "S4 物流配送": "#6366f1", "S5 收货体验": "#c0392b", "S6 售后维权": "#475569",
               "S0 未分类": "#94a3b8"}

STAGE_ACTION = {
    "S1 售前呈现": ["逐条核对详情页参数、主图与实物一致性，重点核查被投诉的规格描述",
                    "对夸大/绝对化用语做合规清洗；页面增加「实物与显示器色差」等预期管理提示"],
    "S2 下单支付": ["校验活动价、券后价与到手价是否一致，避免下单后改价争议",
                    "超卖/缺货前置拦截，库存同步频率提升到小时级"],
    "S3 履约发货": ["建立发货时效 SLA 看板，超时订单自动预警到仓储主管",
                    "缺货订单在承诺时效前主动告知并提供改配/退款选项，禁止虚假发货"],
    "S4 物流配送": ["按承运商统计异常件率并纳入月度考核，末位淘汰或调线",
                    "派送停滞超 48 小时自动触发客服主动介入"],
    "S5 收货体验": ["对本次高发规格做来货抽检加严，出具质检问题描述清单并追溯批次",
                    "把高频故障点前置到说明书/安装视频；质量问题转供应商索赔"],
    "S6 售后维权": ["设定响应时长与一次解决率 SLA，超时自动升级",
                    "退款/运费规则在售后页透明展示；差评订单 24 小时内回访闭环"],
    "S0 未分类": ["人工回溯原始聊天/工单后补录归因，再进入下一轮分析"],
}

STAGE_KPI = {
    "S1 售前呈现": "描述相符评分回升、该类投诉单量下降",
    "S2 下单支付": "价格类纠纷单量、改价工单数",
    "S3 履约发货": "发货超时率、虚假发货投诉数",
    "S4 物流配送": "异常件率、派送停滞工单数",
    "S5 收货体验": "该规格退货率、质检不合格率",
    "S6 售后维权": "客服响应时长、一次解决率、差评率",
    "S0 未分类": "归因完整率（无未分类残留）",
}

e = html.escape


def bar(pct, label, cls="", maxpct=100.0):
    w = max(6.0, pct / max(maxpct, 1) * 100)
    return ('<div class="bar %s"><i style="width:%.1f%%"></i><em>%s</em></div>'
            % (cls, w, e(str(label))))


def build(m):
    o = m["overview"]
    meta = m["meta"]
    title_parts = [p for p in (meta.get("products") or [""]) [:1] if p]
    code = (meta.get("codes") or [""]) [0]
    title = "消费者体验问题订单分析报告"
    if title_parts:
        title += " · " + title_parts[0]

    P = []
    A = P.append

    A('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">')
    A('<meta name="viewport" content="width=device-width,initial-scale=1">')
    A('<title>%s</title><style>%s</style></head><body><div class="wrap">' % (e(title), CSS))
    A('<h1>%s</h1>' % e(title))
    A('<div class="sub">商品编码 %s ｜ 数据口径：平台体验问题订单 ｜ 生成时间 %s ｜ '
      '样本 %d 条</div>' % (e(code or "—"), e(meta["generated_at"]), o["total"]))

    # KPI
    low = m["confidence"].get("low", 0)
    A('<div class="kpis">')
    A('<div class="kpi"><div class="v">%d</div><div class="l">问题订单数</div></div>' % o["total"])
    A('<div class="kpi warn"><div class="v">¥%s</div><div class="l">涉及订单金额</div></div>'
      % format(o["total_amount"], ",.0f"))
    A('<div class="kpi"><div class="v">¥%s</div><div class="l">问题订单客单价</div></div>'
      % format(o["avg_amount"], ",.0f"))
    A('<div class="kpi %s"><div class="v">%s</div><div class="l">问题性质</div></div>'
      % ("warn" if "结构性" in o["nature"] or "批次" in o["nature"] else "ok", e(o["nature"])))
    A('<div class="kpi %s"><div class="v">%d</div><div class="l">根因待回溯（低置信）</div></div>'
      % ("warn" if low else "ok", low))
    A('</div>')

    # ① 阶段定位
    A('<section><h2><span class="num">1</span>问题发生在消费者旅程的哪个环节</h2>')
    A('<p class="note">按旅程阶段归因，而非平台的枚举分类——同一句「退款问题」可能来自质量、发货或描述不符。</p>')
    A('<table><thead><tr><th style="width:150px">旅程阶段</th><th style="width:230px">占比</th>'
      '<th class="n">单量</th><th class="n">涉及金额</th><th class="n">平均爆发天数</th>'
      '<th>根因责任归属</th></tr></thead><tbody>')
    maxpct = max([s["pct"] for s in m["stage_stat"]] or [1])
    for i, s in enumerate(m["stage_stat"], 1):
        A('<tr><td><span class="dot" style="background:%s"></span>%s</td>'
          '<td>%s</td><td class="n">%d</td><td class="n">¥%s</td><td class="n">%s</td><td>%s</td></tr>'
          % (STAGE_COLOR.get(s["stage"], "#94a3b8"), e(s["stage"]),
             bar(s["pct"], "%.1f%%" % s["pct"], "s%d" % i, maxpct),
             s["count"], format(s["amount"], ",.0f"),
             s["avg_days"] if s["avg_days"] is not None else "—", e(s["owner"])))
    A('</tbody></table>')
    top = m["stage_stat"][0] if m["stage_stat"] else None
    if top:
        A('<div class="callout"><b>判读：</b>问题最集中在「%s」（%d 单，占 %.1f%%），根责在<b>%s</b>。'
          '若该阶段不是当前被追责的部门，说明责任被误挂在下游（通常是客服），需穿透。</div>'
          % (e(top["stage"]), top["count"], top["pct"], e(top["owner"])))
    A('</section>')

    # ② 责任穿透
    A('<section><h2><span class="num">2</span>责任穿透：受理部门 ≠ 根因部门</h2>')
    A('<p class="note">平台字段里的「问题部门」记录的是<b>谁在处理</b>，不是<b>谁造成</b>。'
      '这一层用来防止把上游的锅算在客服头上。</p>')
    A('<table><thead><tr><th>平台标注部门（受理）</th><th class="n">单量</th><th class="n">占比</th>'
      '<th>推断的主要根因归属</th><th class="n">疑似误挂条数</th></tr></thead><tbody>')
    for p in m["penetration"]:
        tag = '<span class="tag hi">需穿透</span>' if p["misrouted"] > 0 else '<span class="tag lo">一致</span>'
        A('<tr><td><b>%s</b></td><td class="n">%d</td><td class="n">%.1f%%</td><td>%s</td>'
          '<td class="n">%d %s</td></tr>'
          % (e(p["dept"]), p["count"], p["pct"], e(p["main_root"]), p["misrouted"], tag))
    A('</tbody></table>')
    if low:
        A('<div class="callout red"><b>注意：</b>有 <b>%d 条</b>只命中「退款/售后/投诉」等宽泛词，'
          '根因无法从字段判断，标注为低置信。这部分必须人工回溯原始聊天/工单后再下结论，'
          '不要直接算作客服责任。</div>' % low)
    A('</section>')

    # ③ 二级归因
    A('<section><h2><span class="num">3</span>问题详情 · 二级归因拆解</h2>')
    A('<p class="note">按「问题类型 + 问题详情」聚合，看具体是什么在反复发生。</p>')
    A('<table><thead><tr><th>问题组合</th><th class="n">单量</th><th class="n">涉及金额</th>'
      '<th>归入阶段</th></tr></thead><tbody>')
    maxc = max([d["count"] for d in m["detail_stat"]] or [1])
    for d in m["detail_stat"]:
        A('<tr><td>%s</td><td class="n">%d<br>%s</td><td class="n">¥%s</td><td>%s</td></tr>'
          % (e(d["key"] or "—"), d["count"], bar(d["count"], "", "", maxc),
             format(d["amount"], ",.0f"), e(d["stage"])))
    A('</tbody></table></section>')

    # ④ SKU 归因
    if m["sku_matrix"]:
        A('<section><h2><span class="num">4</span>SKU 规格归因 · 问题落在哪些属性上</h2>')
        A('<p class="note">把问题摊到规格属性上，区分是「全系通病」还是「某个规格独有」。</p>')
        A('<div class="grid2">')
        for sm in m["sku_matrix"][:4]:
            A('<div><table><thead><tr><th>%s</th><th class="n">单量</th><th class="n">金额</th>'
              '</tr></thead><tbody>' % e(sm["attr"]))
            for v in sm["values"]:
                hot = v["count"] >= max(2, sum(x["count"] for x in sm["values"]) * 0.4)
                A('<tr><td>%s %s</td><td class="n">%d</td><td class="n">¥%s</td></tr>'
                  % ('<span class="tag hi">重点</span>' if hot else "", e(v["value"] or "—"),
                     v["count"], format(v["amount"], ",.0f")))
            A('</tbody></table></div>')
        A('</div>')
        hot_attr = next((s for s in m["sku_matrix"] if s["values"]), None)
        if hot_attr:
            A('<div class="callout"><b>抽检重点：</b>「%s」维度以 <b>%s</b> 最集中，'
              '建议作为下一轮来货质检与批次追溯的第一优先级。</div>'
              % (e(hot_attr["attr"]), e(hot_attr["values"][0]["value"])))
        A('</section>')

    # ⑤ 爆发节奏
    A('<section><h2><span class="num">5</span>爆发节奏 · 付款到问题发生用了多久</h2>')
    A('<p class="note">爆发节奏用来判断问题性质：集中在某几天 = 批次/活动问题；均匀铺开 = 结构性缺陷。</p>')
    A('<div class="grid2"><div>')
    A('<table><thead><tr><th>付款→问题间隔</th><th class="n">单量</th><th>分布</th></tr></thead><tbody>')
    maxd = max([d["count"] for d in m["day_stat"]] or [1])
    for d in m["day_stat"]:
        A('<tr><td>%s</td><td class="n">%d</td><td>%s</td></tr>'
          % (e(d["bucket"]), d["count"], bar(d["count"], "", "", maxd)))
    A('</tbody></table></div><div>')
    ds = m.get("days_summary") or {}
    A('<table><thead><tr><th>节奏指标</th><th class="n">天数</th></tr></thead><tbody>')
    for k, label in (("min", "最快爆发"), ("median", "中位天数"), ("avg", "平均天数"),
                     ("p90", "P90（90% 在此之内）"), ("max", "最慢爆发")):
        if k in ds:
            A('<tr><td>%s</td><td class="n">%s</td></tr>' % (label, ds[k]))
    A('</tbody></table>')
    A('<div class="callout"><b>%s</b>（单周最高占比 %.1f%%）：%s</div>'
      % (e(o["nature"]), o["top_week_pct"], e(o["nature_note"])))
    if ds.get("median") is not None and ds["median"] <= 3:
        A('<div class="callout red">中位爆发仅 %s 天，说明问题在<b>收货前后就已显现</b>，'
          '属于商品/发货端问题，客服端拦截无效，必须前置整改。</div>' % ds["median"])
    A('</div></div></section>')

    # ⑥ 明细
    A('<section><h2><span class="num">6</span>问题订单明细</h2>')
    A('<p class="note">按付款时间升序，保留归因命中词与置信度，便于逐条复核。</p>')
    A('<div class="scroll"><table><thead><tr><th>订单号</th><th>付款时间</th><th>问题时间</th>'
      '<th class="n">天数</th><th class="n">金额</th><th>问题类型</th><th>详情</th>'
      '<th>阶段</th><th>置信</th></tr></thead><tbody>')
    conf_tag = {"high": "lo", "low": "md", "none": "hi"}
    conf_txt = {"high": "高", "low": "低", "none": "未归类"}
    for r in sorted(m["records"], key=lambda x: x["pay_time"] or "9999"):
        A('<tr><td class="small">%s</td><td class="small">%s</td><td class="small">%s</td>'
          '<td class="n">%s</td><td class="n">%s</td><td>%s</td><td>%s</td><td>%s</td>'
          '<td><span class="tag %s">%s</span></td></tr>'
          % (e(r["order"][-12:] or "—"), e(r["pay_time"][:10]), e(r["issue_time"][:10]),
             r["days"] if r["days"] is not None else "—", format(r["amount"], ",.0f"),
             e(r["issue_type"] or "—"), e(r["issue_detail"] or "—"), e(r["stage"]),
             conf_tag.get(r["confidence"], "hi"), conf_txt.get(r["confidence"], "?")))
    A('</tbody></table></div></section>')

    # ⑦ 行动建议
    A('<section><h2><span class="num">7</span>改进动作建议 · 含责任部门与验收指标</h2>')
    A('<p class="note">按「单量 × 金额 × 上游可控度」排出的优先级，先看第 1 条。</p>')
    for i, p in enumerate(m["priority"][:5], 1):
        acts = STAGE_ACTION.get(p["stage"], [])
        A('<div style="border:1px solid #e3e8ef;border-radius:10px;padding:14px 16px;margin-bottom:12px">')
        A('<div style="font-size:14.5px;margin-bottom:6px"><b>%d. %s</b> '
          '<span class="small">（%d 单 · ¥%s）</span> '
          '<span class="tag md">责任：%s</span></div>'
          % (i, e(p["stage"]), p["count"], format(p["amount"], ",.0f"), e(p["owner"])))
        A('<ul>')
        for a in acts:
            A('<li>%s</li>' % e(a))
        A('<li><b>验收指标：</b>%s</li>' % e(STAGE_KPI.get(p["stage"], "该类问题单量下降")))
        A('</ul></div>')
    if low:
        A('<div class="callout red"><b>前置动作：</b>先把 %d 条低置信订单回溯归因，'
          '再执行上面的整改——否则可能整改错部门。</div>' % low)
    A('</section>')

    A('<div class="sub" style="text-align:center;margin-top:30px">'
      '本报告由「消费者体验问题分析」技能生成 · 归因结论供决策参考，低置信项须人工复核</div>')
    A('</div></body></html>')
    return "".join(P)


def main():
    ap = argparse.ArgumentParser(description="渲染体验问题分析 HTML 报告")
    ap.add_argument("--metrics", required=True)
    ap.add_argument("--output", default="体验问题分析报告.html")
    args = ap.parse_args()
    with open(args.metrics, "r", encoding="utf-8") as f:
        m = json.load(f)
    out = build(m)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(out)
    print("报告已生成：%s（%.1f KB）" % (args.output, len(out.encode("utf-8")) / 1024))


if __name__ == "__main__":
    main()
