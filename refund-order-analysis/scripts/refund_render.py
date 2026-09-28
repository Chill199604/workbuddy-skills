# -*- coding: utf-8 -*-
"""
退款单数据分析 · 渲染模块（可复用）
================================
依赖 refund_analysis.run() 的结果，生成：
  1) 交互式 HTML 报告（自绘 SVG，无需联网）
  2) 可下载 XLSX 明细表（板块汇总 / 超1000清单 / 品退清单 / TOP商品明细）

图表选型原则（来自本 Skill 的 chart_selection 参考）：
  比大小偶尔条形 · 看趋势必折 · 讲构成用饼 · 说分布画方 · 对比用进度条。
  不堆叠同类型图，保持数据完整 + 观感清爽。

使用：
  cd <skill>/scripts && python refund_render.py
或在别处调用： python refund_render.py
"""
import importlib, sys, os, math
from collections import defaultdict, Counter

# 把本脚本所在目录加入 path，确保能 import 同目录的 refund_analysis
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import refund_analysis as A

# ============================ 输出路径配置 ============================
OUT_DIR = r"D:\serious"
OUT_HTML = os.path.join(OUT_DIR, "退款单分析报告.html")
OUT_XLSX = os.path.join(OUT_DIR, "退款单明细表.xlsx")
SHOP_NAME = "喵住萤石"          # 报告标题用店铺名
PERIOD_LABEL = "近90天（2026-06-27 ~ 2026-09-27）"
# ======================================================================

idx = A.IDX
fnum = A.fnum
COLORS = A.COLORS
MAIN = A.MAIN
TOTAL_SALES = A.TOTAL_SALES
bucket = A.bucket
is_pintui = A.is_pintui

R = A.run()
globals().update(R)   # 把 run() 返回的分析结果提升为模块变量

def money(x):
    return f"{x:,.0f}"

# ---------------- SVG helpers ----------------
def svg_pie(items, size=280):
    total = sum(v for _, v, _ in items)
    cx = cy = size / 2; r = size / 2 - 34
    out = [f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" xmlns="http://www.w3.org/2000/svg">']
    if total <= 0:
        out.append(f'<text x="{cx}" y="{cy}" text-anchor="middle" font-size="13" fill="#999">无数据</text></svg>')
        return "".join(out)
    ang = 0.0
    for label, value, color in items:
        frac = value / total
        a0 = ang; a1 = ang + frac * 360; ang = a1
        x0 = cx + r * math.cos(math.radians(a0 - 90)); y0 = cy + r * math.sin(math.radians(a0 - 90))
        x1 = cx + r * math.cos(math.radians(a1 - 90)); y1 = cy + r * math.sin(math.radians(a1 - 90))
        large = 1 if frac > 0.5 else 0
        out.append(f'<path d="M{cx:.1f},{cy:.1f} L{x0:.1f},{y0:.1f} A{r:.1f},{r:.1f} 0 {large} 1 {x1:.1f},{y1:.1f} Z" fill="{color}" stroke="#fff" stroke-width="1.5"/>')
        if frac > 0.04:
            mid = (a0 + a1) / 2 - 90
            lx = cx + (r * 0.62) * math.cos(math.radians(mid)); ly = cy + (r * 0.62) * math.sin(math.radians(mid))
            out.append(f'<text x="{lx:.1f}" y="{ly:.1f}" font-size="11" fill="#fff" text-anchor="middle" font-weight="bold">{frac*100:.1f}%</text>')
    out.append('</svg>')
    return "".join(out)

def svg_hbar(items, maxw=360, rowh=28, color="#4C72B0", colors=None, label_width=150):
    maxv = max((v for _, v in items), default=1) or 1
    max_val_len = max((len(f"{v:,.0f}") for _, v in items), default=1) * 7 + 10
    W = label_width + maxw + max_val_len + 14
    H = rowh * len(items) + 14
    out = [f'<svg viewBox="0 0 {W} {H}" width="100%" height="{H}" xmlns="http://www.w3.org/2000/svg">']
    for i, (label, value) in enumerate(items):
        y = i * rowh + 8
        w = value / maxv * maxw
        c = (colors[i] if colors else color) or "#4C72B0"
        max_chars = max(1, int((label_width - 8) / 6.5))
        show_label = label if len(label) <= max_chars else label[:max_chars - 1] + "…"
        out.append(f'<g><title>{label}: {value:,.0f}</title>')
        out.append(f'<text x="0" y="{y+rowh-11}" font-size="12" fill="#333">{show_label}</text>')
        out.append(f'<rect x="{label_width}" y="{y}" width="{max(w,0.5):.1f}" height="{rowh-10}" fill="{c}" rx="3"/>')
        val_text = f"{value:,.0f}"
        if w > maxw * 0.72:
            out.append(f'<text x="{label_width+w-6:.1f}" y="{y+rowh-11}" font-size="11" fill="#fff" text-anchor="end">{val_text}</text>')
        else:
            out.append(f'<text x="{label_width+w+6:.1f}" y="{y+rowh-11}" font-size="11" fill="#555">{val_text}</text>')
        out.append('</g>')
    out.append('</svg>')
    return "".join(out)

def short_id(s, n=12):
    if s and len(str(s)) > n:
        return str(s)[:n] + "…"
    return str(s)

def svg_line(series, labels, W=760, H=300, colors=None, ylabel="千元"):
    n = len(labels)
    padL = 46; padB = 38; padT = 14; padR = 14
    plotW = W - padL - padR; plotH = H - padT - padB
    allv = [v for _, vals in series for v in vals]
    maxv = max(allv) if allv else 1
    if maxv <= 0: maxv = 1
    bw = plotW / n
    def color(i): return (colors[i] if colors else "#4C72B0")
    out = [f'<svg viewBox="0 0 {W} {H}" width="100%" height="{H}" xmlns="http://www.w3.org/2000/svg">']
    for g in range(6):
        yv = maxv * g / 5; y = padT + plotH - (yv / maxv * plotH)
        out.append(f'<line x1="{padL}" y1="{y:.1f}" x2="{W-padR}" y2="{y:.1f}" stroke="#eee"/>')
        out.append(f'<text x="{padL-6}" y="{y+3:.1f}" font-size="9" fill="#aaa" text-anchor="end">{yv/1000:.0f}k</text>')
    for si, (name, vals) in enumerate(series):
        c = color(si)
        pts = []
        for i, v in enumerate(vals):
            x = padL + (i + 0.5) * bw; y = padT + plotH - (v / maxv * plotH)
            pts.append(f"{x:.1f},{y:.1f}")
        out.append(f'<polyline points="{" ".join(pts)}" fill="none" stroke="{c}" stroke-width="2.2" stroke-linejoin="round"/>')
        for i, v in enumerate(vals):
            x = padL + (i + 0.5) * bw; y = padT + plotH - (v / maxv * plotH)
            out.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.6" fill="{c}"/>')
    for i, lab in enumerate(labels):
        x = padL + (i + 0.5) * bw
        out.append(f'<text x="{x}" y="{H-16}" font-size="9" fill="#888" text-anchor="middle">{lab}</text>')
    out.append(f'<text x="{padL-6}" y="{padT-3}" font-size="9" fill="#aaa" text-anchor="end">{ylabel}</text>')
    out.append('</svg>')
    return "".join(out)

def svg_progress2(items, W=720):
    barw = W - 180; rowh = 30; H = rowh * len(items) + 14
    out = [f'<svg viewBox="0 0 {W} {H}" width="100%" height="{H}" xmlns="http://www.w3.org/2000/svg">']
    for i, (label, pct, c) in enumerate(items):
        y = 8 + i * rowh
        out.append(f'<text x="0" y="{y+15}" font-size="12.5" fill="#333">{label}</text>')
        out.append(f'<rect x="150" y="{y}" width="{barw}" height="18" fill="#eef1f6" rx="9"/>')
        out.append(f'<rect x="150" y="{y}" width="{max(barw*pct/100,2):.1f}" height="18" fill="{c}" rx="9"/>')
        out.append(f'<text x="{150+barw+8}" y="{y+15}" font-size="12.5" fill="#333" font-weight="bold">{pct:.2f}%</text>')
    out.append('</svg>')
    return "".join(out)

def legend(items):
    total = sum(v for _, v, _ in items)
    if not total: return ""
    return "".join(f'<span class="lg"><i style="background:{c}"></i>{l} {money(v)}（{v/total*100:.1f}%）</span>' for l, v, c in items)

# ---------------- Section: 数据概况 ----------------
closed_amt = sum(fnum(r[idx["退款总额"]]) for r in closed)
s0 = f"""
<section><h2>数据概况与口径</h2>
<ul class="rs">
<li>数据范围：{PERIOD_LABEL}，单 Sheet <code>{A.SHEET_NAME}</code>，共 {len(DATA)} 行退款单明细。</li>
<li>退款状态：退款成功 <b>{len(succ)}</b> 单 / 退款关闭 <b>{len(closed)}</b> 单（关闭单申请金额 {money(closed_amt)} 元，<b>不计入退款金额与退款率</b>）。</li>
<li>金额字段采用 <code>退款总额</code>（= 退给买家 + 退给平台，即整单逆转额）；四类板块金额相加 = 退款成功总额，无遗漏。</li>
<li>换货/补寄/维修共 {b_count.get('其他',0)} 单，退款金额均为 0，<b>不计入退款率</b>，仅作说明。</li>
<li>分母：店铺总销售额 {money(TOTAL_SALES)}。</li>
</ul></section>
"""

# ---------------- Section 1: 未发货 vs 发货后 ----------------
ship_items = [("未发货", ship_amt["未发货"], COLORS["未发货"]), ("发货后", ship_amt["发货后"], COLORS["发货后"])]
s1 = f"""
<section><h2>一、未发货退款率 vs 发货后退款率（首要结论）</h2>
<div class="note">口径：仅统计<strong>退款成功</strong>单；剔除「换货/补寄/维修」与「退款关闭」。分母=店铺总销售额 {money(TOTAL_SALES)}。</div>
<div class="cards">
  <div class="card"><div class="k">未发货退款率</div><div class="v" style="color:{COLORS['未发货']}">{unship_rate*100:.1f}%</div><div class="s">{money(ship_amt['未发货'])} 元 · {ship_count['未发货']} 单</div></div>
  <div class="card"><div class="k">发货后退款率</div><div class="v" style="color:{COLORS['发货后']}">{ship_rate*100:.1f}%</div><div class="s">{money(ship_amt['发货后'])} 元 · {ship_count['发货后']} 单</div></div>
  <div class="card"><div class="k">整体退款率(不含换补维)</div><div class="v">{overall_rate*100:.1f}%</div><div class="s">{money(total_refund)} 元 · {sum(ship_count.values())} 单</div></div>
</div>
<div class="sub">退款率对比：</div>
{svg_progress2([("未发货退款率", unship_rate*100, COLORS['未发货']), ("发货后退款率", ship_rate*100, COLORS['发货后']), ("整体退款率", overall_rate*100, "#888")])}
<div class="chartwrap">{svg_pie(ship_items, size=300)}<div class="legend">{legend(ship_items)}</div></div>
<p class="insight"><strong>解读：</strong>近 {ship_amt['未发货']/total_refund*100:.1f}% 的退款金额发生在<strong>发货之前</strong>，发货后退款率仅 {ship_rate*100:.1f}%。说明大量退款是"下单即反悔/拍错/协商一致"等<strong>售前决策型退款</strong>。降未发货退款应主攻<strong>详情页准确性、拍错拦截、客服确认</strong>；发货后退款（尤其退货退款）才更值得从品控与履约角度治理。</p></section>
"""

# ---------------- Section 2: 四大板块 ----------------
b_items = [(b, b_amt[b], COLORS[b]) for b in MAIN]
b_rows = ""
for b in MAIN:
    b_rows += f"<tr><td>{b}</td><td>{b_count[b]}</td><td>{money(b_amt[b])}</td><td>{b_amt[b]/total_refund*100:.1f}%</td><td>{b_amt[b]/TOTAL_SALES*100:.1f}%</td><td>{money(b_amt[b]/max(b_count[b],1))}</td><td>{big_n[b]}（{money(big_a[b])}）</td><td>{fp_c[b].get('全额退款',0)}全 / {fp_c[b].get('部分退款',0)}部</td></tr>"
rh = ""
for b in MAIN:
    top3 = reason_c[b].most_common(3)
    rh += f"<li><b>{b}</b>：" + "、".join(f"{k}({v})" for k, v in top3) + "</li>"
s2 = f"""
<section><h2>二、四大板块退款明细</h2>
<table><thead><tr><th>板块</th><th>单数</th><th>退款金额</th><th>占退款总额</th><th>板块退款率</th><th>均单</th><th>超1000元</th><th>全额/部分</th></tr></thead><tbody>{b_rows}</tbody></table>
<div class="chartwrap">{svg_pie(b_items, size=300)}<div class="legend">{legend(b_items)}</div></div>
<p class="sub">各板块主要退款原因 TOP3：</p><ul class="rs">{rh}</ul>
<h3>板块周趋势（折线）</h3>
{svg_line([(b, week_amt_arr[b]) for b in MAIN], [w.split(" ")[0] for w in week_labels], colors=[COLORS[b] for b in MAIN])}
<div class="legend">{"".join('<span class="lg"><i style="background:%s"></i>%s</span>' % (COLORS[b], b) for b in MAIN)}</div>
<h3>板块月趋势（折线，千元）</h3>
{svg_line([("退款金额", [sum(month_amt_arr[b][i] for b in MAIN) for i in range(len(months))])], months, colors=["#4C72B0"], ylabel="千元")}
</section>
"""

# ---------------- Section 3: 超1000 ----------------
bands = {"1000-2000":0,"2000-3000":0,"3000+":0}
big_by_b = Counter(); big_by_b_amt = defaultdict(float); big_reason = Counter(); big_prod_c = Counter()
for r in big_rows:
    a = fnum(r[idx["退款总额"]])
    if a < 2000: bands["1000-2000"] += 1
    elif a < 3000: bands["2000-3000"] += 1
    else: bands["3000+"] += 1
    bb = bucket(r); big_by_b[bb] = big_by_b.get(bb,0)+1; big_by_b_amt[bb] += a
    big_reason[r[idx["买家退款原因"]]] += 1
    big_prod_c[r[idx["商品id"]]] += 1
bb_items = [(k, big_by_b_amt[k], COLORS[k]) for k in MAIN if k in big_by_b]
band_items = [(k, bands[k]) for k in ["1000-2000","2000-3000","3000+"] if bands[k]]
big_reason_items = [(k, v) for k, v in big_reason.most_common(5)]
s3 = f"""
<section><h2>三、超 1000 元退款单分析（{len(big_rows)} 单 / {money(sum(fnum(r[idx['退款总额']]) for r in big_rows))} 元）</h2>
<div class="note">占退款成功总单数 {len(big_rows)/len(succ)*100:.1f}%，占退款总额 {sum(fnum(r[idx['退款总额']]) for r in big_rows)/total_refund*100:.1f}%，属高金额长尾。</div>
<div class="twocol">
<div><h3>金额区间分布（单数组）</h3>{svg_hbar(band_items, color="#C44E52", maxw=300)}<div class="chart-hint">单位：单数</div></div>
<div><h3>板块归属（金额占比）</h3>{svg_pie(bb_items, size=260)}<div class="legend">{legend(bb_items)}</div></div>
</div>
<p class="sub">大额单高频退款原因（TOP5）：</p>
{svg_hbar(big_reason_items, color="#8172B3", maxw=560)}
<p class="insight"><strong>提示：</strong>大额单集中在 {big_by_b.most_common(1)[0][0] if big_by_b else '—'} 板块，重点排查是否含误发、重复下单或异常订单。涉及最多的商品 ID 已在第四节完整排名，此处不再单列。完整 {len(big_rows)} 单明细见附件《退款单明细表.xlsx》的「超1000清单」。</p></section>
"""

# ---------------- Section 4: TOP 商品 ----------------
tp_rows = ""
for pid, p in top_prod:
    tr = p["reasons"].most_common(1)
    pr = (p["pa"] / p["amt"] * 100) if p["amt"] else 0
    tp_rows += f"<tr><td>{pid}</td><td>{p['n']}</td><td>{money(p['amt'])}</td><td>{money(p['amt']/p['n'])}</td><td>{p['pn']}</td><td>{pr:.1f}%</td><td>{tr[0][0] if tr else '—'}</td></tr>"
s4 = f"""
<section><h2>四、退款金额最高的商品 ID（TOP20，按退款金额）</h2>
<div class="note">局限：源数据无单品销售额，此处为<strong>绝对退款额排名</strong>，非单品退款率。如需单品退款率请补充各商品销售额。</div>
<table><thead><tr><th>商品ID</th><th>单数</th><th>退款金额</th><th>均单</th><th>品退单数</th><th>品退占额比</th><th>首要退款原因</th></tr></thead><tbody>{tp_rows}</tbody></table>
<h3>TOP12 商品退款金额</h3>{svg_hbar([(short_id(pid), p["amt"]) for pid, p in top_prod[:12]], color="#4C72B0", maxw=560)}
</section>
"""

# ---------------- Section 5: 品退 ----------------
pin_items = [("品退(广义)", pintui_a, "#C44E52"), ("非品退", total_refund - pintui_a, "#999999")]
pin_reason_items = [(k, v) for k, v in pintui_reason.most_common(5)]
s5 = f"""
<section><h2>五、申请品退（广义·商品问题类）订单分析</h2>
<div class="note">口径：买家退款原因含 质量/瑕疵/破损/污渍/描述不符/材质/少件/漏发/发错货/缺货/做工。共 {pintui_n} 单 / {money(pintui_a)} 元，占退款总额 {pintui_a/total_refund*100:.1f}%。</div>
<div class="cards">
  <div class="card"><div class="k">品退单数</div><div class="v" style="color:#C44E52">{pintui_n}</div><div class="s">占成功单 {pintui_n/len(succ)*100:.1f}%</div></div>
  <div class="card"><div class="k">品退金额</div><div class="v" style="color:#C44E52">{money(pintui_a)}</div><div class="s">占退款额 {pintui_a/total_refund*100:.1f}%</div></div>
  <div class="card"><div class="k">商家责任</div><div class="v">{pintui_resp.get('商家责任',0)}</div><div class="s">需重点复核</div></div>
</div>
<div class="twocol">
<div><h3>品退 vs 非品退（金额）</h3>{svg_pie(pin_items, size=280)}</div>
<div><h3>品退高频原因 TOP</h3>{svg_hbar(pin_reason_items, color="#C44E52")}</div>
</div>
<h3>品退周趋势（折线，千元）</h3>{svg_line([("品退金额", pintui_week_arr)], [f"W{i+1}" for i in range(WEEKS)], colors=["#C44E52"], ylabel="千元")}
<p class="insight">品退责任方中「商家责任」仅 {pintui_resp.get('商家责任',0)} 单，多数品退未被判商家过错；但「质量问题+描述不符+瑕疵」占品退主体，建议复盘头部商品的<strong>详情页描述准确性与品控</strong>。品退金额最高的商品已在第四节 TOP 榜中体现（品退为其中子集），完整品退清单见附件「品退清单」。</p></section>
"""

HTML = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{SHOP_NAME} 退款单分析报告</title>
<style>
body{{font-family:-apple-system,'Segoe UI','Microsoft YaHei',sans-serif;background:#f5f7fa;color:#222;margin:0}}
.wrap{{max-width:980px;margin:0 auto;padding:24px 20px 60px}}
h1{{font-size:24px;margin:0 0 4px}}
h2{{font-size:19px;margin:28px 0 10px;padding-left:10px;border-left:4px solid #4C72B0}}
h3{{font-size:15px;margin:18px 0 8px;color:#444}}
section{{background:#fff;border-radius:10px;padding:18px 20px;margin:14px 0;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
.note{{background:#eef3fb;color:#335;border-left:3px solid #4C72B0;padding:8px 12px;border-radius:4px;font-size:13px;margin:6px 0 12px}}
.cards{{display:flex;gap:12px;flex-wrap:wrap}}
.card{{flex:1;min-width:160px;background:#f8fafc;border:1px solid #e6ebf2;border-radius:8px;padding:14px}}
.card .k{{font-size:13px;color:#666}}
.card .v{{font-size:28px;font-weight:700;margin:4px 0}}
.card .s{{font-size:12px;color:#888}}
.chartwrap{{display:flex;gap:18px;align-items:center;flex-wrap:wrap;margin:10px 0}}
.legend{{font-size:12px;color:#555;line-height:1.9}}
.lg{{display:inline-block;margin-right:14px;white-space:nowrap}}
.lg i{{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:middle}}
table{{width:100%;border-collapse:collapse;font-size:13px;margin:8px 0}}
th,td{{border:1px solid #e6ebf2;padding:7px 9px;text-align:right}}
th{{background:#f0f4f9;color:#334}}
td:first-child,th:first-child{{text-align:left}}
.twocol{{display:flex;gap:24px;flex-wrap:wrap;align-items:stretch}}
.twocol>div{{flex:1;min-width:300px;display:flex;flex-direction:column}}
.twocol>div>h3{{margin-top:0}}
.twocol svg{{margin-top:4px}}
.chart-hint{{font-size:11px;color:#999;text-align:right;margin-top:4px}}
.insight{{background:#fff7ec;border-left:3px solid #DD8452;padding:10px 14px;border-radius:4px;font-size:13.5px;line-height:1.7;margin-top:12px}}
.sub{{font-size:13px;color:#555;margin:10px 0 4px}}
.rs{{font-size:13.5px;line-height:1.8;color:#333}}
.rs li{{margin:3px 0}}
code{{background:#eef;padding:1px 5px;border-radius:3px;font-size:12px}}
footer{{font-size:12px;color:#aaa;text-align:center;margin-top:24px}}
</style></head><body><div class="wrap">
<h1>{SHOP_NAME} 退款单分析报告</h1>
<div style="color:#888;font-size:13px">{PERIOD_LABEL} · 店铺销售额 {money(TOTAL_SALES)}</div>
{s0}{s1}{s2}{s3}{s4}{s5}
<footer>数据来源：{os.path.basename(A.PATH)} · 口径以本报告"数据概况"为准</footer>
</div></body></html>
"""
with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(HTML)
print("HTML written:", OUT_HTML)

# ---------------- XLSX ----------------
import openpyxl as ox
wb2 = ox.Workbook()
def style_header(ws2):
    for c in ws2[1]:
        c.font = ox.styles.Font(bold=True, color="FFFFFF")
        c.fill = ox.styles.PatternFill("solid", fgColor="4C72B0")
        c.alignment = ox.styles.Alignment(horizontal="center")
    ws2.freeze_panes = "A2"
ws_sum = wb2.active; ws_sum.title = "板块汇总"
ws_sum.append(["板块","归属","单数","退款金额","占退款总额%","板块退款率%","均单","超1000单数","超1000金额","全额退款","部分退款"])
for b in MAIN:
    ws_sum.append([b, "未发货" if b=="未发货退款" else "发货后", b_count[b], round(b_amt[b]),
                   round(b_amt[b]/total_refund*100,2), round(b_amt[b]/TOTAL_SALES*100,2),
                   round(b_amt[b]/max(b_count[b],1)), big_n[b], round(big_a[b]),
                   fp_c[b].get("全额退款",0), fp_c[b].get("部分退款",0)])
ws_sum.append(["其他(换货/补寄/维修)","排除",b_count.get("其他",0),0,0,0,0,0,0,0,0])
ws_sum.append(["退款关闭(不计入)","排除",len(closed),round(closed_amt),0,0,0,0,0,0,0])
style_header(ws_sum)
ws_big = wb2.create_sheet("超1000清单")
cols = ["订单编号","退款编号","订单付款时间","退款完结时间","商品id","宝贝标题","退款总额","退给买家金额","货物状态","售后类型","买家退款原因","板块","是否品退","责任方"]
ws_big.append(cols)
for r in big_rows:
    ws_big.append([r[idx["订单编号"]],r[idx["退款编号"]],r[idx["订单付款时间"]],r[idx["退款完结时间"]],
                   r[idx["商品id"]],r[idx["宝贝标题"]],fnum(r[idx["退款总额"]]),fnum(r[idx["退给买家金额"]]),
                   r[idx["货物状态"]],r[idx["售后类型"]],r[idx["买家退款原因"]],bucket(r),
                   "是" if is_pintui(r) else "否", r[idx["责任方"]] or "未标记"])
style_header(ws_big)
ws_pin = wb2.create_sheet("品退清单")
ws_pin.append(cols)
for r in pintui_rows:
    ws_pin.append([r[idx["订单编号"]],r[idx["退款编号"]],r[idx["订单付款时间"]],r[idx["退款完结时间"]],
                   r[idx["商品id"]],r[idx["宝贝标题"]],fnum(r[idx["退款总额"]]),fnum(r[idx["退给买家金额"]]),
                   r[idx["货物状态"]],r[idx["售后类型"]],r[idx["买家退款原因"]],bucket(r),
                   "是", r[idx["责任方"]] or "未标记"])
style_header(ws_pin)
ws_tp = wb2.create_sheet("TOP商品明细")
ws_tp.append(["排名","商品id","单数","退款金额","均单","品退单数","品退金额","品退占额%","首要退款原因"])
for i,(pid,p) in enumerate(top_prod,1):
    tr = p["reasons"].most_common(1)
    ws_tp.append([i,pid,p["n"],round(p["amt"]),round(p["amt"]/p["n"]),p["pn"],round(p["pa"]),
                  round(p["pa"]/p["amt"]*100,1) if p["amt"] else 0, tr[0][0] if tr else "—"])
style_header(ws_tp)
for sh in wb2.worksheets:
    for col in sh.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        sh.column_dimensions[col[0].column_letter].width = min(max(width+2,10),40)
wb2.save(OUT_XLSX)
print("XLSX written:", OUT_XLSX)
