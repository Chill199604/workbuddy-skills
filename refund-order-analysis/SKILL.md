---
name: refund-order-analysis
description: 电商退款单明细分析 Skill。当用户提供店铺退款单 xlsx（淘宝/天猫退款明细导出）并要求核算退款率、拆解退款板块、分析大额退款/高退款商品/品退时，使用本 Skill。This skill should be used when the user asks to analyze e-commerce refund orders, compute refund rates, break down refund categories (未发货/在途仅退/已发仅退/退货退款), or investigate high-value refunds, top-refund products, and quality-return (品退) orders.
agent_created: true
---

# 退款单数据分析（Refund Order Analysis）

## Overview

把一份店铺退款单明细 xlsx，按统一口径核算成一份可汇报的分析报告（HTML 交互报告 + XLSX 明细表）。
核心能力：未发货退款率 vs 发货后退款率对比、四大板块拆解、超 1000 元大额单、退款金额最高的商品 ID、申请品退订单分析；并用「少条形、多种类」的图表选型提高观感。

## When to use

- 用户提供 `*.xlsx` 退款单/退款明细，并提到「退款率」「退款分析」「退款板块」「品退」「退款金额最高的商品」等。
- 用户是电商运营/管理岗，需要把售后数据做成老板能看懂的汇报。
- 注意：本 Skill 处理**本地文件**，不依赖联网接口；图表为自绘 SVG，离线可看。

## 使用流程

### Step 1 · 确认口径（必须，向用户确认）
先用 `references/schema_and_caliber.md` 的口径决策表，与用户敲定 4 件事：
1. 退款关闭单：排除（推荐）/ 纳入金额。
2. 品退口径：广义（商品问题类，推荐）/ 狭义（仅质量问题）。
3. 商品分析：仅排绝对退款额 / 补销售额算真实单品退款率。
4. 趋势粒度：周 / 月 / 周+月。

### Step 2 · 读取并校验数据结构
读取 xlsx，确认 `references/schema_and_caliber.md` 第 2 节字段都在。若表头不一致，改 `scripts/refund_analysis.py` 顶部 `COLUMN_MAP` 右侧实际表头。

### Step 3 · 配置并运行
修改 `scripts/refund_analysis.py` 顶部 `CONFIG`：
- `PATH`：输入 xlsx 路径
- `TOTAL_SALES`：店铺总销售额（退款率分母）
- `START_DATE`：周期起点（用于按周切片）
- `PIN_KEYS`：品退命中关键词（若选广义）
- `COLUMN_MAP`：字段映射

运行：
```bash
python scripts/refund_analysis.py      # 先跑计算，看 CHECK 摘要数字是否合理
python scripts/refund_render.py        # 再跑渲染，生成 HTML + XLSX
```
> 运行前确保环境有 `openpyxl`（`pip install openpyxl`）。两个脚本都在 `scripts/` 目录，render 会自动 import analysis。

### Step 4 · 出框架 → 定稿 → 分析（参考本次实操）
- 先输出分析框架（模块 + 口径 + 样例数字）让用户过目，再正式生成报告，避免返工。
- 生成后按用户反馈迭代图表（见 `references/chart_selection.md`）：趋势一律折线、占比一律饼图、同类目别堆 3 个以上条形。

## 输出物

- `退款单分析报告.html`：5 大模块，自绘 SVG 图表，离线可看。
- `退款单明细表.xlsx`：板块汇总 / 超1000清单 / 品退清单 / TOP商品明细 四个 sheet。

## 关键口径（默认）

- 仅统计 `退款成功`；`退款关闭` 钱未退，排除出退款率分子，仅单列。
- 换货/补寄/维修金额通常为 0，排除出退款率。
- 首要对比：未发货 vs 发货后（发货后 = 在途仅退 + 已发仅退 + 退货退款）。
- 趋势按 `退款完结时间` 切片，避免历史老订单污染近 90 天窗口。

## Resources

- `scripts/refund_analysis.py`：计算模块（CONFIG 顶部可配，含四大板块/未发货发货后/品退/超1000/TOP商品 统计）。
- `scripts/refund_render.py`：渲染模块（SVG 图表 + HTML + XLSX 输出）。
- `references/schema_and_caliber.md`：字段映射、口径决策表、板块判定逻辑。
- `references/chart_selection.md`：图表选型决策表与观感红线。

## 相关技能（同一套电商售后分析体系）

- **ecommerce-experience-issue-analysis**（体验问题订单分析）：定性·根因·责任穿透视角——回答「这笔退款/投诉是什么性质、根子在谁」。本 Skill 是它的**定量互补**：回答「退款总共掏了多少钱、哪个板块/商品最烧钱、趋势怎么走」。两者在「退款」上是一体两面，合起来就是完整的电商售后/退款分析体系。该 Skill 已把自己描述为 `ecommerce-return-rate` / `ecommerce-returns-profit-recovery` / **本 Skill** 的生态兄弟。
- 后续可补的体系成员（同仓收纳）：退货率分析、客诉分析、评价/VOC 分析、品控归因等，均往 `workbuddy-skills` 仓库以子目录形式追加。
