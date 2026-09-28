# workbuddy-skills

WorkBuddy 可复用技能（Skill）收纳仓库，定位 **电商售后 / 数据分析** 方向。

每个子目录是一个自包含的 Skill（含 `SKILL.md` + `scripts/` + `references/`），互不干扰，可单独拷贝到 `~/.workbuddy/skills/` 使用。

## 已收录

| 目录 | 技能 | 视角 | 产出 |
|---|---|---|---|
| `refund-order-analysis/` | 退款单数据分析 | **定量**：退款率、四大板块、大额单、TOP 商品、品退的金额/比例/趋势 | HTML 交互报告 + XLSX 明细表 |
| `ecommerce-experience-issue-analysis/` | 消费者体验问题订单分析 | **定性**：体验问题的性质与根因责任穿透 | HTML 七章报告（在各自仓库） |

> 两者构成一套「电商售后/退款分析体系」：一个管"花了多少钱"（定量），一个管"根子在谁"（定性）。在退款这件事上是一体两面。

## 后续规划（同仓追加子目录）

- 退货率分析（比例与趋势）
- 客诉 / 评价 / VOC 分析
- 品控归因
- 利润修复

## 目录约定

```
<repo>/
├── README.md                 # 本文件
└── <skill-name>/             # 单个 skill，自包含
    ├── SKILL.md              # 元信息 + 流程指引
    ├── scripts/              # 可执行脚本
    └── references/           # 字段映射、口径、方法论
```

## 使用

把需要的 `<skill-name>/` 目录整体拷贝到 WorkBuddy 的 `~/.workbuddy/skills/` 即可被识别加载。
