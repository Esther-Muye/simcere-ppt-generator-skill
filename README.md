#  PPT 生成器

基于simcere 2026 版官方 PPT 模板，通过 JSON 配置自动生成符合集团统一视觉风格的可编辑 PPT 文件。

## 核心功能

### 14 种页面类型

| 页面类型 | 说明 | 典型场景 |
|----------|------|----------|
| `cover` | 封面页，含标题、副标题、日期、汇报人 | PPT 首页 |
| `toc` | 目录页，列出章节结构 | 汇报大纲 |
| `section` | 章节分隔页 | 章节过渡 |
| `content` | 标准内容页，含标题 + 正文 + 可选表格/图表 | 大多数内容页 |
| `table` | 独立数据表格页 | 数据对比、进度跟踪 |
| `chart` | 图表页（柱状图/折线图/饼图/面积图） | 趋势、占比、对比 |
| `gantt` | 矩阵式甘特图 | 项目推进节奏、月度计划 |
| `timeline_horizontal` | 横向卡片式时间线 | Campaign 规划、传播节奏 |
| `big_number` | 超大字号核心数字 + 说明 + 底部明细卡片 | 活动规模、成本拆解 |
| `comparison` | 多组数据并排对比卡片 | 竞品对比、方案对比 |
| `process` | 水平排列的圆角矩形节点 + 箭头 | 里程碑、上市节奏 |
| `kpi_dashboard` | 网格排列的指标卡片 | 关键指标总览 |
| `item_matrix` | 物品矩阵图（场景→对象→物品→成本→评分） | 品牌提示物盘点、资源矩阵 |
| `calendar_grid` | 日历网格（行=场景，列=月份） | 月度推广排期、资源分配 |
| `ending` | 结束页 | "谢谢" 尾页 |

### 设计规范

- **主题色**：绿色 `#00B052` 为主色，深绿 `#006647` 为辅助
- **字体**：中文微软雅黑，英文 Arial
- **模板继承**：完整保留母版、版式、Logo、页脚、版权信息
- **可编辑性**：所有文本框、表格、图表均可直接编辑

## 快速开始

### 安装依赖

```bash
pip install python-pptx
```

### 使用方式

1. 准备内容 JSON 文件（参考 `examples/content-example.json`）
2. 运行生成脚本：

```bash
python scripts/generate.py \
  --template assets/simcere-template.pptx \
  --input examples/content-example.json \
  --output output.pptx
```

## 目录结构

```
.
├── assets/
│   └── simcere-template.pptx      # 先声药业官方PPT模板
├── scripts/
│   └── generate.py                # 主生成脚本（Python + python-pptx）
├── references/
│   └── template-spec.md           # 模板视觉规范文档
├── examples/
│   └── content-example.json       # 内容JSON示例
└── README.md
```

## JSON 内容格式示例

```json
{
  "title": "PPT总标题",
  "slides": [
    {
      "type": "cover",
      "title": "封面主标题",
      "subtitle": "封面副标题",
      "date": "2026年6月",
      "presenter": "汇报部门/人"
    },
    {
      "type": "content",
      "title": "市场现状分析",
      "subtitle": "2026年上半年零售市场核心数据",
      "body": [
        {"text": "市场规模持续增长，但增速放缓", "level": 0},
        {"text": "竞品加速布局县域市场", "level": 0}
      ],
      "table": {
        "headers": ["指标", "Q1", "Q2", "同比增长"],
        "rows": [
          ["销售额(亿元)", "12.5", "14.2", "+13.6%"]
        ]
      }
    },
    {
      "type": "gantt",
      "title": "下半年推进甘特图",
      "rows": ["宣传", "覆盖", "陈列", "培训", "复购"],
      "months": ["7月", "8月", "9月", "10月", "11月", "12月"],
      "data": [
        {"row": "覆盖", "month": "7月", "text": "方案落地", "highlight": true}
      ]
    },
    {
      "type": "item_matrix",
      "title": "品牌提示物使用情况概览",
      "groups": [
        {
          "scene": "品牌活动",
          "audience": "患者",
          "cost_tier": "≤10元",
          "avg_score": "8.5",
          "items": [
            {"name": "定制抽纸", "score": "8.5"},
            {"name": "抽绳背包", "score": "8.5"}
          ]
        }
      ]
    },
    {
      "type": "calendar_grid",
      "title": "月度推广排期",
      "rows": [
        {"name": "品牌活动", "subtitle": "（对象：患者）"}
      ],
      "months": ["7月", "8月", "9月", "10月", "11月", "12月"],
      "grid": [
        [
          {"items": ["抽绳背包", "艾草锤"], "highlight": true},
          {"items": ["定制抽纸"]}
        ]
      ]
    },
    {
      "type": "ending",
      "title": "谢谢",
      "subtitle": "Thank You"
    }
  ]
}
```

## 适用场景

- 市场部汇报
- 零售业务策略规划
- 医药连锁项目复盘
- 产品上市计划
- 领导汇报材料

## 技术栈

- Python 3
- [python-pptx](https://python-pptx.readthedocs.io/)

## 注意事项

- 模板文件 `simcere-template.pptx` 为 16:9 宽屏尺寸（13.333" x 7.5"）
- 如用户未提供具体数据，JSON 中可使用 `"待补充"` 占位
- 所有生成内容均为可编辑对象，非整页截图
- 完整保留模板母版、Logo、页脚版权信息

---

**Copyright (C) 2026 Simcere. All rights reserved.**
