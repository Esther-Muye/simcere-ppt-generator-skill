# 先声药业 PPT 生成器 v3.0

基于先声药业 2026 版官方 PPT 模板，通过 JSON 配置自动生成符合集团统一视觉风格的可编辑 PPT 文件。

> **v3.0 更新**：精确排版引擎（25色 + 18级字号 + 段落间距系统），融合官方模板与3份实战参考PPT风格。

## 核心功能

### 16 种页面类型

| 页面类型 | 说明 | 典型场景 |
|----------|------|----------|
| `cover` | 封面页（44pt大标题） | PPT 首页 |
| `toc` | 目录页（18pt条目） | 汇报大纲 |
| `section` | 章节分隔页（32pt标题） | 章节过渡 |
| `content` | 标准内容页（26pt标题 + 14pt正文 + 可选表格/图表） | 大多数内容页 |
| `gantt` | 矩阵式甘特图 | 项目推进节奏、月度计划 |
| `timeline_horizontal` | 横向卡片式时间线 | Campaign 规划、传播节奏 |
| `big_number` | 超大字号核心数字 + 说明 + 明细卡片 | 活动规模、成本拆解 |
| `comparison` | 多组数据并排对比卡片 | 竞品对比、方案对比 |
| `process` | 水平排列的圆角矩形节点 + 箭头 | 里程碑、上市节奏 |
| `kpi_dashboard` | 网格排列的指标卡片 | 关键指标总览 |
| `item_matrix` | 物品矩阵图（场景→对象→物品→成本→评分） | 品牌提示物盘点、资源矩阵 |
| `calendar_grid` | 日历网格（行=场景，列=月份） | 月度推广排期、资源分配 |
| `review_matrix` | 复盘矩阵（有效动作 × 做的好/不好） | 项目复盘 |
| `action_category` | 分类复盘（持续做/优化提升/开始做） | 改进计划 |
| `strategy_diagram` | 策略架构图（中心+维度+支柱） | 策略展示 |
| `ending` | 结束页（44pt标题） | "谢谢" 尾页 |

### 设计规范（v3.0 精确版）

- **颜色系统**：25色（含主题色6种 + 业务色 + 背景色 + 架构图色）
- **字体**：中文微软雅黑（`a:ea`）+ 英文 Arial（`a:latin`），同时设置
- **字号**：18级精确层级（44pt → 32pt → 26pt → 20pt → 16pt → 14pt → ... → 7pt）
- **段落间距**：标题 0/8pt、正文 6/3pt、紧凑 2/1pt，行距 1.0x/1.3x/1.15x
- **表格**：真实表格对象 + 先声绿表头(#00B052) + 隔行浅绿(#F0F7F4)
- **图表**：真实图表对象，accent1→accent6 配色顺序
- **模板继承**：完整保留母版、版式、Logo、页脚、版权信息

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
  --input your_content.json \
  --output result.pptx
```

## 目录结构

```
.
├── SKILL.md                        # Skill 完整定义（含9大业务模块 + 设计规范）
├── README.md
├── assets/
│   └── simcere-template.pptx      # 先声药业2026官方PPT模板
├── scripts/
│   ├── generate.py                # 主生成脚本（Python + python-pptx，16种页面构建器）
│   └── simcere_layout.py          # 排版引擎（25色 + 18级字号 + 段落间距系统）
├── references/
│   └── template-spec.md           # 精确模板视觉规范（占位符坐标、颜色、字号、间距）
├── examples/
│   ├── content-example.json       # 内容JSON示例（基础版）
│   └── content-example-v2.json    # 内容JSON示例（14页科唯可复盘完整版）
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
      "tag": "市场洞察",
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
      "type": "review_matrix",
      "title": "有效动作复盘矩阵",
      "badge": "复盘",
      "category_header": "有效动作",
      "good_header": "做的好（持续做）",
      "bad_header": "做的不好（优化提升）",
      "rows": [
        {"action": "宣传/传播", "good": "321睡眠日整合营销...", "bad": "线上仅135场(6.4%)..."},
        {"action": "覆盖", "good": "新开发连锁69家...", "bad": "县域仅完成62%..."}
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

- 市场部汇报（业务复盘、营销计划、品牌策略）
- 零售业务策略规划（渠道、产品、品牌提示物）
- 医药连锁项目复盘（五维链路：传播→覆盖→陈列→培训→复购）
- 产品上市计划（竞品对比、差异化定位、关键里程碑）
- 数据汇报与洞察（KPI仪表盘、趋势分析、投入产出分析）
- 组织能力建设（培训体系、能力评估、考核激励）

## 技术栈

- Python 3.11+
- [python-pptx](https://python-pptx.readthedocs.io/)
- lxml（段落间距 XML 操作）

## 注意事项

- 模板文件 `simcere-template.pptx` 为 16:9 宽屏尺寸（13.333" × 7.5"）
- 如用户未提供具体数据，JSON 中可使用 `"[待补充]"` 占位，**绝不编造数据**
- 所有生成内容均为可编辑对象（文本框、表格、图表），非整页截图
- 完整保留模板母版、Logo、页脚版权信息
- 字体通过 `a:ea`（东亚）和 `a:latin`（拉丁）分别设置，确保中英文混排正确
- 表格为真实 PPT 表格对象，非色块拼贴

---

**Copyright (C) 2026 Simcere. All rights reserved.**
