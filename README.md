# daily-merge — 日报 Excel 合并

日报流程第二步：读取 `daily-record` 生成的 md 需求记录，合并到当月 Excel 日报表。

## 解决了什么问题

**md 写了需求，Excel 还要手动抄一遍。** 每天把 md 表格里的内容复制粘贴到 Excel，调格式、改序号、对齐日期——纯体力活。这个脚本一键完成：读 md → 追加到当月 Excel → 冻结表头 → 列宽自适应 → 日期格式匹配。

## 功能清单

- 自动找到当天 md 文件和当月 Excel 基准文件
- 序号全局递增，永不重复
- G 列按 8h/工作日叠加，跳过周六日
- A 列同日期自动合并 + 居中
- D/N 列自动换行，行高自适应
- E 列写百分比（`100%`）
- 仓库名→项目名映射（如 `workingpaper-v5.5` → `中信底稿v5`）
- 数据行和备注之间保留间距，备注原地不动
- 冻结表头行
- 去重：重复跑不会翻倍

## 使用

**每月第一天：**
```bash
python scripts/new_month.py
```

**其他日期：**
```bash
python scripts/daily.py
```

或在 AI 编码助手里说「合并日报」自动判断执行。

## 依赖

```bash
pip install openpyxl
```

## 文件结构

```
daily-merge/
├── SKILL.md           ← 技能指令
├── README.md          ← 本文档
├── templates/
│   └── 日报模板.xlsx   ← 月初用模板
└── scripts/
    ├── new_month.py   ← 月初：模板起新表
    └── daily.py       ← 每日：昨日追加
```

## 列映射

| md 列 | Excel 列 | 说明 |
|--------|----------|------|
| 仓库 | B (项目名称) | 支持映射表转换 |
| 需求概述 | D (任务描述) | — |
| 状态 | E (完成百分比) | 直接写入 |
| 人工工时 | H (预计/实际工时) | — |
| AI辅助工时 | K | — |
| 备注 | N (备注) | — |

## 安装

```bash
git clone https://github.com/huzhw/daily-merge-skill.git ~/.claude/skills/daily-merge
```

安装后在 AI 编码助手里说「合并日报」触发。

## 相关项目

- [daily-record](https://github.com/huzhw/daily-record-skill) — 日报第一步：需求记录 + 工时评估
- [git-commit-skill](https://github.com/huzhw/git-commit-skill) — Git 提交规范，同系列

## 许可

MIT
