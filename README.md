# data-analyst

> 一个只做**本地数据计算**的 Claude Code 子代理：读 Excel/CSV → 生成可重跑的 Python 脚本 → 执行 → 产出清洗后数据、图表与报告。
> 外加一个 `data-stack` 技能，把「这台机器到底能做什么」讲清楚 —— **包括做不到的部分**。

---

## 为什么会有这个

把「一堆表 → 结论」这条链路，从手工敲 pandas 变成可复用流水线。核心设计取向：

- **独立上下文 = 隔离脏输出。** 数据清洗/EDA 会产生大量中间打印和报错回溯 —— 这些留在子代理里，主对话只拿到结论。
- **脚本化、可重跑。** 清洗与统计必须是脚本产物，不是"在对话里手动改几行"。同一脚本对同一原始数据必须给出同一结果。
- **不画饼。** Power BI 的自动化被授权层级切得很碎，能力分级写进技能里，**每一级都标清楚前提与硬限制**（见下）。
- **缺依赖先报后装。** 不擅自 `pip install`；缺库时报告「要装什么 / 为什么 / 影响」，装也装进项目 `.venv`，不污染全局。

## 装

作为 Claude Code 插件：

```bash
claude plugin marketplace add <你的用户名>/data-analyst
claude plugin install data-analyst@data-analyst
```

或直接把 `plugins/data-analyst/agents/data-analyst.md` 拷进你项目的 `.claude/agents/`。

## 用

对 Claude 说：

- 「分析一下 `sales.xlsx`，清理缺失的订单数据，按地区汇总」
- 「做个 EDA，找找影响转化的关键因素」
- 「把结果整理成带格式的 Excel 报表」
- 「把这句话翻成 DAX：计算去年同期的销售额增长率」

四种模式：`clean`（清洗 + 质量报告）· `eda`（探索性分析）· `excel`（报表自动化）· `dax`（自然语言 → DAX）。

写区默认只有两处：`{分析根}/scripts/` 与 `{分析根}/output/`。输入数据（`data/`）**只读，绝不被覆盖**。

## 结构

```
.
├── .claude-plugin/marketplace.json
├── plugins/data-analyst/
│   ├── .claude-plugin/plugin.json
│   ├── agents/data-analyst.md          ← 子代理定义（含四种模式与硬性护栏）
│   └── skills/data-stack/SKILL.md      ← 可行性速查（Power BI 四级、Excel 两条路、pandas 3.0 断代坑）
├── examples/
│   └── xlsx_read.py                    ← 零依赖 xlsx 读取器（纯标准库）
└── README.md
```

## 一个"零依赖"的例子

`examples/xlsx_read.py` 用 **Python 标准库**（`zipfile` + `xml.etree`）直解 xlsx —— 因为 xlsx 本质就是 zip 里的 XML。

在**没装** `openpyxl`（pandas 读写 xlsx 的硬前提）的机器上，这就是"不擅自装依赖"却仍要拿到结果的落地方案。首个实战（统计 Excel 导出里的每日完成量）就是靠它跑通的。

## 它做不到什么（诚实清单）

- **Power BI 的"全自动"分四级**（T0 DAX 文本 → T1 Desktop 本地 XMLA → T2 Service REST API → T3 Fabric），**没有一级是零前提的**。细节与硬限制见技能 `data-stack`。
- **无法用 API 从零建 Power BI 报表** —— REST API 只能克隆已有报表或导入 PBIX。
- **它不"连接" VS Code** —— 它本来就跑在同一个 Claude Code 进程里。子代理有独立上下文，**拿不到你正在编辑的选区**；要看某个文件，得把路径给它（或由主代理贴进派单）。

## 许可

MIT
