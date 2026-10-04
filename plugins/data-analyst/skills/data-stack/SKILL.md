---
name: data-stack
description: 本地数据分析栈的能力契约与坑位清单（VS Code / Python / Excel / Power BI）。当要判断「这台机器到底能不能自动分析 Excel / 能不能自动刷 Power BI」、要选哪条自动化路径（DAX 文本 / Power BI Desktop 本地 XMLA / Power BI REST API / Fabric）、要查该装哪些 Python 库、或要排查 pandas 3.0 / openpyxl / xlwings 的兼容问题时使用。含**分级**的 Power BI 可行性、**待装清单**与**已知冲突**。**注意：本技能不复制官方文档，只登记"现在有什么"+"能到哪一级"+"踩过什么坑"。**
---

# 数据分析栈：能力契约与坑位

> 📌 **定位**：给「搭管道的人」（主代理 / 用户）看的**可行性速查**；运行时方案由 `data-analyst` 子代理消费，它**先读本文件**再决定怎么做。
>
> **核实标注法**：✅ 已核实（本机实测 **或** 打开了官方/开源页面）· ⚠️ 存疑 · ❌ 未核实。
> 库与 API 版本会变，动手前回官方复核。

## 一、本机现状 —— 先自查，再决定方案

> 🚩 **这一节是空的，因为"本机"因机器而异。** 用之前先跑一遍下面的命令，把结果填进表格。
> **别跳过这步** —— `data-analyst` 的一切方案都建立在"这台机器有什么"之上。

```bash
# Python 与关键库
python --version
python -c "import pandas;print('pandas',pandas.__version__)"      # 版本断代很重要，见第四节
python -c "import openpyxl;print('openpyxl',openpyxl.__version__)" # pandas 读写 xlsx 的硬前提
python -c "import matplotlib;print('matplotlib',matplotlib.__version__)"
python -c "import numpy;print('numpy',numpy.__version__)"

# Excel / Power BI 是否存在（Windows）
ls "/c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"
ls -d "/c/Program Files/Microsoft Power BI Desktop" 2>/dev/null || echo "Power BI Desktop 未装"
```

| 组件 | 本机状态（自己填） |
|------|------------------|
| Python | |
| pandas | ← ⚠️ **是不是 3.x？** 关系到第四节 |
| openpyxl | ← ❌ 未装 = **读写 xlsx 会直接失败** |
| matplotlib | ← ❌ 未装 = **出不了图** |
| xlwings / pywin32 | |
| Excel | |
| Power BI Desktop | ← ❌ 未装 = **只能到 T0**（见第二节） |
| git | |

> 🚩 **一句话判据**：**"零依赖"能跑通的，只有 CSV/JSON 的 pandas 分析 + 纯文本 DAX。**
> 凡涉及 Excel 读写、出图、连 Power BI 的，**都要先装东西** —— 见第五节待装清单。

## 二、Power BI：分级可行性（别画饼）

⭐ **记牢这条**：Power BI 的自动化**不是"能不能"的问题，是"到哪一级 + 付什么代价"的问题**。四级，从低到高：

| 级 | 做法 | 前提 | 能做到 | **做不到 / 硬限制** |
|----|------|------|--------|---------------------|
| **T0 · DAX 文本** | Agent 直接写 DAX 表达式给你 | **零前提** ✅ | 度量、计算列、注释、分解说明 | **无法验证它能跑** —— 必须人工贴进 Power BI 验 |
| **T1 · Desktop 本地自动化** | Power BI Desktop 打开时，内嵌一个 **`msmdsrv.exe`（Analysis Services）**，监听**随机回环端口**并暴露 **XMLA 端点**。用 **PowerShell + TOM/ADOMD.NET** 连 `localhost:<port>` | **需装 PBI Desktop**（免费）+ 下 NuGet 包 | 跑 DAX 查询、**改度量/列/关系**、刷新模型、读 DMV 取性能/会话信息 | ⚠️ **不是无人值守** —— Desktop 必须**开着且模型已打开**，关了就没有端点；端口每次启动都变（要动态发现：读 `AnalysisServicesWorkspaces\*\Data\msmdsrv.port.txt`，UTF-16 编码）；**thin report / Direct Lake / 远程 XMLA 不适用**；只覆盖 Import 与 DirectQuery 模型 |
| **T2 · Service REST API** | 调 Power BI REST API（Entra 应用 / 服务主体） | **需 Power BI Service 工作区 + 许可**；服务主体还要**管理员开租户开关**+加进工作区 | **Execute Queries (JSON)** 跑 DAX —— **Pro/PPU/Premium/Fabric 均可** ✅，上限 **10 万行 / 100 万值 / 15MB / 每用户每分钟 120 次**（超限返回 **HTTP 200 + 内嵌 error**，静默截断，别只看状态码）；触发数据集刷新 | 🚩 **服务主体刷新在非 Premium 工作区实测 403**（需 Premium/PPU/Fabric）；🚩 **push dataset 不支持服务主体**（403 `API is not accessible for application`）；🚩 **无法用 API 从零建报表** —— 只能**克隆已有报表**或**导入 PBIX**；`Execute Queries (Arrow)` 是**另一条**，**强制要 Premium/Fabric/Embedded 容量** |
| **T3 · Fabric 全自动** | Fabric notebook / 数据管道 / XMLA **读写**端点 | **需 Fabric 或 Premium 容量（付费）** | 真·无人值守的 ETL + 语义模型管理 | 个人自用**成本不划算**，别默认选它 |

**选级顺序**：先看有没有 Service 账号（有 → T2；没有 → 只能 T0/T1）→ **T0 永远能做**（不装任何东西）→ 想本地跑 DAX 才装 Desktop 上 T1 → **T3 别碰**，除非本来就有 Fabric 容量。

**参考实现（已核实存在，未跑）**：
- `data-goblin/power-bi-agentic-development`（GitHub）—— 内含 `connect-pbid` skill，脚本 `query-dax.ps1`（发现端口 → 装 ADOMD.NET → 连 `localhost:<port>` → 跑 DAX / DMV）
- `pbi-tools`（`github.com/pbi-tools/pbi-tools`）—— **源码管理 / 抽取 PBIX** 的 CLI，`extract` 需 **Desktop 版**；无 Desktop 的 Core 版**不支持 extract**

> ⚠️ **别信"用 GUI 自动化（pyautogui 点鼠标）操作 Power BI Desktop"** —— 脆、版本一变就废，**不作为方案**。

## 三、Excel：两条路，别选错

| 需求 | 库 | 前提 |
|------|----|------|
| **读写 xlsx 数据**（pandas 的默认引擎） | `openpyxl` **>= 3.1.5** | 纯 Python，**无需 Excel** |
| **写带格式的报表**（颜色 / 边框 / 条件格式 / 图表），不需要 Excel 进程 | `xlsxwriter` **>= 3.2.0** | 纯 Python |
| **驱动真实 Excel**（透视表、公式重算、宏、`Workbook.active()` 连正在开的文件） | `xlwings` | **必须有 Excel** + `pywin32`（Windows 的 COM 桥） |

> 🚩 **`xlwings` 和 `openpyxl` 的根本差别**：`xlwings` **依赖本机 Excel 进程**（走 COM / pywin32），`openpyxl` **不依赖**。所以：
> - **要无头、要快、要跨环境** → `openpyxl` / `xlsxwriter`
> - **要真·透视表 / 公式引擎算出值 / 操作你正开着的表** → `xlwings`（代价：慢、要装 pywin32、会起 Excel 进程）
> - 🚩 **OneDrive 路径 + Excel 进程**：若项目在 OneDrive 下，`xlwings` 打开/保存会触发同步，**批量写入时先用本地临时目录再落回**。

## 四、Python 数据栈：pandas **3.0** 的断代坑

> 若你的 pandas 是 **3.x**，**网上绝大多数教程是 1.x/2.x 写的，代码可能直接错**。已知断代差异：

1. 🚩 **默认字符串 dtype 变了**：3.0 起 `read_csv` 等**直接推断出专用 `"str"` 扩展 dtype**，**不再是 NumPy 的 `object`**。
   - → `df.dtypes` 里看到的是 `str`；`ser.dtype == "object"` **会返回 False**。
   - → 跨版本写法：用 `pd.api.types.is_string_dtype(ser)` 或 `select_dtypes(include=["object","string"])`。
   - → 该 dtype **只存字符串**，塞非字符串会**报错**（旧 `object` 不会）。
2. 🚩 **Copy-on-Write 默认开启**：链式赋值（`df[df.a>1]["b"] = 0`）**不再生效**，`SettingWithCopyWarning` 也没了（因为不再有"视图"语义）。改写要用 `.loc[...] = value` 或重新赋值。
3. **缺失值哨兵统一为 `NaN`**：字符串列里原来的 `None` 会被转成 `NaN` —— 一律用 `pd.isna()` 判空。
4. **`pytz` 不再是必需依赖**；`pyarrow` 仍是**可选**（不影响 csv/xlsx）。
5. ⚠️ **`matplotlib` 的 pandas plot extra 要求 >= 3.9.3**；**`openpyxl` 的 excel extra 要求 >= 3.1.5** —— 装低了 pandas 的 `read_excel` 会拒绝。
6. 🚩🚩 **`ydata-profiling` 与 pandas 3.x 冲突**：它依赖域写死 **`pandas < 3.0`** → **在 pandas 3.0 上装不上**。想用"自动化 EDA（ydata-profiling）"的，二选一：
   - **(a)** 单开一个 venv 降到 `pandas 2.x` 专门跑 profiling（**牺牲**新版 pandas 特性）
   - **(b)** 放弃 `ydata-profiling`，**手写 EDA**（`.describe()` + `corr()` + `matplotlib` + 缺失矩阵）—— **推荐，零冲突**

## 五、待装清单（缺什么装什么，**装前先报**）

> 🔒 **安装纪律**：装之前先报「装什么 / 为什么 / 影响」→ 用户确认 → **装进 `{分析根}/.venv`**（别污染全局）→ 同步 `requirements.txt`。
> 基础命令（Windows 用 `Scripts`，不是 `bin`）：
> ```bash
> python -m venv "{分析根}/.venv"
> "{分析根}/.venv/Scripts/python.exe" -m pip install openpyxl matplotlib xlsxwriter
> ```

| 必装（Phase 1 就有用） | 版本 | 用途 |
|---|---|---|
| `openpyxl` | >= 3.1.5 | **pandas 读写 xlsx 的硬前提**，没有它 `.xlsx` 读不了 |
| `matplotlib` | >= 3.9.3 | 图表落盘 |
| `xlsxwriter` | >= 3.2.0 | 生成带格式的 Excel 报表 |

| 按需 | 版本 | 用途 / 备注 |
|---|---|---|
| `xlwings` + `pywin32` | latest | 驱动**真实 Excel**（透视表/公式/宏）。前提：装有 Excel |
| `scipy` | >= 1.14.1 | 统计检验（t / 卡方 / ANOVA） |
| `pyarrow` | >= 13 | 提速字符串运算、读写 parquet。可选 |
| `ydata-profiling` | — | ⚠️ **与 pandas 3.x 冲突**，见第四节第 6 条 |

## 六、本文件里"未核实"的（诚实标注）

- ❌ **T1（PBI Desktop 本地 XMLA）机制与脚本均来自已核实的官方/开源来源，但"装完就能跑"这一步未在每台机器验证过** —— 装完 Desktop 后**跑一次冒烟测试**再当它可用。
- ⚠️ **T2 的授权细节会变**：Microsoft 的租户开关与容量要求调整频繁，**动手前回 `learn.microsoft.com/power-bi` 复核**。
- ⚠️ **pandas 3.x 较新，第三方库（尤其 profiling / 可视化类）适配进度参差** —— 遇到第三方库报错，**先怀疑 pandas 3.x 兼容性**，再查别处。
