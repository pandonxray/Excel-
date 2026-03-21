# 交易研究看板

一个面向日度时间序列研究的中文看板，适合对 Wind 导出的商品、化工、能源相关 Excel 数据做研究分析。

当前版本重点支持：

- 单品种分析
- 月差分析
- 跨表/价差组合分析
- 预设组合分析
- 驱动拆解
- 风控指标分析
- 季节性分析
- 本地 Excel 刷新与上传

## 适用场景

适合以下类型的研究工作：

- 跟踪单个品种的历史位置、波动和分布
- 研究月差、跨品种价差、成本端价差
- 观察预设组合的驱动来源
- 做日度层面的风险评估和季节性复盘

本项目默认只依赖日度 Excel 数据，不依赖实时行情。

## 当前功能

### 1. 单品种

- 按数据源和品种分组选择序列
- 查看时间序列、分布、风控指标、季节性

### 2. 月差

- 新增 `月差` 作为分析对象
- 支持先选品种，再选近月和远月
- 例如：
  `L05 - L09`
  `L09 - L01`

### 3. 跨表/价差组合

- 支持价差组合
- 支持比值组合
- 支持多腿自定义组合
- 支持乘数项和系数设置

### 4. 预设组合

- 从 `config/strategy.yaml` 加载策略
- 支持分类查看
- 当前已扩展：
  `LPG内外价差`
  `FEI_PDH`
  `PP-L`
  `MTO`

### 5. 驱动拆解

新增通用驱动拆解框架，支持：

- 组件序列拆解
- 派生项展示
- 驱动路径标准化
- 区间贡献拆解
- 驱动诊断
- 敏感度分析
- 情景分析

其中：

- `驱动路径标准化` 支持用户自选起点
- `贡献拆解` 支持用户自选分析区间
- 页面内提供术语说明和计算逻辑说明

### 6. 风控分析

支持常用研究指标：

- 历史分位
- Z-Score
- VaR
- ES
- 波动率
- 最大回撤
- 自定义目标值的定位分析

### 7. 季节性分析

支持：

- 历年季节路径
- 季节均值与波动带
- 月度分布箱线图
- 季节性分位与季节性偏离

当前版本新增：

- 季节图纵轴上下限可手动输入

### 8. Excel 输入方式

左侧栏支持两种方式：

- 本地路径
- 拖拽/上传 Excel

### 9. Excel 刷新

支持在 Windows 本机环境下调用 Excel 刷新工作簿。

## Excel 数据要求

当前默认配置见 [config/app.yaml](/D:/codex/Excel-/trade_dashboard/config/app.yaml)。

默认读取：

- `wind_raw_data`
- `manual_data`

默认字段：

- 日期列：`Wind`
- 手工数据日期列：`price_date`
- 汇率列：`USDCHY`

典型结构示意：

| Wind | PP01 | L01 | LPG01 | FEI01 | USDCHY |
|---|---:|---:|---:|---:|---:|
| 46098 | 7771 | 8120 | 4387 | 615 | 7.23 |

## 项目结构

```text
trade_dashboard/
├─ config/
│  ├─ app.yaml
│  ├─ metric.yaml
│  └─ strategy.yaml
├─ scripts/
│  ├─ build_exe.ps1
│  ├─ setup_env.ps1
│  └─ start_dashboard.ps1
├─ src/
│  ├─ dashboard.py
│  ├─ data_loader.py
│  ├─ driver_engine.py
│  ├─ excel_refresh.py
│  ├─ formula_engine.py
│  ├─ portfolio_engine.py
│  ├─ risk_engine.py
│  ├─ seasonal_engine.py
│  └─ utils.py
├─ tests/
│  ├─ test_data_loader.py
│  ├─ test_driver_engine.py
│  ├─ test_formula_engine.py
│  ├─ test_risk_engine.py
│  └─ test_seasonal_engine.py
├─ requirements.txt
├─ requirements-build.txt
├─ start_dashboard.bat
└─ README.md
```

## 安装

### 方式一：直接安装

```powershell
cd trade_dashboard
python -m pip install -r requirements.txt
```

### 方式二：使用项目虚拟环境

```powershell
cd trade_dashboard
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

说明：

- 当前依赖已限制 `numpy<2`
- 启动脚本会优先使用项目下的 `.venv`

## 启动

### 方式一：双击启动

直接运行：

- `start_dashboard.bat`

### 方式二：PowerShell 启动

```powershell
cd trade_dashboard
.\scripts\start_dashboard.ps1
```

### 方式三：命令行启动

```powershell
cd trade_dashboard
.\.venv\Scripts\python.exe -m streamlit run src/dashboard.py --server.headless true --server.port 8501
```

启动后默认地址：

- [http://localhost:8501](http://localhost:8501)

## README 对应的当前更新点

本次版本包含这些重点更新：

- 修复启动时误用系统 Python/Anaconda 的问题
- 增加 `LPG内外价差` 预设组合
- 新增通用驱动拆解引擎
- 驱动拆解页增加图表、说明、区间分析和标准化起点控制
- 新增 `月差` 分析对象
- 季节图支持纵轴范围手动输入
- 增加术语说明与计算逻辑说明

## 驱动拆解说明

驱动拆解页主要回答三个问题：

1. 当前组合由哪些底层变量构成
2. 某一段时间内是谁推动了变化
3. 如果某个驱动继续波动，目标序列可能如何变化

当前框架支持：

- 自动识别简单价差结构
- 配置化拆解复杂组合
- 日期切换系数
- 区间归因
- 敏感度与情景分析

驱动拆解的核心代码在：

- [src/driver_engine.py](/D:/codex/Excel-/trade_dashboard/src/driver_engine.py)

## 测试

运行全部测试：

```powershell
cd trade_dashboard
.\.venv\Scripts\python.exe -m pytest --rootdir . tests
```

如果只想跑驱动拆解相关测试：

```powershell
cd trade_dashboard
.\.venv\Scripts\python.exe -m pytest --rootdir . tests/test_driver_engine.py
```

## 打包 EXE

安装打包依赖：

```powershell
cd trade_dashboard
python -m pip install -r requirements-build.txt
```

执行打包：

```powershell
cd trade_dashboard
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

## 注意事项

- 当前 `config/app.yaml` 中的 Excel 路径是本地默认路径，跨机器使用时建议改成自己的路径
- 如果要使用 Excel 刷新功能，需要本机具备桌面版 Excel 和对应环境
- `.venv` 已通过 `.gitignore` 排除，不会被提交到仓库

## 后续建议

后面还可以继续增强：

- 月差曲线总览
- 区间内自动结论摘要
- 研究结论卡片
- 驱动领先/滞后分析
- 结构断点识别
