# 中文多品种组合分析与风控看板（Excel 驱动）

本项目用于在本地 Python 环境中：
- 从 Excel 自动刷新并读取多列时间序列；
- 按可配置公式动态生成组合；
- 输出风控指标与季节分析；
- 通过 Streamlit 提供中文交易看板。

## 1）项目目录结构

```text
trade_dashboard/
├─ config/
│  ├─ app.yaml
│  ├─ strategy.yaml
│  └─ metric.yaml
├─ data/
│  ├─ raw/
│  │  └─ wind_data.xlsx        # 你的 Wind 导出 Excel
│  └─ processed/
├─ outputs/
├─ src/
│  ├─ excel_refresh.py         # Excel RefreshAll
│  ├─ data_loader.py           # Excel读取与清洗
│  ├─ formula_engine.py        # 公式解析与计算
│  ├─ portfolio_engine.py      # 批量组合生成
│  ├─ risk_engine.py           # 风控指标
│  ├─ seasonal_engine.py       # 季节模块
│  ├─ dashboard.py             # Streamlit看板
│  └─ utils.py                 # logging/yaml工具
├─ tests/
├─ requirements.txt
└─ README.md
```

## 2）Excel 模板建议

### Sheet: `数据`
| Date | PP | PG | LPG | Brent | 甲醇 | USD/CNY | 运费 | 利润指标 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|

要求：
- 第一列固定为 `Date`；
- 列名唯一，不合并单元格；
- 每列为时间序列，行按日期递增。

### Sheet: `策略`
| StrategyName | Formula | Enabled | Category | Notes |
|---|---|---|---|---|
| PP_PG_ratio | PP / PG | Y | 烯烃 | 基础价比 |
| PP_PG_adj | PP / (1.8 * PG) | Y | 烯烃 | 调整价比 |
| PP_PG_spread | 17 * PP - 5 * PG | Y | 套利 | 手数组合 |
| Brent_Naphtha | Brent - Naphtha | Y | 能化 | 跨品种 |

公式规则：
- 可使用任意数据列名；
- 支持加减乘除、括号、常数；
- 推荐格式：`17 * PP - 5 * PG`。

## 3）模块职责

- `excel_refresh.py`：在 Windows + Excel 下调用 `win32com` 执行 `RefreshAll`、等待、保存、关闭。
- `data_loader.py`：读取数据/策略 sheet，日期规范化、数值清洗、字段校验。
- `formula_engine.py`：解析公式并按列运算，支持动态组合。
- `portfolio_engine.py`：遍历策略配置，输出多组合 DataFrame。
- `risk_engine.py`：计算当前值、历史分位、Z-score、VaR/ES、波动率、最大回撤、rolling corr。
- `seasonal_engine.py`：季节矩阵、季节均值带、月度分布、同期分位与偏离、去除2月29日。
- `dashboard.py`：中文页面（总览、组合分析、季节分析、数据浏览）。


## 环境一键配置（推荐）

### Linux / macOS
```bash
cd trade_dashboard
bash scripts/setup_env.sh
```

### Windows PowerShell
```powershell
cd trade_dashboard
powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1
```

脚本会自动：
- 创建虚拟环境 `.venv`
- 升级 pip
- 检测缺失依赖
- 按多个镜像顺序安装依赖（PyPI / 清华 / 阿里云）

若公司网络需要代理，可先配置：
```bash
python -m pip config set global.proxy http://<user>:<pass>@<proxy_host>:<port>
```

## 快速开始

```bash
cd trade_dashboard
python -m pip install -r requirements.txt
streamlit run src/dashboard.py
```

## Windows Excel 自动刷新

1. 安装桌面版 Excel；
2. 安装 pywin32：
   ```bash
   pip install pywin32
   ```
3. 在看板点击“刷新 Excel 数据”。

## 测试

```bash
cd trade_dashboard
pytest -q
```
