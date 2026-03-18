# Wind 品种与组合分析看板

一个面向本地 Excel 数据的中文分析看板，适合用于 Wind 导出的期货/化工品价格序列分析。

项目当前支持：

- 读取 Wind 样式的双表头 Excel
- 单品种分析
- 临时价格比 / 价差组合分析
- 预设组合分析
- 风控指标：百分位、ZScore、VaR、ES、波动率、最大回撤
- 季节性分析：历年曲线、均值带、月度分布、同期百分位
- 侧边栏切换 Excel 路径
- Windows 一键启动

## 当前适配的 Excel 格式

当前项目已适配如下结构：

- Sheet 名：`wind_raw_data`
- 第 1 行：简称代码，例如 `PP01`、`LPG01`、`MA05`
- 第 2 行：中文指标说明
- 第 1 列：`Wind`，值为 Excel 序列日期

示例：

| Wind | PP01 | PP02 | LPG01 | MA05 |
|---|---:|---:|---:|---:|
| 46098 | 7771 | 7768 | 5184 | 2847 |

## 主要功能

### 1. 单品种分析

- 支持按品种类别二级选择，例如先选 `PP`，再选 `PP01`
- 查看走势、分布、季节性、风控指标

### 2. 自定义组合

- 支持临时创建价格比或价差组合
- 支持左右腿权重设置
- 不需要先改配置文件

### 3. 风控参数交互

侧边栏可直接设置：

- 百分位窗口
- ZScore 窗口
- VaR 回看窗口
- VaR 期限
- VaR 置信度
- 指定某个输入值，查看其百分位和 ZScore

### 4. 季节性分析

- 季节图已做连续化处理
- 自动补全年日历索引
- 对缺口做线性插值，便于观察连续走势

## 项目结构

```text
trade_dashboard/
├─ config/
│  ├─ app.yaml
│  ├─ metric.yaml
│  └─ strategy.yaml
├─ scripts/
│  ├─ bootstrap_env.py
│  ├─ setup_env.ps1
│  ├─ setup_env.sh
│  ├─ start_dashboard.ps1
│  └─ build_exe.ps1
├─ src/
│  ├─ dashboard.py
│  ├─ data_loader.py
│  ├─ excel_refresh.py
│  ├─ formula_engine.py
│  ├─ portfolio_engine.py
│  ├─ risk_engine.py
│  ├─ seasonal_engine.py
│  ├─ utils.py
│  └─ launcher.py
├─ tests/
├─ requirements.txt
├─ requirements-build.txt
├─ start_dashboard.bat
└─ README.md
```

## 安装依赖

### 方式一：直接安装

```powershell
cd trade_dashboard
python -m pip install -r requirements.txt
```

### 方式二：使用脚本初始化

```powershell
cd trade_dashboard
powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1
```

## 启动方式

### 方式一：直接运行

```powershell
cd trade_dashboard
python -m streamlit run src/dashboard.py --server.headless true --server.port 8501
```

### 方式二：双击启动

直接双击：

- `start_dashboard.bat`

### 方式三：PowerShell 启动

```powershell
cd trade_dashboard
.\scripts\start_dashboard.ps1
```

启动后默认地址：

- [http://localhost:8501](http://localhost:8501)

## Excel 刷新说明

看板中的“刷新 Excel 数据”按钮依赖：

- Windows
- 桌面版 Excel
- `pywin32`

安装：

```powershell
python -m pip install pywin32
```

如果没有这些环境，按钮会提示跳过，但不影响直接读取现有 Excel 文件。

## Excel 文件选择

看板左侧支持两种方式：

- `本地路径`：直接输入或粘贴 Excel 绝对路径
- `拖拽/上传 Excel`：把 Excel 文件拖进侧边栏，或者点击后从本地文件夹选择文件

这意味着你现在既可以手工填路径，也可以直接像普通桌面工具一样选文件。

## 运行测试

```powershell
cd trade_dashboard
python -m pytest -q
```

## 打包为 EXE

### 安装打包依赖

```powershell
cd trade_dashboard
python -m pip install -r requirements-build.txt
```

### 执行打包

```powershell
cd trade_dashboard
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

### 生成位置

打包完成后可在这里找到：

- `dist\WindDashboard\WindDashboard.exe`

注意：

- 真正可迁移的是整个 `dist\WindDashboard` 文件夹，不是只拷贝单独一个 `WindDashboard.exe`
- 这个文件夹里还包含 `_internal` 运行时依赖，缺了它 exe 无法正常工作
- 复制到其他 Windows 电脑时，建议整文件夹一起拷贝
- 目标电脑不需要单独安装 Python

推荐做法：

1. 在你的开发电脑执行打包
2. 把整个 `dist\WindDashboard` 文件夹复制到目标电脑
3. 在目标电脑双击 `WindDashboard.exe`
4. 第一次打开后，在看板左侧选择或拖入目标 Excel 文件

如果你希望跨电脑使用更稳，建议：

- 尽量在和目标电脑相同架构的 Windows 环境下打包
- 打包机和目标机都使用 64 位 Windows
- 如果目标电脑没有桌面版 Excel，就不要依赖“刷新 Excel 数据”按钮，直接读取现成文件即可

## 当前默认预设组合

当前 `config/strategy.yaml` 中包含示例组合：

- `PP01_LPG01_ratio`
- `PP05_MA05_spread`
- `L01_PP01_spread`
- `PP09_L09_ratio`

你也可以继续扩展自己的组合库。

## 后续可继续扩展

- 多腿组合公式编辑器
- 历史 Excel 路径记录
- 自动打开浏览器
- 打包成带图标的桌面应用
- 输出日报或快照
