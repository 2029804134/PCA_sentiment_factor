<div align="center">

# PCA 市场情绪因子与 A 股反转效应研究

**本科论文实证分析公开仓库 · PCA Sentiment Factor for A-share Reversal Effect**

基于市场情绪代理变量、主成分分析和面板回归，构建 `PCAMS` 与 `exPCAMS` 情绪因子，并检验其对 A 股收益反转效应的解释能力。

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)
![PCA](https://img.shields.io/badge/Method-PCA-5B5FC7?style=flat-square)
![Panel Regression](https://img.shields.io/badge/Model-Panel%20Regression-00897B?style=flat-square)
![Data](https://img.shields.io/badge/Data-Restricted-orange?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-blue?style=flat-square)
![Status](https://img.shields.io/badge/Release-Public%20Subset-brightgreen?style=flat-square)

[项目简介](#项目简介) ·
[代码说明](#代码说明) ·
[数据说明](#数据说明) ·
[快速开始](#快速开始) ·
[许可证](#许可证)

</div>

---

## 项目简介

本仓库是论文中“基于 PCA 的市场情绪因子与 A 股反转效应实证分析”部分的公开整理版本。仓库仅保留可公开的分析代码、说明文档和部分公开结果，用于展示从情绪代理变量构建、主成分分析到回归检验的完整研究流程。

> 本项目关注一个核心问题：市场情绪因子是否能够解释或预测 A 股市场中的收益反转效应。

本项目主要支持以下任务：

- 构建市场情绪因子 `PCAMS` 与扩展情绪因子 `exPCAMS`；
- 对沪深 300 样本进行基准回归；
- 对全 A 股样本进行稳健性回归；
- 生成描述性统计、相关图表和统计检验结果；
- 进行异质性回归分析；
- 进行双重中介效应 bootstrap 检验。

## 代码说明

- `code/Get_PCAMS.py`
  对核心情绪代理变量进行标准化处理，使用 PCA 构建 `PCAMS` 与 `exPCAMS`。
- `code/Baseline.py`
  运行沪深 300 样本的基准回归，并导出回归结果表。
- `code/Robust.py`
  运行全 A 股样本的稳健性回归，并导出回归结果表。
- `code/Stats.py`
  生成描述性统计、图表以及相关统计检验结果。
- `code/heterogeneity.py`
  基于分位数网格搜索进行异质性回归分析。
- `code/mediation_bootstrap.py`
  运行双重中介效应 bootstrap 检验。
- `scripts/run_pipeline.py`
  从 PCA 因子构建到后续分析脚本的简化流程入口。

## 数据说明

由于版权和数据许可限制，以下复现实证分析所需的 Excel 数据文件不随公开仓库提供：

- `aggregateData/TS.xlsx`
  `Get_PCAMS.py` 所需的本地输入数据。
- `aggregateData/TS_CSI_300_PCA.xlsx`
  沪深 300 样本的 PCA 与回归数据。
- `aggregateData/TS_CSI_ALL_PCA.xlsx`
  全 A 股样本的 PCA 与回归数据。

如需在本地复现完整流程，请将上述文件放入 `aggregateData/` 目录。公开仓库中仅保留不包含受限源数据的代码、说明文件和可公开结果。若需要相关数据，可通过邮件联系作者：`2029804134h@gmail.com`。

## 仓库结构

```text
.
|-- code/              PCA 市场情绪分析相关脚本
|-- scripts/           可选的一键运行入口
|-- aggregateData/     本地数据与可公开结果目录
|-- docs/              数据布局与开源说明文档
|-- requirements.txt   Python 依赖列表
|-- LICENSE            MIT 开源许可证
`-- README.md          项目说明
```

## 快速开始

安装依赖：

```bash
pip install -r requirements.txt
```

在仓库根目录运行简化流程：

```bash
python scripts/run_pipeline.py
```

也可以按以下顺序分别运行脚本：

1. `python code/Get_PCAMS.py`
2. `python code/Baseline.py`
3. `python code/Robust.py`
4. `python code/Stats.py`
5. `python code/heterogeneity.py`
6. `python code/mediation_bootstrap.py`

## 注意事项

- 本公开版本不包含早期原始数据清洗与构造脚本。
- Word 格式的回归表和中介效应输出默认不纳入版本控制。
- 运行脚本前，请确认本地已准备好必要的数据文件，并且当前工作目录为仓库根目录。

## 许可证

本项目基于 [MIT License](LICENSE) 开源。
MIT·作者：Shaopei Huang