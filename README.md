# 公路桥梁技术状况评定系统

Highway Bridge Condition Assessment System

基于 Python 和 PySide6 开发的桌面应用，支持公路桥梁病害数据录入、构件评分、结构及全桥技术状况评定、项目保存和 Excel 报告导出。

## 主要功能

- 支持梁桥、拱桥和缆索承重桥，共 10 类桥梁。
- 按桥面系、上部结构、下部结构组织评定数据。
- 管理构件数量、病害信息、构件是否存在及评定权重。
- 从 Excel 模板导入病害数据，并导出评定数据和评分报告。
- 保存和加载 JSON 项目文件，查询历史项目。
- 提供图形界面、计算过程展示及内置帮助文档。

## 支持的桥梁类型

| 类别 | 桥梁类型 | 模块 |
| --- | --- | --- |
| 梁桥 | 混凝土梁桥 | `Concrete_beam_bridge.py` |
| 梁桥 | 钢结构梁桥 | `Steel_beam_bridge.py` |
| 拱桥 | 圬工拱桥 | `Masonry_arch_bridge.py` |
| 拱桥 | 钢筋混凝土拱桥 | `Reinforced_concrete_arch_bridge.py` |
| 拱桥 | 双曲拱桥 | `Double_curved_arch_bridge.py` |
| 拱桥 | 刚架拱桥 | `Rigid-framed_arch_bridge.py` |
| 拱桥 | 桁架拱桥 | `Trussed_arch_bridge.py` |
| 拱桥 | 钢-混组合拱桥 | `Steel-concrete_composite_arch_bridge.py` |
| 缆索承重桥 | 斜拉桥 | `Cable-stayed_bridge.py` |
| 缆索承重桥 | 悬索桥 | `Suspension_bridge.py` |

## 安装与运行

下面以 Windows 和 Python 3.12 为例。需要可显示桌面窗口的运行环境。

```powershell
git clone https://github.com/ZXLCQU/Highway_bridge_assessment.git
cd Highway_bridge_assessment
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe highway_bridge_assessment_system.py
```

请从项目根目录启动程序，桥梁模块使用相对路径读取基础配置和导入模板。

依赖包括 PySide6、NumPy、pandas、Matplotlib 和 openpyxl，见 [requirements.txt](requirements.txt)。

## 使用流程

1. 启动主程序，选择桥梁类别与具体桥型。
2. 填写项目和桥梁信息，设置构件数量及构件是否存在。
3. 录入病害信息，或使用 `Import_data/` 中对应桥型的 Excel 模板导入。
4. 查看构件、结构和全桥的评分及评定结果。
5. 保存项目，或导出 Excel 数据和评分报告。

## 目录结构

```text
Highway_bridge_assessment/
├── highway_bridge_assessment_system.py  # 主程序和桥型选择界面
├── *_bridge.py                         # 10 个桥型评定模块
├── Basic_data/                         # 桥型基础配置 JSON 和程序图标
├── Import_data/                        # 按桥型分类的 Excel 病害数据模板
├── Saved_Projects/                     # 运行时保存的项目 JSON
├── Export_data/                        # 运行时导出的数据及评分报告
├── requirements.txt                    # Python 依赖
└── LICENSE                             # MIT 许可证
```

仓库包含全部 Python 源码、基础配置和导入模板。`Saved_Projects/` 与 `Export_data/` 中的运行数据由本地使用者管理，已在 `.gitignore` 中忽略。

## 作者与许可证

作者：Mingtao Zhang

本项目采用 [MIT License](LICENSE)。
