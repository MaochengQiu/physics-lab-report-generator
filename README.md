# Physics Lab Report Generator (physics-lab-report-generator)

## 简介
这是一个专为大学物理实验报告设计的自动化生成工具。它能够根据实验照片（讲义、手写数据记录表）或电子表格数据，自动进行高精度的数据处理（含坏值检验、不确定度计算及物理修约），并生成包含专业 OMML 公式和精美排版的 Word (.docx) 实验报告。

## 核心功能
- **数据处理引擎**：基于 `data_processor.py`，支持 Grubbs' 准则 (3σ) 迭代坏值检验、A/B类不确定度合成及严格的物理修约规则。
- **OMML 公式生成**：通过 `omml_generator.py` 程序化构建 Office Math XML 树，生成的公式在 Word 中完全可编辑且美观。
- **高保真排版**：`report_generator.py` 严格对齐专业实验报告范本的字体、字号、表格样式及页面布局。

## 目录结构
- `SKILL.md`: 技能定义的详细说明文档。
- `scripts/`:
  - `data_processor.py`: 核心计算与修约逻辑。
  - `omml_generator.py`: OMML 公式组件库。
  - `report_generator.py`: Word 文档组装与样式控制。

## 依赖项
- `numpy`
- `pandas`
- `matplotlib`
- `scipy`
- `python-docx`
- `lxml`

## 使用说明
该技能设计用于在 Manus 平台中使用。只需将此仓库中的 `SKILL.md` 和 `scripts/` 提供给 Manus，即可触发高标准的物理实验报告生成流程。
