# Physics Lab Report Generator (physics-lab-report-generator)

## 简介
这是一个面向大学物理实验报告的生成工具集，支持：
- 实验数据处理（均值、样本标准差、3σ坏值检验、不确定度合成、物理修约）
- 图像生成（线性拟合图或趋势/误差图）
- Word `.docx` 报告输出（含 OMML 公式、表格、图表）

## 当前提供的三种入口

### 1) 全能版 Skill（严格流程）
- 文件：`SKILL.md`
- 特点：约束最完整，包含确认门/计算门/作图门/公式门/保存门。
- 适合：希望强流程、强一致性、强审计痕迹的场景。

### 2) Lite 版 Prompt（轻量流程）
- 文件：`实验报告Lite.md`
- 特点：提示词更短，但保留关键硬约束（必须 `.docx`、必须有图、失败 Fail-Closed）。
- 适合：上下文预算紧、模型能力中等、希望更快对齐的场景。

### 3) 弱模型友好一键脚本（推荐 Qwen/DeepSeek）
- 文件：`scripts/one_click_report.py`
- 特点：只改顶部 `REPORT_CONFIG`，运行一次即可输出报告；
  计算过程小节固定写入，减少模型“偷懒省略公式”的风险。

## 目录结构
- `SKILL.md`：全能版技能定义
- `实验报告Lite.md`：轻量版提示词
- `scripts/`
  - `one_click_report.py`：弱模型友好一键生成
  - `data_processor.py`：统计与不确定度处理
  - `omml_generator.py`：OMML 公式构建
  - `report_generator.py`：Word 组装与排版

## 依赖项
- `numpy`
- `pandas`
- `matplotlib`
- `scipy`
- `python-docx`
- `lxml`

## 快速使用（推荐：弱模型一键脚本）
1. 打开 `scripts/one_click_report.py`
2. 只修改 `REPORT_CONFIG`
3. 运行：
   ```bash
   python scripts/one_click_report.py
   ```
4. 查看终端状态字段：
   - 成功：`DOCX_STATUS: SUCCESS` + `DOCX_PATH: ...`
   - 失败：`DOCX_STATUS: FAILED` + `FAILED_REASON` + `REQUIRED_INPUTS`

## 结果质量与交付约束
- 报告必须输出为真实 `.docx` 文件，不能用纯文本替代
- 每份报告至少一张图（线性实验必须拟合图）
- 每组数据必须包含：
  - 算术均值
  - 样本标准差
  - 坏值检验（3σ）
  - 不确定度计算
  - 修约与最终结果
