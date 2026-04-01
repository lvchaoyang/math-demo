# MathType -> LaTeX 测试记录

- 测试日期：
- 测试人：
- 分支：
- Parser 启动方式：
- 前端地址：
- API 地址：
- 备注：

## 环境配置

- `MATHTYPE_LATEX_MODE`：
- `MATHTYPE_LATEX_CMD`：
- `MATHTYPE_LATEX_TIMEOUT`：
- 是否重启 Parser 后生效：是 / 否

---

## 测试样本清单


| 样本ID | 文档名 | 公式类型           | 预期（LaTeX/回退图片） | 实际结果 | 是否通过 | 备注  |
| ---- | --- | -------------- | -------------- | ---- | ---- | --- |
| S01  |     | 分式             |                |      | ✅/❌  |     |
| S02  |     | 根式             |                |      | ✅/❌  |     |
| S03  |     | 上下标            |                |      | ✅/❌  |     |
| S04  |     | 求和/积分          |                |      | ✅/❌  |     |
| S05  |     | 矩阵             |                |      | ✅/❌  |     |
| S06  |     | 混合题干（文字+公式）    |                |      | ✅/❌  |     |
| S07  |     | 纯 MathType OLE |                |      | ✅/❌  |     |
| S08  |     | 异常样本（故意失败）     |                |      | ✅/❌  |     |


---

## 前端渲染检查（逐项打勾）

- 成功提取 LaTeX 的公式由 MathJax 渲染（非图片）
- 行内公式与文本基线对齐正常
- 块级公式居中与换行正常
- 未提取成功时自动回退图片
- 页面无空白公式占位
- 选项中的公式显示正常
- 导出前预览中的公式显示正常

---

## API/解析结果记录（每个样本填一次）

- 样本ID：
- `file_id`：
- `method`：
- `formula_render_summary.total`：
- `formula_render_summary.rendered`：
- `formula_render_summary.source_only`：
- `formula_render_summary.by_source_type`：
- `formula_render_summary.by_note`：
- `formula_render_summary.mathtype_latex_status`：

### 关键状态计数（抄写）

- `ok_external`：
- `ok_heuristic`：
- `external_not_configured`：
- `external_timeout`：
- `no_latex_candidate`：
- `extract_exception`：
- 其他：

---

## 失败样本明细（有失败就填）


| 样本ID | 题号  | 现象  | 状态码 | 可能原因 | 建议修复方向 |
| ---- | --- | --- | --- | ---- | ------ |
|      |     |     |     |      |        |


---

## 模式对照测试（建议至少跑一次）

### A. `MATHTYPE_LATEX_MODE=none`

- 预期：全部走图片回退
- 结果：
- 结论：

### B. `MATHTYPE_LATEX_MODE=heuristic`

- 预期：只走启发式提取
- 结果：
- 结论：

### C. `MATHTYPE_LATEX_MODE=external`

- 预期：只走外部命令，失败不走启发式
- 结果：
- 结论：

---

## 最终结论

- 总样本数：
- 通过数：
- 失败数：
- 当前可上线判断：可 / 不可
- 阻塞问题：
- 下一步动作：

