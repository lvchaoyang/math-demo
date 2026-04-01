# MathType 转 LaTeX 任务清单

## 目标

- 将 `MathType OLE` 公式从“仅图片渲染”升级为“优先 LaTeX 渲染，失败回退图片”。
- 保证现有题目拆分与导出链路不回归。
- 建立可观测、可回归、可逐步优化的改造基础。

## 阶段 1（本次先落地）

- [x] 在 Parser 增加 MathType LaTeX 提取入口（`extract_latex`）。
- [x] 在 OLE 解析数据结构中携带 `latex` 字段。
- [x] 在公式资产构建时新增 `mathtype_latex` 源类型。
- [x] 在题目 HTML 生成时，优先渲染 `mathtype_latex`（MathJax）。
- [x] 保留图片回退路径，确保提取失败时不影响现有结果。

## 阶段 2（稳定性增强）

- [x] 引入可插拔的 `MTEF -> LaTeX` 转换入口（外部命令优先 + 启发式回退；完整语义转换需自建 CLI 或接入 MathType SDK）。
- [x] 建立 LaTeX 标准化模块（转义、函数名、分隔符、display/inline 归一）。
- [x] 增加失败分类码（不支持结构、解析异常、超时、空结果）。
- [x] 增加转换缓存键，避免同一公式重复转换。

### 环境与灰度开关（Parser 进程）

| 变量 | 说明 | 典型值 |
|------|------|--------|
| `MATHTYPE_LATEX_MODE` | `auto`：有外部命令则先外部，否则启发式；`heuristic`：仅启发式；`external`：仅外部（失败不回退启发式）；`none`：不提取 LaTeX（仅图片） | `auto` |
| `MATHTYPE_LATEX_CMD` | 可执行程序及参数；可写 `{ole_path}` 占位符，否则在末尾追加 OLE 临时文件路径 | 空（不调用外部） |
| `MATHTYPE_LATEX_TIMEOUT` | 外部命令超时秒数 | `15` |

外部程序约定：退出码 0，**标准输出** 为一行或多行 LaTeX；非 0 时回退启发式（`auto`）或失败（`external`）。

### 已落地桥接工具（Windows）

- 路径：`tools/mathtype-latex-bridge`
- 构建命令：`pnpm build:mathtype-latex-bridge`
- 作用：作为 `MATHTYPE_LATEX_CMD` 的目标 exe，后续在工具内部接入 MathType SDK 即可。

## 阶段 3（质量保障）

- [ ] 建立 MathType 样本库（最少 50 例，覆盖分式、根式、积分、矩阵、多行、上下限）。
- [ ] 建立样本期望输出（golden set）并纳入 CI 回归。
- [ ] 增加端到端验证：上传 docx -> 解析 -> 前端渲染快照。
- [ ] 为失败样本添加自动归档，便于迭代修复。

## 阶段 4（可观测与运维）

- [ ] 在 API 返回中补充 MathType 转换统计字段（成功数、失败数、回退数）。
- [ ] 在日志中记录可追踪上下文（file_id、ole_filename、失败原因）。
- [ ] 在部署文档增加依赖说明与开关策略（灰度、回滚）。

## 验收标准

- 数学公式优先以 LaTeX 在前端展示，显示质量优于图片。
- 无法提取 LaTeX 的公式自动回退图片，页面不出现空白公式。
- 现有 `questions` 拆题、导出、图片接口行为保持兼容。
