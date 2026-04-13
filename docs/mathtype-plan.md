# MathType / 公式导出 — 当前进度与续作备忘

> 下次打开本仓库时，先读本文可接上上下文；细节论证见其它文档与源码。

---

## 1. 当前结论（生产默认）

- **Word 导出公式**：已固定为 **纯 OMML**（LaTeX → MathML → OMML，写入 `m:oMath`），**不再**以 ConvertEquations / MathType OLE / 公式 WMF 作为主路径。
- **实测**：在现有试卷数据上，纯 OMML **版式与速度**优于 MathType 链路；MathType 路径易出现重复、错嵌、叠行等问题（根因见下文文档）。
- **解析阶段**：仍可能接触 MathType OLE、WMF 预览等（`parser` / `mathtype_parser`），与「导出引擎选 OMML」独立；若日志里大量「OLE 无 WMF/EMF 预览」，属解析侧资源问题，不表示导出必须走 MathType。

---

## 2. 已落地改动（按模块）

### 2.1 Parser：`apps/parser/app/core/exporter.py`

- 公式插入：`_try_formula_engines_order` → 仅 `_try_insert_math_omml`。
- 已移除：ConvertEquations、MathType OLE 嵌入、公式 WMF 等与导出主路径相关的旧逻辑。
- 已移除：写入 Word 的 **`[导出诊断]`** 段落。
- 仍保留：OMML 失败时的 **`data-image` / 斜体 LaTeX** 等回退；正文插图里的 **磁盘 WMF** 仍可能经 `WMFConverter` 转 PNG（非「公式 WMF 引擎」）。

### 2.2 Parser：进程与 CLI

- `apps/parser/app/core/process_cleanup.py`：`EXPORT_KILL_WINWORD_AFTER`（`parse` / `export` / `convert` / `all`）可选结束 `WINWORD.EXE`（**会关本机所有 Word**，默认关闭）。
- `apps/parser/main.py`：解析/导出结束处可调用上述清理（由环境变量控制）。
- `apps/parser/app/core/convert_equations_cli.py`：`run_latex_to_mathtype_payload` 的 `try/finally` 合法化；子进程默认传 **`CONVERTEQUATIONS_CLI_SAFE=1`**（见下节工具）。

### 2.3 工具：`tools/latexToMathType/ConvertEquations`（C#）

已做工程向优化（需 **重新 `dotnet build -c Release`** 后使用新 `ConvertEquations.exe`）：

- **`CONVERTEQUATIONS_CLI_SAFE`**：为真时失败路径 **不**执行旧版 `Restart()`（避免杀光 WINWORD/MathType 并自重启进程）。**`Program.cs` 的 `--latex` 入口已默认开启**。
- **`GetOLEAndWMFFromOneWord` / `GetOLEAndWMFFromOneWordMML`**：每次调用使用 **独立临时 Flat XML**（`%TEMP%` 下 guid），`finally` 删除；减轻「单文件反复复用」导致的状态污染。
- **`InitWord()`**：主要为 Word **不可见、关弹窗**；不再预先创建全局模板 XML。
- **`DealWordFile` 的 catch**：与上述 safe 模式一致。

### 2.4 文档

- `docs/export-omml-vs-mathtype-analysis.md`：OMML vs MathType/ConvertEquations **原因分析** + 工具内部逻辑与 **上述 C# 优化** 的说明（§7 等）。
- `docs/mathtype-latex-dev-notes.md`：既有 MathType/LaTeX 开发笔记（若与本文冲突，**以本文「当前默认策略」为准**）。

---

## 3. Git / 分支（若需对齐远程）

- 曾创建分支 **`export-omml`** 提交 Parser 侧 OMML 与进程清理相关改动；推送若遇网络问题需本机重试 `git push -u origin export-omml`。
- **后续 C# 工具改动**可能尚未单独成提交；合并前在仓库根执行 `git status` 核对。

---

## 4. 环境变量速查

| 变量 | 作用 |
|------|------|
| `EXPORT_KILL_WINWORD_AFTER` | `parse` / `export` / `convert` / `1` / `all` 等：可选在请求结束后 `taskkill` WINWORD |
| `CONVERTEQUATIONS_CLI_SAFE` | Parser 调 ConvertEquations 时默认 `1`；工具内抑制毁灭性 `Restart()` |
| `EXPORT_CONVERT_EQUATIONS_EXE` | 显式指定 `ConvertEquations.exe` 路径 |
| `PARSER_RELOAD` | `1` 时 Parser 开发热重载（见 `main.py`） |

---

## 5. 续作建议（未完成 / 可选）

按优先级大致如下，**不必一次做完**：

1. **若再次启用「导出为 MathType OLE」**：Python 侧需恢复/重写与 `embed_mathtype_ole_in_paragraph`、`convert_equations_cli` 的对接；并坚持 **每式唯一预览图**、**单次落盘**，见 `export-omml-vs-mathtype-analysis.md` §5。
2. **ConvertEquations 深层改造**：去掉 **剪贴板 + PasteAndFormat**，改为 SDK 能直接落 OLE/WMF 的稳定路径（工作量大，需对照 MathType SDK）。
3. **解析侧**：减少「OLE 无预览」类失败（环境、试卷来源、MT6 路径）；与导出 OMML 策略独立，可单独立项。
4. **前端**：若存在仅用于调试的 `console.log`（如 `ExportPanel.vue`），提交前删除，避免污染控制台。

---

## 6. 本地验证命令

```bash
# 编译 ConvertEquations（Windows，需已装 .NET SDK + Word/MathType 依赖）
dotnet build tools/latexToMathType/ConvertEquations/ConvertEquations.sln -c Release

# Parser
cd apps/parser && py -3 -m uvicorn main:app --host 0.0.0.0 --port 8000
```

---

## 7. 相关路径索引

| 路径 | 说明 |
|------|------|
| `apps/parser/app/core/exporter.py` | Word 导出、OMML 插入 |
| `apps/parser/app/core/latex_omml_export.py` | LaTeX → OMML |
| `apps/parser/app/core/convert_equations_cli.py` | 调用 `ConvertEquations.exe --latex` |
| `apps/parser/app/core/mathtype_ole_embed.py` | MathType OLE 嵌入（当前导出默认不走） |
| `tools/latexToMathType/ConvertEquations/ConvertEquations/mathType.cs` | Word + OLE/WMF 抽取主逻辑 |
| `tools/latexToMathType/ConvertEquations/ConvertEquations/Program.cs` | `--latex` 入口、CLI_SAFE |

---

*文档目的：承上启下；更新时请同步改「§2 已落地改动」与「§5 续作」，避免与代码长期脱节。*
