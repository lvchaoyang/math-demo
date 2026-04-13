# 导出：纯 OMML 稳定 vs ConvertEquations/MathType 易乱版 — 原因分析

本文基于本仓库 **Parser 导出器**（`apps/parser/app/core/exporter.py`）、**OMML 生成**（`latex_omml_export.py`）、**MathType OLE 嵌入**（`mathtype_ole_embed.py`）及 **ConvertEquations** 调用链（`convert_equations_cli.py`）的现有实现，说明为何「纯 OMML」在位置与排版上表现正常，而经 **ConvertEquations.exe → MathType OLE** 的路径更容易出现公式重复、错嵌、叠行等问题。

---

## 1. 两条链路在架构上的根本差异

| 维度 | 纯 OMML 导出 | ConvertEquations → MathType |
|------|----------------|-----------------------------|
| **是否驱动 Word** | 否，全程在 Python 进程内生成 Office Math ML | 是，典型实现通过 **Word 自动化/COM** 完成 LaTeX→公式对象 |
| **写入 docx 的对象类型** | 原生 **`m:oMath` / `m:oMathPara`**（与 Word 公式编辑器一致） | **`w:object`（OLE 嵌入）+ VML 形状 + 预览图关系** |
| **二进制与关系部件** | 无独立 OLE 部件；OMML 内联在文档 XML 中 | 需 `word/embeddings/*.bin`、图片部件、多组 **Relationship**，去重与复用策略敏感 |
| **失败与重试** | 失败可局部回退（如斜体文本、原卷图），不易「多插一层」 | 若上层再叠加 OMML/WMF 回退，易出现 **同一语义多次落盘** |

结论：**OMML 是「单进程、结构化、与 Word 原生公式同一模型」；MathType 路径是「多进程、OLE+图片、与排版引擎交互更复杂」** — 出错面天然更大。

---

## 2. 为何纯 OMML 嵌入位置与排版通常「正常」

### 2.1 生成路径短且确定

实现上（`latex_omml_export.latex_to_omml_element`）：LaTeX → MathML → OMML 字符串 → `parse_xml` 得到元素。不依赖外部 Word 实例，**没有跨请求的进程级缓存**干扰结果。

### 2.2 插入语义与 Word 行模型一致

导出器（`_try_insert_math_omml`）按 HTML 节点区分行内/块级：

- **行内**：新建 `run`，把 OMML 挂到 `run._element`；
- **块级**：清空段落内文本 run 后，将 **`oMathPara` 挂到 `paragraph._p`**。

这与 Word 对「段落内公式 / 独立公式段」的预期一致，**基线、换行、居中**由 OMML 与段落属性共同决定，不会出现 OLE 与周围文字争用同一 run 的复杂情况。

### 2.3 深拷贝避免「同一节点被移走」

`_clone_oxml_fragment` 对 OMML 子树做 **序列化再解析** 的深拷贝后再挂接。注释中明确：避免 lxml 把**同一元素对象**从上一处移走，导致**跨段/跨题串公式**。这是纯内存结构问题，OMML 路径已针对性处理。

### 2.4 无「预览图去重」类问题

原生 OMML 不依赖单独 PNG 预览部件，**不存在**「多个公式共用同一图片 relationship、表面顺序对但显示成同一张图」这类 OLE 路径特有问题。

---

## 3. 为何 ConvertEquations / MathType 路径更容易「重复、乱嵌、叠版」

### 3.1 外部 Word 进程与状态

ConvertEquations 典型行为是启动或附着 **WINWORD**。若未严格做到 **每式独立文档实例** 或 **用完即关**，易出现：

- 同一会话内缓存、剪贴板或临时文档状态影响下一次结果；
- 多题连续导出时 **时序与资源竞争**（与本仓库曾增加的「可选 taskkill WINWORD」动机一致）。

OMML 路径无此维度。

### 3.2 OLE 嵌入与 python-docx 部件模型

`embed_mathtype_ole_in_paragraph`（`mathtype_ole_embed.py`）需要：

- 新建 **OLE 部件**（`embeddings/embeddingN.bin`）；
- 预览 **PNG** 经 `get_or_add_image_part` —— **按图片二进制去重**。

代码中 `unique_preview_png_for_ole` 的文档字符串已说明：若多道公式 **共用相同预览字节**，Word 可能多处显示 **同一张预览图**，表现为 **顺序对但式子重复/错位**。即：**同一排版位置问题，在 OLE 路径上被「关系去重」放大**。

### 3.3 块级 OLE 与段落子节点的处理方式与 OMML 不完全对称

块级公式时，OLE 嵌入会 **删除段落内所有 `w:r`** 再追加 `w:object`（见 `mathtype_ole_embed` 中 `is_block` 分支）。若上游 HTML/片段拆分与「当前 paragraph 里还剩什么」任一不一致，容易出现：

- 本应保留的文字 run 被清空；
- 或与 `_fill_inline` 其它分支组合时 **重复写入**。

OMML 块级路径使用 `_clear_paragraph_text_runs` + `paragraph._p.append`，语义更接近「整段即公式」，但仍与 OLE 的 **VML+object** 结构不同，**与周围 DOM 的交互更敏感**。

### 3.4 ConvertEquations 返回层面的「可疑复用」

历史导出逻辑中曾存在对 **同一 payload 哈希对应不同 LaTeX** 的检测（防 ConvertEquations **错误复用**导致错嵌）。说明在真实环境中：**输出字节级结果并非总能与输入 LaTeX 一一对应**，上层若再当作强一致数据嵌入 docx，会出现 **张冠李戴式错嵌**。

### 3.5 多引擎回退叠加（若仍启用）

若策略为「先 MathType，失败再 OMML 再 WMF」，**同一公式**可能在不同层被尝试多次；一旦条件判断或「是否已写入成功」标记有偏差，就会产生 **重复插入**。纯 OMML 策略则通常是 **单次尝试 + 明确失败回退**，路径更短。

### 3.6 解析阶段与 OLE 预览（日志现象）

解析管线对 MathType OLE 常需 **WMF/EMF 预览** 才能稳定转图或抽 LaTeX。终端中大量 **「OLE 流内没有找到 WMF/EMF 预览图」** 类日志时，说明 **源文档中 OLE 与预览不完整** —— 后续再经 ConvertEquations 生成新 OLE，**与上游数据质量、MathType 安装路径** 强相关，问题会传导到导出观感（缺图、回退路径不一致等）。

---

## 4. 归纳：问题类别对照

| 现象 | OMML 路径更可能原因 | MathType/ConvertEquations 路径更可能原因 |
|------|---------------------|----------------------------------------|
| 公式重复出现 | 片段/HTML 重复遍历、去重逻辑 bug | OLE 预览图去重、ConvertEquations 缓存、多引擎多次插入 |
| 顺序错乱 | OMML 节点未克隆导致串题（已通过 clone 缓解） | Word 会话状态、段落 run 清空范围与 DOM 不一致 |
| 叠行/选项区乱 | `$` 拆分与 `math-inline` 嵌套问题 | OLE 块级与 div/p 边界 + 对象高度/行距 |
| 显示对但「内容错」 | 少见 | payload 与 LaTeX 不一致、错 OLE 绑定 |

---

## 5. 与当前仓库默认策略的关系

当前导出器默认采用 **纯 OMML**（`exporter` 中 `_try_formula_engines_order` → `_try_insert_math_omml`），**不再走 ConvertEquations/MathType OLE 作为公式主路径**。这与上述分析一致：**在可控性、可复现性和版式稳定性上，OMML 更适合作为默认生产路径**。

若将来需再次启用 MathType 路径，建议同时：

1. 保证每次嵌入 **唯一预览图**（或等价 nonce），避免 `get_or_add_image_part` 合并；
2. 明确 **单次公式单次落盘**，避免多引擎叠加入；
3. 对 ConvertEquations 输出做 **输入/输出一致性校验**（如 LaTeX 与 payload 哈希登记）；
4. 在 Windows 上规范 **Word 子进程生命周期**，避免残留实例。

---

## 6. 相关源文件索引

- OMML 生成与插入：`apps/parser/app/core/latex_omml_export.py`，`apps/parser/app/core/exporter.py`（`_try_insert_math_omml`、`_clone_oxml_fragment`）
- MathType OLE 嵌入：`apps/parser/app/core/mathtype_ole_embed.py`
- ConvertEquations CLI：`apps/parser/app/core/convert_equations_cli.py`
- 可选进程清理：`apps/parser/app/core/process_cleanup.py`，`apps/parser/main.py`

---

## 7. ConvertEquations.exe 内部在做什么？是否「工具不行」？

### 7.1 它不是什么

它**不是**一个独立的「LaTeX→OLE 纯算法库」。CLI 模式（`--latex`）本质是：

1. **MathType SDK**：把 LaTeX 转成 MTEF，再参与生成 OLE/剪贴板数据（见 `mathType.cs` 中 `GetMTEFBytesFromLatex` 等）。
2. **Microsoft Word COM**：使用**全局单例** `Word.Application`；在**每次** `GetOLEAndWMFFromOneWord` / `GetOLEAndWMFFromOneWordMML` 调用中，于 `%TEMP%` 下生成**独立** Flat XML（新建空文档 → 另存为 guid.xml → 打开 → 写入后再读回），避免旧版「长期复用同一路径文件」带来的状态污染（仍保留全局 Word 实例与剪贴板粘贴）。
3. **剪贴板**：`DealWordFile` 中先删掉文档里所有 `InlineShape`，再通过剪贴板 **`PasteAndFormat`** 把新公式贴进 Word，再保存成 XML，最后从 XML 里解析出 **OLE（embedding）与 WMF 的 base64**。

核心片段逻辑（语义概述，非逐行照抄）：

- `InitWord()`：**仅**将 `wordAppGlobal.Visible = false`、关闭告警弹窗（不再预先创建全局模板 XML）。
- `GetOLEAndWMFFromOneWord(latex)`：为本次调用分配**独立**临时 `*.xml` → `Documents.Add` → Flat XML `SaveAs` → `Close` → `Open` → `DealWordFile(MTEF)` → `readMathTypeDontDelete()` → `finally` 删除临时文件。
- `DealWordFile`：遍历 `wordDocGlobal.InlineShapes` **全部 Delete** → `GetWMFBase64FromClipboard` + **粘贴** → `SaveAs` 回当前 `path`。
- **CLI 安全模式**：环境变量 **`CONVERTEQUATIONS_CLI_SAFE=1`**（`Program --latex` 已设置；Parser 子进程默认传入）时，失败分支**不**执行 `Restart()`（旧逻辑会杀光本机 WINWORD/MathType 并自重启进程）。

`Program.latexToMathType` 还要求 **`ole` 与 `wmf` 均非空** 才返回 JSON，否则整次算失败。

### 7.2 问题更多来自「机制」而非单一 exe 写坏了

- **全局 Word + 单文件模板 + 剪贴板**：链路过长，任何一步（Word 未就绪、上一次残留 shape、剪贴板被占用、Flat XML 解析边界）都会造成 **内容与前一次混淆** 或 **抽取出错**，表现成「重复 / 错嵌」。这**换另一个同样依赖 Word+剪贴板的封装**，往往仍脆弱。
- **InvalidLatex / CleanLatex**：工具内对 LaTeX 做了大量替换与**拒绝规则**（如部分符号、`\not`、`boldsymbol` 等），与「标准 LaTeX 全集」不一致，容易 **静默失败或走异常分支**。
- **与 Python 侧嵌入叠加**：即便 exe 输出正确，python-docx 写 OLE、图片去重、段落结构仍可能引入前文所述问题。

因此：**不一定要把失败简单归结为「ConvertEquations.exe 质量差」**；更准确说法是：**当前这条「Word 自动化 + 剪贴板 + 单文档复用」管线，天生比 OMML 难做对**。

### 7.3 换工具会不会更好？

| 目标 | 更稳妥的方向 |
|------|----------------|
| **版式稳定、与 Word 兼容** | **OMML**（本仓库默认）：LaTeX→MathML→OMML，无 Word 进程。 |
| **必须可编辑为 MathType 对象** | 仍离不开 **MathType + Word（或厂商 SDK）**；换「另一个 exe」若仍是同类 COM+剪贴板架构，**改善有限**。若商业方案提供 **稳定 API、无剪贴板、每式独立文档实例**，会明显好于当前脚本式流程。 |
| **只要视觉一致、不要可编辑** | **高质量 PNG/SVG**（解析阶段渲染或单独服务），不走 OLE。 |

**结论**：若业务允许，**继续以 OMML 为主**通常是最省心的；若强需求 MathType OLE，应评估 **是否值得投资在「去剪贴板、去单文件复用、每式隔离 Word 文档」** 的改造或商业组件，而不是仅替换一个同名功能的 exe。

---

*文档版本：与仓库「纯 OMML 默认导出」策略同步撰写，仅供架构与排障参考。*
