# MathType / LaTeX 解析与联调备忘

> 记录日期：2026-04-01；**更新**：2026-04-12（第六节：导出公式「乱窜」根因与可复制验证模板）  
> 说明：阶段性记录；**第五、六节**为可复制给协作者/AI 的一键上下文。

## 一、已做事项（摘要）

### Parser

- Word 导出：`answer` / `analysis` 中 **`$...$` 行内公式** 与题干同源（MathType OLE → OMML → 斜体）；此前答案整段纯文本导致公式不显示，已修复（见第三节 §5）。
- MathType OLE 优先尝试提取 LaTeX，失败回退图片；增加噪声过滤与状态码、缓存。
- 可插拔外部转换：`MATHTYPE_LATEX_CMD`（例如 `mtef-go.exe -f {ole_path}`）、`MATHTYPE_LATEX_MODE`、`MATHTYPE_LATEX_TIMEOUT`。
- 公式资产与汇总里暴露 `mathtype_latex` 相关状态，便于确认是否走到外部工具。
- 拆题：`splitter` 修复选择题选项中公式未随 HTML 返回的问题；MathType 有 LaTeX 时优先前端 MathJax。

### Web

- `index.html` 中为 MathJax 配置 `$...$` / `$$...$$` 等分隔符，避免公式以纯文本显示。

### 工具与文档

- `tools/mathtype-latex-bridge`：.NET 桥接骨架（后续可接 SDK）。
- `docs/mathtype-latex-tasklist.md`：任务清单与灰度开关说明。
- `docs/mathtype-latex-test-report.md`：测试记录模板。
- Git 分支：`feat/mathtype-latex-external`（具体以仓库 `git log` 为准）。

### 未纳入仓库的本地文件（若存在）

- `tools/mtef-go.exe` 等二进制，通常勿提交；以本机路径配置 `MATHTYPE_LATEX_CMD`。

---

## 二、启动解析前：环境变量要不要重新设？

**取决于设置方式：**

| 方式 | 下次启动 |
|------|----------|
| 仅在当前 PowerShell 会话中 `$env:XXX=...` | **需要**：关窗口或新开终端后失效，需重新设置再执行 `pnpm dev`。 |
| 在 Windows「系统/用户环境变量」中永久配置 | **一般不需要**：新终端会自动继承；修改后需**重启**正在跑的 `pnpm dev`。 |

与是否在项目目录执行无关；关键是 **启动 `pnpm dev` 的进程能读到这些变量**。

### 常用变量（Parser 进程）

- `MATHTYPE_LATEX_MODE`：如 `auto` / `external` / `heuristic` / `none`
- `MATHTYPE_LATEX_CMD`：示例  
  `E:\path\to\mtef-go.exe -f {ole_path}`
- `MATHTYPE_LATEX_TIMEOUT`：秒数，如 `20`

### 一次性设置（三条，复制执行）

**PowerShell**（与 `pnpm dev` 同一窗口；`{ole_path}` 为占位符，Parser 会替换为实际 OLE 临时文件路径）：

```powershell
$env:MATHTYPE_LATEX_MODE = "external"
$env:MATHTYPE_LATEX_CMD = "E:\Lvcy\practice\math-demo\tools\mtef-go.exe -f {ole_path}"
$env:MATHTYPE_LATEX_TIMEOUT = "20"
```

将 `MATHTYPE_LATEX_CMD` 里的路径改成你本机 `mtef-go.exe` 的实际位置；若路径含空格，给 exe 加引号，例如：

```powershell
$env:MATHTYPE_LATEX_CMD = '"D:\My Tools\mtef-go.exe" -f {ole_path}'
```

**cmd（命令提示符）**（不要用 `$env:`）：

```bat
set MATHTYPE_LATEX_MODE=external
set MATHTYPE_LATEX_CMD=E:\Lvcy\practice\math-demo\tools\mtef-go.exe -f {ole_path}
set MATHTYPE_LATEX_TIMEOUT=20
```

设完后在同一窗口执行 `pnpm dev`（或先 `cd` 到仓库根目录再执行）。

### PowerShell 快速检查

```powershell
echo $env:MATHTYPE_LATEX_MODE
echo $env:MATHTYPE_LATEX_CMD
echo $env:MATHTYPE_LATEX_TIMEOUT
```

注意：在 **cmd** 里用 `set`，在 **PowerShell** 里用 `$env:`，不要混用。

---

## 三、导出与 ConvertEquations（latexToMathType）

Parser 导出 Word 时，若配置了 **`EXPORT_CONVERT_EQUATIONS_EXE`**，会在 **OMML / 公式图** 之前尝试调用本机 **`ConvertEquations.exe --latex`**。  
**当前实现**：若 CLI 返回 **OLE** 且嵌入成功，docx 中会写入 **MathType OLE（`w:object` + `word/embeddings/`）**；失败则回退 **OMML** 或 **data-image**。

### 1. 编译 exe（仓库根目录）

```text
pnpm run build:convert-equations
```

- 会先 `dotnet restore`：`ensure-net48-ref`、`ensure-office-interop`（Office Interop NuGet），再 `dotnet msbuild` 解决方案。  
- **NuGet / MSB3644 / COM** 等排错仍见下文「编译与还原」小节。

**默认输出路径**（以本仓库 csproj 为准）：

```text
tools/latexToMathType/ConvertEquations/ConvertEquations/bin/Release/ConvertEquations.exe
```

- `Release|Any CPU` 已固定 **`PlatformTarget=x86`**（与 **MT6.dll 为 32 位**一致）；勿改为纯 AnyCPU 在 x64 系统上跑，否则易报无法加载 `MT6.dll`。

### 2. MathType / MT6.dll（必看）

SDK（`MTSDKDN`）通过 **`DllImport("MT6.dll")`** 加载 **MathType 原生库**。

| 说明 | 要点 |
|------|------|
| **`MT6.dsc` ≠ `MT6.dll`** | 只有 **`MT6.dll`** 可被加载；根目录仅有 `.dsc` 不够。 |
| **常见位置** | **`{MathType安装根}\System\MT6.dll`**（例如 `D:\MathType\System\MT6.dll`）。 |
| **PATH** | 必须把 **`...\MathType\System`** 加入 PATH，**不要**只加 `...\MathType` 根目录。 |
| **仍报 0x8007007E** | 多为依赖缺失：把 **`System\*.dll`** 复制到 **`ConvertEquations.exe` 同目录**，或确保已装 **VC++ 2015–2022 可再发行组件（x86）**。 |

**当前会话临时 PATH（PowerShell）**：

```powershell
$env:Path = "D:\MathType\System;" + $env:Path   # 按本机修改盘符与路径
```

**验证 CLI（应先输出一行 JSON，且含 `ole` 字段）**：

```powershell
$env:Path = "D:\MathType\System;" + $env:Path
& "E:\Lvcy\practice\math-demo\tools\latexToMathType\ConvertEquations\ConvertEquations\bin\Release\ConvertEquations.exe" --latex "x^2+1"
```

### 3. Parser 侧环境变量（导出走 ConvertEquations）

**启动 Parser 的进程**必须能：① 找到 **`ConvertEquations.exe`**；② 加载 **MT6.dll**（同上，**系统/用户 PATH 含 `...\MathType\System`**，或与 exe 同目录已有 DLL）。

```powershell
$env:EXPORT_CONVERT_EQUATIONS_EXE = "E:\Lvcy\practice\math-demo\tools\latexToMathType\ConvertEquations\ConvertEquations\bin\Release\ConvertEquations.exe"
$env:EXPORT_CONVERT_EQUATIONS_TIMEOUT = "120"   # 可选，默认 120 秒
# 可选：OLE 的 ProgID，默认 Equation.DSMT4；老环境可试 Equation.3
# $env:EXPORT_MATHTYPE_OLE_PROGID = "Equation.DSMT4"
```

然后 **`pnpm dev`** 或仅起 Parser；并确认 **`apps/api/.env` 中 `PARSER_URL`** 指向实际 Parser 端口（见根目录 `AGENTS.md`）。

### 4. 行为与限制

- **很慢**：每个公式可能调一次 CLI（内部会起 Word）；同一份 docx 内相同 LaTeX 会缓存，避免重复调用。  
- C# 侧 **`InvalidLatex`** 会拒绝部分 LaTeX，失败则走 OMML/图。  
- 未设置 **`EXPORT_CONVERT_EQUATIONS_EXE`** 时，不调用该工具，导出以 OMML/图为主。

### 5. 导出 Word 公式形态（目标：MathType 优先）与答案/解析

**目标**：导出得到的 **`.docx` 中公式优先为可编辑的 MathType 对象**（OLE，Word 中双击可进 MathType）；在无法生成 OLE 时依次回退 **Word 原生 OMML**、**公式图**、**斜体占位文本**。

| 环节 | 说明 |
|------|------|
| 题干 / 选项 | 来自 `content_html` / `content_html`，按 `math-inline` 等节点插入。 |
| **答案 / 解析** | 解析结果为纯文本字段 `answer` / `analysis`；其中 **`$...$` 行内公式** 与题干共用同一套逻辑：**ConvertEquations（MathType OLE）→ OMML → 斜体**。此前答案曾整段写入单个 `run`，公式无法排版，已修复。 |

**限制**：若题目里答案、解析在解析阶段 **没有** 以 `$...$`（或其它已支持结构）给出公式，而只是一段纯文字，导出 **无法凭空变成** MathType 对象；需改进拆题或源文档规范。

### 6. 编译与还原（子问题速查）

- **推荐**：`pnpm run build:convert-equations`（华为云 + nuget.org 双源还原）。  
- **仅官方源**：`pnpm run build:convert-equations:nuget-org`。  
- **`Directory.Build.props`** 使用 **`NUGET_PACKAGES` 或 `%USERPROFILE%\.nuget\packages`** 定位 net48 与 Office Interop，**勿依赖错误的 `NuGetPackageRoot`**。  
- **MSB3644**：需 **`TargetFrameworkIdentifier=.NETFramework`**（已在 `tools/latexToMathType/Directory.Build.props`）+ 成功 restore **`Microsoft.NETFramework.ReferenceAssemblies.net48`**。  
- **MSB4216（ResolveComReference）**：工程已改为 **NuGet `Microsoft.Office.Interop.Word`**，应用 **`dotnet msbuild`** 即可，一般不再依赖本机 COM 解析生成 Interop。

---

## 四、后续可能方向（待你决定再动）

- 导出：Word 内「可编辑公式」若不要求必须是 MathType OLE，可考虑 **LaTeX → OMML** 等服务端方案；若必须 MathType 对象，需 SDK 或等价转换链。
- 一键启动：可加 `scripts/start-dev-with-mtef.ps1`，在脚本内设置变量后调用 `pnpm dev`。

---

## 五、可复制给 AI / 协作者的上下文（一键粘贴）

下次联调、排错或换机器时，把 **本节整段**（含模板）复制到对话里，并填好 **本机路径** 与 **现象**。第三节有更细的步骤说明。

### 推荐操作顺序（ checklist ）

1. **编译**：仓库根目录 `pnpm run build:convert-equations`。  
2. **MathType**：确认存在 **`{MathType}\System\MT6.dll`**；**PATH 永久或临时包含 `{MathType}\System`**（不要只加安装根目录）。  
3. **验证 CLI**（应有一行 JSON，且含 **`ole`**）：

```powershell
$env:Path = "D:\MathType\System;" + $env:Path   # 改成你的 System 路径
& "E:\Lvcy\practice\math-demo\tools\latexToMathType\ConvertEquations\ConvertEquations\bin\Release\ConvertEquations.exe" --latex "x^2+1"
```

4. **Parser**：在同一类环境中设置 **`EXPORT_CONVERT_EQUATIONS_EXE`**（及可选 **`EXPORT_CONVERT_EQUATIONS_TIMEOUT`**），并保证 **运行 Parser 的进程** 也能加载 **MT6.dll**（同上 PATH 或 DLL 与 exe 同目录）。  
5. **全链路**：`pnpm dev`，`apps/api/.env` 里 **`PARSER_URL`** 正确；Web 导出 docx，用 Word 打开检查公式。

### 发给 AI 的模板（填空后粘贴）

将下方代码块整体复制，把 `【】` 与示例路径改成你的实际情况。

```text
【项目】math-demo 单仓库；Parser=apps/parser；ConvertEquations=tools/latexToMathType；OLE 嵌入逻辑见 apps/parser/app/core/mathtype_ole_embed.py 与 exporter。

【本机】
- 仓库根目录：【如 E:\Lvcy\practice\math-demo】
- MT6.dll：【如 D:\MathType\System\MT6.dll】
- ConvertEquations.exe：【…\bin\Release\ConvertEquations.exe】

【环境】
- MathType\System 已加入【用户/系统 PATH / 仅当前 PowerShell 会话】
- 已执行 pnpm run build:convert-equations：【是/否】
- CLI --latex 测试结果：【能输出含 ole 的 JSON / 不能，完整报错粘贴】

【Parser】
- EXPORT_CONVERT_EQUATIONS_EXE：【已设置路径 / 未设置】
- 若作为服务/IDE 启动：说明 Parser 进程是否继承 PATH（或已拷 DLL 到 exe 目录）

【现象】
【例如：导出 docx 仍是 OMML / Word 打不开 OLE / Parser 报 …】

【期望】
【一句话】
```

---

如需把本节同步到 `DEPLOYMENT.md` 或拆成「运维」与「开发」两篇，再说一声即可。

---

## 六、Word 导出公式「乱窜 / 串题 / 重复嵌」——根因与验证

> 本节记录 **可能原因**（按排查优先级）与 **可复制给 AI 的验证模板**。代码位置以仓库为准：`apps/parser/app/core/exporter.py`、`mathtype_ole_embed.py`、`latex_omml_export.py`、`convert_equations_cli.py`。

### 1. 现象（便于对齐描述）

| 说法 | 含义 |
|------|------|
| 乱窜 / 串位 | 公式与前后文字在阅读顺序或视觉上错位，或跑到句尾/下一段 |
| 串题 | 上一题的式子「嵌进」下一题题干，或内容张冠李戴 |
| 重复嵌 / 叠字 | 同一位置多层公式或图叠在一起 |

### 2. 可能根因（建议按序号排查）

**A. `w:p` 子节点插入方式不一致（结构性，优先怀疑）**

- 正文、普通图多用 **`paragraph.add_run()` → `CT_P.add_r()`**，内部用 `insert_element_before`，会按 OOXML/python-docx 约定插 `w:r`（在含 `w:hyperlink`、书签等时，落点未必是「物理最后一个子节点」）。
- MathType OLE 在 **`embed_mathtype_ole_in_paragraph`** 里对整段 `w:r` 使用 **`paragraph._p.append(r_el)`**，始终接在 `w:p` 子列表末尾。
- **后果**：简单段落往往正常；一旦同段结构变复杂，**`add_r` 与 `append` 的语义分叉**，XML 文档顺序可能与预期读写顺序不一致，表现为窜位、像「上一段内容接错段」。

**B. 同一段混用多种公式载体（版式层）**

- 同一题干内可能交替出现：**MathType OLE（`w:object`）**、**WMF/PNG 图**、**OMML（`m:oMath`）**、斜体回退字。
- Word 对行内 OLE、Drawing、OMML 的 **基线 / 换行 / 环绕** 处理不一致，易出现 **视觉叠盖** 或 **看起来像重复嵌**。

**C. 外部链路 `ConvertEquations.exe` + Word**

- Python 侧已去掉 **ConvertEquations 返回 payload 的内存缓存**，并对 `ole`/`wmf` 字节做拷贝；若仍稳定出现「固定像上一题」，需怀疑 **C# 进程内是否复用 Word Application、剪贴板或临时文件**（需在工具侧加日志或审计）。

**D. 上游数据：`content_export_segments` / `content_html` 顺序或内容**

- 片段顺序若与真卷不一致，导出只能忠实复现，表现为「公式位置错」。应用 **`EXPORT_DEBUG_SEGMENTS=1`** 打日志对照原 docx。

**E. 已缓解但需知晓的历史问题**

- **`get_or_add_image_part` 按 PNG 去重**：多 OLE 共用预览图曾导致「同图多处」；已通过预览图 **nonce** 等方式缓解（见 `mathtype_ole_embed.py`）。
- **OMML 节点被 lxml 搬家**：插入前 **克隆 OMML 子树** 再挂接（见 exporter）。
- **相邻公式片段去重**：易误删不同图；已 **默认关闭**，仅 `EXPORT_SEGMENT_DEDUPE_FORMULAS=1` 时开启。

### 3. 建议验证步骤（可复制执行）

1. **解压 docx**：将导出文件改名为 `.zip`，查看 `word/document.xml` 中该题对应 **`w:p`** 内 **`w:r` / `w:object` / `m:oMath`** 的**子节点顺序**是否与阅读顺序一致。  
2. **对照日志**：`EXPORT_DEBUG_SEGMENTS=1` 重启 Parser，导出后对照 segment 索引与原文。  
3. **单引擎对比**：临时 `EXPORT_FORMULA_PRIORITY=omml`（或关闭 ConvertEquations）导出，观察窜动是否明显减轻，以区分 **B** 与 **A/C**。  
4. **单题 vs 多题**：只导出一题与同一卷多题对比，若仅多题出现「串题」，侧重 **A/C**。

### 4. 发给 AI 的模板（整段复制后填空）

将下方代码块 **整体** 复制到对话，替换 `【】` 内容；可附上 `document.xml` 片段或 `[导出诊断]` 行。

```text
【项目】math-demo；Word 导出相关：apps/parser/app/core/exporter.py、mathtype_ole_embed.py、convert_equations_cli.py、latex_omml_export.py。

【现象】（勾选或描述）
- [ ] 乱窜/句内顺序错  [ ] 像上一题嵌进本题  [ ] 重复嵌/叠盖  [ ] 选择题选项缺失或其它：【】

【环境】
- EXPORT_FORMULA_PRIORITY：【mathtype / omml / auto】
- ConvertEquations：【在用路径 / 未用】
- EXPORT_DEBUG_SEGMENTS：【0 / 1】
- EXPORT_SEGMENT_DEDUPE_FORMULAS：【未设 / 1】

【导出路径】题干走 segments 还是 HTML：【看 Word 里 [导出诊断] stem_mode= 或说明】
【诊断原文粘贴】
【粘贴 [导出诊断] 整行或 document.xml 中该题 w:p 片段（可脱敏）】

【已读文档】docs/mathtype-latex-dev-notes.md 第六节「乱窜根因」。

【诉求】请按第六节优先级（A→E）分析最可能原因，并给出下一步验证或代码修改建议。
```

### 5. 代码改进方向（备忘，非承诺）

- 将 MathType OLE 的 `w:r` 插入改为与 **`CT_P.add_r` / `_insert_r` 同一套语义**，避免与正文 `add_r` 分叉（需改 `mathtype_ole_embed.py` 或 exporter 调用方式）。  
- ConvertEquations 侧：每次调用打印 **LaTeX 输入与 ole/wmf 长度或 hash**，排除进程内状态串扰。
