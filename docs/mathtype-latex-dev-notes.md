# MathType / LaTeX 解析与联调备忘

> 记录日期：2026-04-01  
> 说明：阶段性记录，后续可继续补充或拆分。

## 一、已做事项（摘要）

### Parser

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

## 三、后续可能方向（待你决定再动）

- 导出：Word 内「可编辑公式」若不要求必须是 MathType OLE，可考虑 **LaTeX → OMML** 等服务端方案；若必须 MathType 对象，需 SDK 或等价转换链。
- 一键启动：可加 `scripts/start-dev-with-mtef.ps1`，在脚本内设置变量后调用 `pnpm dev`。

---

如需把本节同步到 `DEPLOYMENT.md` 或拆成「运维」与「开发」两篇，再说一声即可。
