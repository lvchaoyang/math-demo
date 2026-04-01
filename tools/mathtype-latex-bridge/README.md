# MathTypeLatexBridge

Windows 下的 MathType OLE -> LaTeX 桥接工具（CLI）。

## 目标

- 给 Parser 提供稳定的外部命令入口（`MATHTYPE_LATEX_CMD`）。
- 当前仓库先提供可编译骨架，待接入 MathType SDK 后即可产出稳定 TeX。

## 构建

```bash
dotnet build tools/mathtype-latex-bridge/MathTypeLatexBridge.csproj -c Release
```

## 用法

```bash
MathTypeLatexBridge.exe --ole "E:\path\to\oleObject1.bin" --mode sdk
```

参数：

- `--ole`: OLE 文件路径（必填）
- `--mode`: `sdk`（默认）或 `heuristic`

退出码：

- `0`: 成功（stdout 输出 LaTeX）
- `10`: sdk 模式未实现或调用失败
- `11`: heuristic 未提取到候选
- `12`: 不支持的 mode
- `20`: 运行异常

## 与 Parser 对接

在 Parser 进程环境变量中配置：

```bash
MATHTYPE_LATEX_MODE=external
MATHTYPE_LATEX_CMD=E:\Lvcy\practice\math-demo\tools\mathtype-latex-bridge\bin\Release\net8.0-windows\MathTypeLatexBridge.exe --ole {ole_path} --mode sdk
MATHTYPE_LATEX_TIMEOUT=20
```

如果你想先做端到端联调（未接 SDK）：

```bash
MATHTYPE_LATEX_MODE=external
MATHTYPE_LATEX_CMD=E:\Lvcy\practice\math-demo\tools\mathtype-latex-bridge\bin\Release\net8.0-windows\MathTypeLatexBridge.exe --ole {ole_path} --mode heuristic
```

## 下一步（你拿到 SDK 后）

在 `Program.cs` 的 `TrySdk()` 中接入 MathType SDK API，返回纯 LaTeX 字符串即可。
