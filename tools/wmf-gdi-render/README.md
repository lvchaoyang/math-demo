# WmfGdiRender（Windows / GDI+）

在 **Windows** 上用 **GDI+**（`System.Drawing.Metafile`）把单独的 **WMF/EMF** 栅格化为 **PNG**，通常比 Inkscape / LibreOffice / `wmf2svg` 更接近 Word 对同一份元文件的观感（仍是近似，不能保证与 Word 版式 100% 一致）。

**给 Cursor / 后续在 Windows 上续作的说明（必读）**：见同目录 **[CURSOR_HANDOFF.md](./CURSOR_HANDOFF.md)**（含仓库改动清单、环境变量、API 行为、可复制给 AI 的提示语）。

## 环境

- Windows 10/11 或 Windows Server
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

## 构建

```bat
cd tools\wmf-gdi-render
dotnet build -c Release
```

可执行文件在：`tools\wmf-gdi-render\bin\Release\net8.0-windows\WmfGdiRender.exe`（或 `dotnet run` 时用 `bin\Debug\...`）。

## 用法

```bat
WmfGdiRender.exe input.wmf output.png --dpi 300
```

- `--dpi` 可选，默认 `300`，范围约 `72`–`1200`。

## 与 math-demo API 对接

在 **Windows** 上跑 Node API 时设置环境变量（路径改成你的实际 exe）：

```bat
set WMF_GDI_RENDER_EXE=C:\path\to\WmfGdiRender.exe
```

然后访问 WMF 图片时加 `method=gdi`（或 `auto` 在 Windows 上会优先尝试 GDI）：

`/api/v1/images/{fileId}/xxx.wmf?method=gdi&dpi=300&normalize=0`

未设置 `WMF_GDI_RENDER_EXE` 或非 Windows 时，`method=gdi` 会失败并回退其它引擎（若走 `auto`）。

## 说明

- 仅处理**磁盘上的单个** `.wmf` / `.emf`，与 Word 里「整页排版中的对象」仍可能有细微差别。
- 授权：随业务自行满足 Windows / .NET 运行环境即可；不涉及 Word 安装。
