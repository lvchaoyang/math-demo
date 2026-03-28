# Windows WMF/EMF → GDI+ 渲染：给 Cursor / 协作者的交接说明

> **用途**：在 Windows 上用 **系统 GDI+**（`System.Drawing.Metafile`）把**单独的** `.wmf` / `.emf` 栅格成 PNG，使公式图观感**尽量接近 Word**，而不依赖 Inkscape / LibreOffice / `wmf2svg` 的跨平台近似渲染。  
> **平台**：仅 **Windows**（`win32`）。macOS/Linux 上此路径不可用，API 会走原有引擎。

---

## 1. 背景（为什么要做）

- 试卷公式多为 **MathType OLE + WMF 预览**；用开源工具转图容易出现 **重叠、挤压、符号乱码**。  
- **GDI+ 解释元文件**与 Word 同属 Windows 图形栈，对**同一份 WMF/EMF 文件**通常比跨平台工具更接近 Word。  
- 注意：这是**单文件栅格化**，不是「用 Word 打开整篇 DOCX 再截图」，因此与 Word 里**整页排版**仍可能有细微差异。

---

## 2. 仓库里相关文件（改动了什么）

| 路径 | 作用 |
|------|------|
| `tools/wmf-gdi-render/WmfGdiRender.csproj` | .NET 8 **windows** 目标，引用 `System.Drawing.Common` |
| `tools/wmf-gdi-render/Program.cs` | CLI：`WmfGdiRender.exe <input.wmf\|emf> <output.png> [--dpi 72-1200]` |
| `tools/wmf-gdi-render/README.md` | 人类可读简版说明 |
| `tools/wmf-gdi-render/CURSOR_HANDOFF.md` | **本文件**：给后续 Cursor 会话的完整上下文 |
| `apps/api/src/routes/images.ts` | 在 **Windows + 配置 exe** 时接入 GDI；`method=gdi`；`auto` 优先 `gdi`；缓存键含 `v13` 等 |

**Python 解析器**：无需为「仅 GDI」改逻辑；仍输出带 `file_id_` 前缀的 WMF 路径，由 Node 图片路由按需转 PNG。

---

## 3. Windows 上必须满足的前提

1. **Windows 10/11** 或 Windows Server（`os.platform() === 'win32'`）。  
2. 安装 **[.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)**。  
3. 在本机编译出 `WmfGdiRender.exe`（见下节）。  
4. 启动 **Node API**（`@math-demo/api`）前设置环境变量 **`WMF_GDI_RENDER_EXE`** 为 exe 的**绝对路径**。

---

## 4. 编译与本地验证（命令）

在 **Windows** 上执行：

```bat
cd tools\wmf-gdi-render
dotnet build -c Release
```

典型输出位置：

`tools\wmf-gdi-render\bin\Release\net8.0-windows\WmfGdiRender.exe`

单文件测试：

```bat
WmfGdiRender.exe C:\path\to\test.wmf C:\path\to\out.png --dpi 300
```

若报错，检查：路径无中文引号问题、文件确实是 WMF/EMF、.NET 8 已安装。

---

## 5. 与 Node API 的对接（环境变量与 URL）

### 环境变量

```bat
set WMF_GDI_RENDER_EXE=C:\你的路径\WmfGdiRender.exe
```

然后照常启动 monorepo 的 `pnpm run dev` 或仅启动 API。

### HTTP 行为（`apps/api/src/routes/images.ts`）

- 仅当 **`win32` 且 `WMF_GDI_RENDER_EXE` 指向存在的文件** 时，转换候选里才会加入 **`gdi`**。  
- **`method=gdi`**：只接受 GDI 结果；失败则 **500**，不会静默改用其它引擎（避免误以为「已是 GDI」）。  
- **`method=auto`（默认 query 或不写 method 时的自动分支）**：优先 **`gdi`**，再 **`wmf2svg` → inkscape → magick → soffice`**（与无 GDI 时相比，仅在 Windows 配置 exe 后多插一层 `gdi`）。  
- 其它 query 仍有效：`dpi`、`normalize`、`fit`、`outline` 等；**`dpi` 会传给 `WmfGdiRender.exe --dpi`**。  
- 缓存文件名包含版本 **`v13`** 及 method/dpi 等，改引擎后勿混用旧缓存目录可删 `data/image_cache` 下对应文件。

### 示例 URL（前端经 3000 代理到 API 时）

```
GET /api/v1/images/{fileId}/{fileId}_image33.wmf?method=gdi&dpi=300&normalize=0
```

`fileId` 与文件名前缀需与 `data/images/{fileId}/` 下实际文件一致。

---

## 6. 给「在 Windows 上新开 Cursor 会话」时的提示语建议

可把下面整段贴给 Cursor：

```
请阅读 tools/wmf-gdi-render/CURSOR_HANDOFF.md。我在 Windows 上已安装 .NET 8，
需要：1）编译 WmfGdiRender；2）设置 WMF_GDI_RENDER_EXE；3）确认 apps/api 图片路由
对 .wmf 请求能走 method=gdi 或 auto。若转换失败，根据错误信息排查路径与权限。
```

---

## 7. 局限与后续方向（避免错误预期）

- **不是 Word COM**：若仍不满意，下一档是 **Word 自动化导出**（需 Word 授权与无人值守设计）。  
- **非 Windows**：不会使用 GDI；不要在 Mac 上期待 `method=gdi` 成功。  
- **极端 WMF**：帧尺寸异常、损坏元文件可能导致 `Metafile` 构造失败，需个案处理或回退 `method=wmf2svg`。

---

## 8. 关键代码定位（便于跳转）

- GDI 子进程调用：`apps/api/src/routes/images.ts` → `convertWithGdiRender`、`isGdiMetafileRendererAvailable`  
- 候选与 `method=gdi` 强制失败逻辑：`convertWmfBestEffort` 内 `forcedMethod !== 'auto'` 分支  
- CLI 实现：`tools/wmf-gdi-render/Program.cs`

---

*文档版本：与仓库内 `WMF_CACHE_VERSION`（当前 `v13`）及 `images.ts` 行为一致；若你改了缓存版本或路由，请同步更新本节。*
