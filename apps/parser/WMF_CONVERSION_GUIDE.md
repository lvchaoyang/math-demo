# WMF 图片转换安装指南

## 问题说明

数学试卷中的公式图片使用的是 WMF (Windows Metafile) 格式，这是旧版 Word 公式编辑器的格式。浏览器无法直接显示 WMF 文件，需要将其转换为 PNG 格式。

## 解决方案

### 方案一：安装 Inkscape（推荐）

Inkscape 是开源的矢量图形编辑器，对 WMF 格式支持最好。

#### macOS
```bash
brew install inkscape
```

#### Ubuntu/Debian
```bash
sudo apt-get update
sudo apt-get install inkscape
```

#### Windows
1. 访问 https://inkscape.org/release/
2. 下载 Windows 安装包
3. 安装时选择"添加到 PATH"

### 方案二：安装 LibreOffice

LibreOffice 可以转换 WMF 文件为 PNG。

#### macOS
```bash
brew install --cask libreoffice
```

#### Ubuntu/Debian
```bash
sudo apt-get install libreoffice
```

#### Windows
1. 访问 https://www.libreoffice.org/download/download/
2. 下载并安装

### 方案三：使用 ImageMagick + LibreOffice

ImageMagick 需要 LibreOffice 作为后端来处理 WMF 文件。

#### macOS
```bash
brew install imagemagick
brew install --cask libreoffice
```

#### Ubuntu/Debian
```bash
sudo apt-get install imagemagick libreoffice
```

## 验证安装

运行以下命令验证安装是否成功：

```bash
# 检查 Inkscape
inkscape --version

# 检查 LibreOffice
libreoffice --version

# 检查 ImageMagick
magick --version  # v7
convert --version  # v6
```

## 转换效果

安装转换工具后，重新解析文档：

```bash
cd apps/parser
python3 -c "
from app.core.parser import parse_docx
from app.core.splitter import split_questions

result = parse_docx('test.docx', extract_images=True)
questions = split_questions(result['paragraphs'])
print(f'解析完成，共 {len(questions)} 道题')
"
```

## 临时方案

如果无法安装转换工具，可以：

1. **前端处理**：使用 JavaScript 库在前端转换 WMF
2. **云服务**：使用云函数（AWS Lambda、阿里云函数计算）进行转换
3. **手动转换**：使用在线工具批量转换 WMF 文件

## 常见问题

### Q: 为什么需要转换工具？
A: WMF 是 Windows 专有格式，浏览器无法直接显示。需要转换为 PNG/JPEG 等通用格式。

### Q: 哪个工具效果最好？
A: Inkscape 效果最好，转换质量高且速度快。LibreOffice 次之。ImageMagick 需要依赖 LibreOffice。

### Q: 转换失败怎么办？
A: 检查转换工具是否正确安装，查看日志中的错误信息。如果转换失败，系统会保留原始 WMF 文件。

### Q: 可以批量转换已有的 WMF 文件吗？
A: 可以，使用以下脚本：

```python
from pathlib import Path
from app.core.wmf_converter import WMFConverter

converter = WMFConverter()
wmf_dir = Path('data/images')

for wmf_file in wmf_dir.rglob('*.wmf'):
    png_file = wmf_file.with_suffix('.png')
    success, result = converter.convert(str(wmf_file), str(png_file))
    if success:
        print(f'转换成功: {wmf_file.name}')
    else:
        print(f'转换失败: {wmf_file.name} - {result}')
```

## 性能优化

对于大量 WMF 文件，建议：

1. 使用 Inkscape（速度最快）
2. 启用并行处理
3. 缓存转换结果
4. 使用 SSD 存储

## 技术细节

### WMF 格式
- Windows Metafile Format
- 矢量图形格式
- Word 2003 及更早版本使用
- 包含 GDI 绘图命令

### 转换流程
```
WMF 文件 → 提取 → 转换工具 → PNG 文件 → 前端显示
```

### 支持的格式
- 输入：WMF, EMF
- 输出：PNG, JPEG, SVG
