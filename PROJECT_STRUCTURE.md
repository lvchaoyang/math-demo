# 数学试卷解析系统 - 项目结构说明

## 目录

1. [项目概述](#项目概述)
2. [整体架构](#整体架构)
3. [目录结构详解](#目录结构详解)
4. [模块说明](#模块说明)
5. [数据流](#数据流)
6. [技术栈详解](#技术栈详解)
7. [配置文件说明](#配置文件说明)
8. [开发规范](#开发规范)

---

## 项目概述

数学试卷解析系统是一个基于 Monorepo 架构的 Web 应用，用于解析 Word 格式的数学试卷，提取题目、公式和图片，并支持导出功能。

### 核心功能

- 📄 **Word 解析**: 解析 docx 格式的数学试卷
- 🧮 **公式识别**: 支持 OMML 和 LaTeX 格式数学公式
- 🖼️ **图片提取**: 自动提取试卷中的图片
- ✂️ **题目拆分**: 智能识别并拆分各个题目
- 📤 **导出功能**: 支持选择性导出题目到 Word

---

## 整体架构

```
┌─────────────────────────────────────────────────────────────────┐
│                        用户层 (Client)                           │
│  ┌─────────────┐                                                │
│  │   浏览器     │  ← 用户访问界面                                 │
│  └──────┬──────┘                                                │
└─────────┼───────────────────────────────────────────────────────┘
          │ HTTP
┌─────────┼───────────────────────────────────────────────────────┐
│         ▼                                                       │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │                    前端层 (Frontend)                     │    │
│  │  ┌─────────────────────────────────────────────────┐   │    │
│  │  │  Vue3 + TypeScript + Vite                       │   │    │
│  │  │  - 用户界面组件                                  │   │    │
│  │  │  - 公式渲染 (MathJax)                           │   │    │
│  │  │  - 状态管理 (Pinia)                             │   │    │
│  │  │  - UI 组件库 (Element Plus)                     │   │    │
│  │  └─────────────────────────────────────────────────┘   │    │
│  │  端口: 3000                                             │    │
│  └────────────────────────┬────────────────────────────────┘    │
└───────────────────────────┼─────────────────────────────────────┘
                            │ HTTP/REST API
┌───────────────────────────┼─────────────────────────────────────┐
│                           ▼                                       │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │                  API 层 (Node.js)                        │     │
│  │  ┌─────────────────────────────────────────────────┐   │     │
│  │  │  Express + TypeScript                           │   │     │
│  │  │  - 业务逻辑处理                                  │   │     │
│  │  │  - 文件上传/下载                                 │   │     │
│  │  │  - 服务间通信                                    │   │     │
│  │  │  - 数据验证                                      │   │     │
│  │  └─────────────────────────────────────────────────┘   │     │
│  │  端口: 8080                                             │     │
│  └────────────────────────┬────────────────────────────────┘     │
└───────────────────────────┼──────────────────────────────────────┘
                            │ HTTP/REST API
┌───────────────────────────┼──────────────────────────────────────┐
│                           ▼                                        │
│  ┌─────────────────────────────────────────────────────────┐      │
│  │                解析层 (Python)                           │      │
│  │  ┌─────────────────────────────────────────────────┐   │      │
│  │  │  FastAPI + Python 3.11                          │   │      │
│  │  │  - Word 文档解析                                 │   │      │
│  │  │  - 公式转换 (OMML → LaTeX)                      │   │      │
│  │  │  - 图片提取与处理                                │   │      │
│  │  │  - 题目拆分算法                                  │   │      │
│  │  └─────────────────────────────────────────────────┘   │      │
│  │  端口: 8001 (开发) / 8000 (生产)                        │      │
│  └─────────────────────────────────────────────────────────┘      │
└───────────────────────────────────────────────────────────────────┘
```

---

## 目录结构详解

### 根目录

```
math-demo/
├── apps/                          # 应用程序目录
│   ├── web/                       # 前端应用
│   ├── api/                       # Node.js API 服务
│   └── parser/                    # Python 解析服务
├── packages/                      # 共享包目录
│   └── shared-types/              # 共享类型定义
├── data/                          # 统一数据存储目录
│   ├── uploads/                   # 上传文件存储
│   ├── images/                    # 提取的图片存储
│   └── exports/                   # 导出文件存储
├── docker-compose.yml             # Docker Compose 配置
├── package.json                   # 根 package.json
├── pnpm-workspace.yaml           # pnpm workspace 配置
├── README.md                      # 项目说明
├── DEPLOYMENT.md                  # 部署说明
└── PROJECT_STRUCTURE.md           # 本文档
```

> **说明**: 所有数据文件（上传文件、图片、导出文件）统一存储在 `data/` 目录下，便于管理和备份。API 和 Parser 服务都访问同一数据位置。

---

## 模块说明

### 1. 前端应用 (apps/web)

```
web/
├── public/                        # 静态资源
│   └── favicon.ico
├── src/
│   ├── assets/                    # 资源文件
│   │   ├── styles/                # 全局样式
│   │   │   ├── main.scss
│   │   │   └── variables.scss
│   │   └── images/
│   ├── components/                # 组件目录
│   │   ├── MathRenderer/          # 公式渲染组件
│   │   │   ├── MathRenderer.vue   # MathJax 封装
│   │   │   └── index.ts
│   │   ├── QuestionList/          # 题目列表组件
│   │   │   ├── QuestionList.vue
│   │   │   └── index.ts
│   │   └── ExportPanel/           # 导出面板组件
│   │       ├── ExportPanel.vue
│   │       └── index.ts
│   ├── views/                     # 页面视图
│   │   ├── HomeView.vue           # 首页
│   │   └── UploadView.vue         # 上传页面
│   ├── router/                    # 路由配置
│   │   └── index.ts
│   ├── stores/                    # Pinia 状态管理
│   │   └── index.ts
│   ├── types/                     # 类型定义
│   │   └── index.ts
│   ├── utils/                     # 工具函数
│   │   └── request.ts             # HTTP 请求封装
│   ├── App.vue                    # 根组件
│   ├── main.ts                    # 入口文件
│   └── env.d.ts                   # 环境变量类型
├── .env.development               # 开发环境变量
├── .env.production                # 生产环境变量
├── index.html
├── package.json                   # @math-demo/web
├── tsconfig.json
├── tsconfig.node.json
└── vite.config.ts                 # Vite 配置
```

#### 核心组件说明

**MathRenderer (公式渲染)**
- 封装 MathJax 库
- 支持 LaTeX 公式渲染
- 支持行内公式和块级公式
- 自动加载 MathJax 库

**QuestionList (题目列表)**
- 展示解析后的题目列表
- 支持题目选择
- 显示题目类型、内容、选项
- 图片懒加载

**ExportPanel (导出面板)**
- 配置导出选项
- 选择包含答案/解析
- 设置水印
- 触发导出操作

---

### 2. API 服务 (apps/api)

```
api/
├── src/
│   ├── routes/                    # 路由模块
│   │   ├── upload.ts              # 上传/解析路由
│   │   ├── export.ts              # 导出路由
│   │   ├── health.ts              # 健康检查路由
│   │   └── images.ts              # 图片服务路由
│   ├── middleware/                # 中间件
│   │   └── errorHandler.ts        # 错误处理中间件
│   ├── utils/                     # 工具函数
│   │   └── logger.ts              # 日志工具
│   ├── types/                     # 类型定义
│   │   └── index.ts
│   └── index.ts                   # 入口文件
├── logs/                          # 日志文件 (运行时生成)
│   ├── combined.log               # 综合日志
│   └── error.log                  # 错误日志
├── .env                           # 环境变量
├── .env.example                   # 环境变量示例
├── package.json                   # @math-demo/api
├── tsconfig.json
└── Dockerfile                     # Docker 配置

> **注意**: API 服务的数据文件存储在统一的 `data/` 目录下，不在 api 目录内。
```

#### 路由模块说明

**upload.ts (上传路由)**
```typescript
// 主要接口
POST   /api/v1/upload              # 上传文件
GET    /api/v1/upload/progress/:id # 获取解析进度
GET    /api/v1/upload/questions/:id # 获取题目列表
```

**export.ts (导出路由)**
```typescript
// 主要接口
POST   /api/v1/export              # 导出题目
```

**health.ts (健康检查)**
```typescript
// 主要接口
GET    /api/v1/health              # 服务健康检查
```

**images.ts (图片服务)**
```typescript
// 主要接口
GET    /api/v1/images/:fileId/:imageName  # 获取图片文件
```

#### 中间件说明

**errorHandler (错误处理)**
- 捕获所有未处理的错误
- 统一错误响应格式
- 记录错误日志
- 区分开发和生产环境

---

### 3. 解析服务 (apps/parser)

```
parser/
├── app/
│   └── core/                      # 核心模块
│       ├── __init__.py
│       ├── parser.py              # Word 文档解析器
│       ├── splitter.py            # 题目拆分器
│       ├── omml2latex.py          # OMML 转 LaTeX
│       └── exporter.py            # Word 导出器
├── main.py                        # FastAPI 入口
├── requirements.txt               # Python 依赖
└── Dockerfile                     # Docker 配置

> **注意**: Parser 服务的数据文件存储在统一的 `data/` 目录下，不在 parser 目录内。
```

#### 核心模块说明

**parser.py (Word 解析器)**
```python
class DocxParser:
    """Word 文档解析器"""
    
    # 主要功能
    - 解析 docx 文件结构
    - 提取文本内容
    - 识别数学公式 (OMML/VML)
    - 提取图片资源
    - 处理表格和列表
    
    # 核心方法
    parse_document()      # 解析整个文档
    extract_paragraphs()  # 提取段落
    extract_formulas()    # 提取公式
    extract_images()      # 提取图片
```

**splitter.py (题目拆分器)**
```python
class QuestionSplitter:
    """题目拆分器"""
    
    # 主要功能
    - 识别题目编号
    - 判断题目类型
    - 提取选项
    - 分离答案和解析
    
    # 核心方法
    split_questions()     # 拆分题目
    identify_type()       # 识别题型
    extract_options()     # 提取选项
```

**omml2latex.py (公式转换)**
```python
# 主要功能
- OMML (Office Math Markup Language) 解析
- LaTeX 代码生成
- 支持复杂数学符号
- 处理上下标、分数、根号等

# 核心函数
convert_omml_to_latex()  # 转换函数
```

**exporter.py (Word 导出)**
```python
class WordExporter:
    """Word 导出器"""
    
    # 主要功能
    - 生成新的 Word 文档
    - 插入题目内容
    - 添加水印
    - 格式化输出
    
    # 核心方法
    export_questions()    # 导出题目
    add_watermark()       # 添加水印
```

#### API 接口

```python
# FastAPI 路由

@router.get("/health")
async def health_check():
    """健康检查"""

@router.post("/parse")
async def parse_document(file: UploadFile):
    """解析 Word 文档"""

@router.get("/images/{file_id}/{filename}")
async def get_image(file_id: str, filename: str):
    """获取图片"""

@router.post("/export")
async def export_questions(data: dict):
    """导出题目"""
```

---

### 4. 共享类型 (packages/shared-types)

```
shared-types/
├── src/
│   └── index.ts                   # 类型定义
├── dist/                          # 编译输出 (生成)
├── package.json                   # @math-demo/shared-types
└── tsconfig.json
```

#### 类型定义

```typescript
// 题目类型枚举
enum QuestionType {
  SINGLE_CHOICE = 'single_choice',
  MULTIPLE_CHOICE = 'multiple_choice',
  FILL_BLANK = 'fill_blank',
  ANSWER = 'answer',
  PROOF = 'proof',
  UNKNOWN = 'unknown'
}

// 题目接口
interface Question {
  id: string;
  number: number;
  type: QuestionType;
  type_name: string;
  content: string;
  content_html: string;
  options: Option[];
  answer?: string;
  analysis?: string;
  score?: number;
  difficulty?: string;
  images: string[];
  latex_formulas: string[];
}

// 其他类型...
// ParseProgress, UploadResponse, ExportRequest 等
```

---

## 数据流

### 上传解析流程

```
用户选择文件
    │
    ▼
┌─────────────┐
│   Web       │  1. 选择文件
│  (前端)     │  2. 显示上传进度
└──────┬──────┘
       │ POST /api/v1/upload
       │ FormData: file
       ▼
┌─────────────┐
│    API      │  1. 接收文件
│  (Node.js)  │  2. 保存到 uploads/
│             │  3. 转发到 Parser
└──────┬──────┘
       │ POST /parse
       │ FormData: file
       ▼
┌─────────────┐
│   Parser    │  1. 解析 Word 文档
│  (Python)   │  2. 提取公式和图片
│             │  3. 拆分题目
│             │  4. 保存图片到 uploads/images/
└──────┬──────┘
       │ JSON: {questions, total}
       ▼
┌─────────────┐
│    API      │  1. 接收解析结果
│  (Node.js)  │  2. 存储进度信息
│             │  3. 返回给前端
└──────┬──────┘
       │ JSON: {file_id, message}
       ▼
┌─────────────┐
│   Web       │  1. 接收 file_id
│  (前端)     │  2. 轮询进度
│             │  3. 显示题目列表
└─────────────┘
```

### 导出流程

```
用户选择题目
    │
    ▼
┌─────────────┐
│   Web       │  1. 选择题目
│  (前端)     │  2. 配置导出选项
└──────┬──────┘
       │ POST /api/v1/export
       │ JSON: {file_id, question_ids, options}
       ▼
┌─────────────┐
│    API      │  1. 接收导出请求
│  (Node.js)  │  2. 转发到 Parser
└──────┬──────┘
       │ POST /export
       │ JSON: {file_id, question_ids, options}
       ▼
┌─────────────┐
│   Parser    │  1. 读取原文件
│  (Python)   │  2. 提取选中题目
│             │  3. 生成新 Word
│             │  4. 添加水印
│             │  5. 保存到 exports/
└──────┬──────┘
       │ Stream: docx file
       ▼
┌─────────────┐
│    API      │  1. 接收文件流
│  (Node.js)  │  2. 转发给前端
└──────┬──────┘
       │ Stream: docx file
       ▼
┌─────────────┐
│   Web       │  1. 接收文件流
│  (前端)     │  2. 触发下载
└─────────────┘
```

---

## 技术栈详解

### 前端技术栈

| 技术 | 版本 | 用途 |
|-----|------|------|
| Vue 3 | ^3.3.8 | 前端框架 |
| TypeScript | ^5.2.2 | 类型安全 |
| Vite | ^5.0.0 | 构建工具 |
| Element Plus | ^2.4.4 | UI 组件库 |
| Pinia | ^2.1.7 | 状态管理 |
| Vue Router | ^4.2.5 | 路由管理 |
| MathJax | ^3.2.2 | 公式渲染 |
| Axios | ^1.6.2 | HTTP 客户端 |
| Sass | ^1.69.5 | CSS 预处理器 |

### API 技术栈

| 技术 | 版本 | 用途 |
|-----|------|------|
| Node.js | >= 18 | 运行环境 |
| Express | ^4.18.2 | Web 框架 |
| TypeScript | ^5.3.0 | 类型安全 |
| Multer | ^1.4.5 | 文件上传 |
| Axios | ^1.6.2 | HTTP 客户端 |
| Winston | ^3.11.0 | 日志记录 |
| UUID | ^9.0.1 | 唯一标识符 |
| tsx | ^4.7.0 | TypeScript 执行 |

### Parser 技术栈

| 技术 | 版本 | 用途 |
|-----|------|------|
| Python | >= 3.11 | 运行环境 |
| FastAPI | 0.109.0 | Web 框架 |
| Uvicorn | 0.27.0 | ASGI 服务器 |
| python-docx | 1.1.0 | Word 文档处理 |
| lxml | 4.9.3 | XML 处理 |
| Pillow | 10.2.0 | 图片处理 |
| python-multipart | 0.0.6 | 文件上传 |
| Pydantic | 2.5.0 | 数据验证 |

---

## 配置文件说明

### 根目录配置

**package.json**
```json
{
  "name": "math-demo-monorepo",
  "private": true,
  "scripts": {
    "dev": "pnpm -r --parallel dev",
    "build": "pnpm -r build",
    "docker:up": "docker-compose up -d"
  }
}
```

**pnpm-workspace.yaml**
```yaml
packages:
  - 'apps/*'
  - 'packages/*'
```

**docker-compose.yml**
```yaml
version: '3.8'
services:
  web:      # 前端服务
  api:      # API 服务
  parser:   # Parser 服务
```

### 前端配置

**vite.config.ts**
```typescript
export default defineConfig({
  plugins: [vue()],
  server: {
    port: 3000,
    proxy: {
      '/api': 'http://localhost:8080'
    }
  }
})
```

**tsconfig.json**
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "strict": true
  }
}
```

### API 配置

**tsconfig.json**
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "outDir": "./dist",
    "rootDir": "./src"
  }
}
```

**.env**
```env
PORT=8080
PARSER_URL=http://localhost:8001
LOG_LEVEL=info
```

### Parser 配置

**requirements.txt**
```
fastapi==0.109.0
uvicorn[standard]==0.27.0
python-docx==1.1.0
lxml==4.9.3
Pillow==10.2.0
```

---

## 开发规范

### 代码规范

1. **TypeScript 严格模式**: 所有项目启用 `strict: true`
2. **ESLint**: 统一代码风格
3. **Prettier**: 自动格式化代码
4. **命名规范**:
   - 组件: PascalCase (如 `MathRenderer.vue`)
   - 函数: camelCase (如 `parseDocument`)
   - 常量: UPPER_SNAKE_CASE
   - 类型: PascalCase (如 `QuestionType`)

### 目录规范

1. **按功能组织**: 相关文件放在同一目录
2. **索引文件**: 每个目录提供 `index.ts` 导出
3. **类型分离**: 类型定义单独放在 `types/` 目录
4. **工具函数**: 通用工具放在 `utils/` 目录

### Git 规范

1. **分支管理**:
   - `main`: 主分支
   - `develop`: 开发分支
   - `feature/*`: 功能分支
   - `hotfix/*`: 紧急修复

2. **提交信息**:
   ```
   feat: 新功能
   fix: 修复
   docs: 文档
   style: 格式
   refactor: 重构
   test: 测试
   chore: 构建
   ```

### 文档规范

1. **代码注释**: 复杂逻辑必须注释
2. **README**: 每个项目提供 README
3. **API 文档**: 使用 Swagger/OpenAPI
4. **变更日志**: 记录版本变更

---

## 扩展指南

### 添加新功能

1. **前端组件**:
   ```
   apps/web/src/components/NewComponent/
   ├── NewComponent.vue
   └── index.ts
   ```

2. **API 接口**:
   ```
   apps/api/src/routes/newRoute.ts
   ```

3. **Parser 模块**:
   ```
   apps/parser/app/core/new_module.py
   ```

### 添加新依赖

1. **前端**:
   ```bash
   cd apps/web
   npm install <package>
   ```

2. **API**:
   ```bash
   cd apps/api
   npm install <package>
   ```

3. **Parser**:
   ```bash
   cd apps/parser
   pip install <package>
   # 添加到 requirements.txt
   ```

---

## 性能优化

### 前端优化

1. **代码分割**: 路由懒加载
2. **图片优化**: 懒加载、压缩
3. **缓存策略**: 合理设置 HTTP 缓存
4. **构建优化**: Tree shaking、代码压缩

### API 优化

1. **压缩响应**: 启用 gzip
2. **连接池**: 复用 HTTP 连接
3. **缓存**: 缓存频繁请求的数据
4. **流式传输**: 大文件使用流

### Parser 优化

1. **异步处理**: 解析过程异步化
2. **资源复用**: 复用文档对象
3. **图片缓存**: 缓存已提取的图片
4. **并发控制**: 限制并发解析数量

---

## 安全考虑

1. **文件上传**:
   - 限制文件类型 (仅 .docx)
   - 限制文件大小
   - 扫描恶意文件

2. **API 安全**:
   - 输入验证
   - 防 SQL 注入
   - 防 XSS 攻击
   - 速率限制

3. **数据安全**:
   - 敏感信息加密
   - 定期清理临时文件
   - 访问控制

---

## 监控与日志

### 日志级别

- **ERROR**: 错误信息
- **WARN**: 警告信息
- **INFO**: 一般信息
- **DEBUG**: 调试信息

### 监控指标

- 请求响应时间
- 错误率
- 服务可用性
- 资源使用率

---

**文档版本**: 1.0.0
**最后更新**: 2026-03-04
**维护者**: Math Demo Team
