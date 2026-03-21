# 数学试卷解析系统 - Monorepo

基于 pnpm workspace 的混合架构数学试卷解析系统。

## 项目架构

```
math-demo/
├── apps/                          # 应用目录
│   ├── web/                       # 前端应用 (Vue3 + TypeScript)
│   ├── api/                       # Node.js API 服务
│   └── parser/                    # Python 解析服务
├── packages/                      # 共享包
│   └── shared-types/              # 共享类型定义
├── docker-compose.yml             # Docker Compose 配置
├── package.json                   # 根 package.json
├── pnpm-workspace.yaml           # pnpm workspace 配置
└── README.md                      # 项目说明
```

## 技术栈

### 前端 (apps/web)
- **框架**: Vue 3 + TypeScript
- **构建工具**: Vite
- **UI 组件库**: Element Plus
- **状态管理**: Pinia
- **公式渲染**: MathJax
- **容器**: Nginx

### API 服务 (apps/api)
- **运行时**: Node.js 18
- **框架**: Express + TypeScript
- **包管理器**: pnpm
- **日志**: Winston
- **HTTP 客户端**: Axios

### 解析服务 (apps/parser)
- **运行时**: Python 3.11
- **框架**: FastAPI
- **文档解析**: python-docx, lxml
- **公式处理**: 自定义 OMML 转 LaTeX
- **容器**: Python Slim

## 快速开始

### 环境要求
- Node.js >= 18
- Python >= 3.11
- pnpm >= 8.0
- Docker & Docker Compose (可选)

### 安装依赖

```bash
# 安装 pnpm (如果尚未安装)
npm install -g pnpm

# 安装所有依赖
pnpm install
```

### 开发模式

```bash
# 同时启动所有服务
pnpm dev

# 或者分别启动
pnpm dev:web    # 前端: http://localhost:5173
pnpm dev:api    # API: http://localhost:8080
pnpm dev:parser # Parser: http://localhost:8000
```

### Docker 部署

```bash
# 构建并启动所有服务
pnpm docker:up

# 停止服务
pnpm docker:down

# 重新构建
pnpm docker:build
```

## 服务说明

### 前端服务 (Web)
- **端口**: 3000 (Docker) / 5173 (开发)
- **职责**: 用户界面、题目展示、导出配置
- **依赖**: API 服务

### API 服务 (Node.js)
- **端口**: 8080
- **职责**: 
  - 业务逻辑处理
  - 文件上传/下载
  - 调用 Parser 服务
  - 数据持久化 (未来扩展)
- **依赖**: Parser 服务

### 解析服务 (Python)
- **端口**: 8000
- **职责**:
  - Word 文档解析
  - 公式提取 (OMML → LaTeX)
  - 图片提取和处理
  - 题目拆分
- **依赖**: 无

## 工作流程

```
用户上传试卷
    ↓
[Web] → 选择文件
    ↓
[API] → 接收文件 → 转发到 Parser
    ↓
[Parser] → 解析 Word → 提取公式/图片 → 拆分题目
    ↓
[API] ← 返回题目数据 ← [Parser]
    ↓
[Web] ← 展示题目 ← [API]
    ↓
用户选择题目 → 导出
    ↓
[API] → 调用 Parser 生成 Word
    ↓
用户下载
```

## 目录详解

### apps/web/
```
web/
├── src/
│   ├── components/          # 组件
│   │   ├── MathRenderer/    # 公式渲染组件
│   │   ├── QuestionList/    # 题目列表
│   │   └── ExportPanel/     # 导出面板
│   ├── views/               # 页面
│   ├── router/              # 路由
│   ├── stores/              # Pinia 状态
│   ├── types/               # 类型定义
│   └── utils/               # 工具函数
├── package.json
├── vite.config.ts
└── Dockerfile
```

### apps/api/
```
api/
├── src/
│   ├── routes/              # 路由
│   │   ├── upload.ts        # 上传/解析
│   │   ├── export.ts        # 导出
│   │   └── health.ts        # 健康检查
│   ├── middleware/          # 中间件
│   ├── utils/               # 工具函数
│   └── types/               # 类型定义
├── package.json
├── tsconfig.json
└── Dockerfile
```

### apps/parser/
```
parser/
├── app/
│   └── core/
│       ├── parser.py        # Word 解析器
│       ├── splitter.py      # 题目拆分器
│       ├── omml2latex.py    # OMML 转 LaTeX
│       └── exporter.py      # Word 导出器
├── main.py                  # FastAPI 入口
├── requirements.txt
└── Dockerfile
```

### packages/shared-types/
```
shared-types/
├── src/
│   └── index.ts             # 共享类型定义
├── package.json
└── tsconfig.json
```

## 脚本命令

```bash
# 根目录
pnpm dev              # 启动所有服务
pnpm build            # 构建所有应用
pnpm lint             # 运行所有 lint
pnpm clean            # 清理所有 node_modules
pnpm test             # 运行所有测试

# 子项目
pnpm --filter @math-demo/web dev      # 只启动前端
pnpm --filter @math-demo/api dev      # 只启动 API
pnpm --filter @math-demo/parser dev   # 只启动 Parser

# Docker
pnpm docker:up        # 启动 Docker 服务
pnpm docker:down      # 停止 Docker 服务
pnpm docker:build     # 构建 Docker 镜像
```

## API 文档

启动服务后访问:
- API 文档: http://localhost:8080/api/v1/health
- Parser 文档: http://localhost:8000/docs

## 环境变量

### API 服务
```env
PORT=8080
NODE_ENV=development
PARSER_URL=http://localhost:8000
LOG_LEVEL=info
```

### 前端
```env
VITE_API_URL=http://localhost:8080
```

## 贡献指南

1. 使用 pnpm 管理依赖
2. 遵循各项目的代码规范
3. 提交前运行 `pnpm lint`
4. 保持类型定义同步更新

## 许可证

MIT
