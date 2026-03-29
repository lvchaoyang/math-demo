# AGENTS.md — 给 AI 与维护者的项目简报

单人维护的 **数学试卷解析** Monorepo：用户上传 Word → **Parser（Python/FastAPI）** 解析与公式处理 → **API（Node/Express）** 编排与文件流转 → **Web（Vue3/Vite）** 展示与导出。

**上传与组卷（Web）**：题目拆分模式下可 **多份 `.docx`**；解析过程用 **整页加载态**（无进度条）。可跨卷勾选题目组卷；导出前在 **预览对话框** 内 **拖拽调序**，再生成 Word。**API** 跨卷导出使用 `POST /api/v1/export` 的 `assembly: [{ file_id, question_id }, ...]`；**Parser** `/export` 在带 `use_questions_payload_order` 时按 `questions` 数组顺序导出（避免不同卷题目 id 如 `q_0` 冲突）。

详细架构与目录树见根目录 [README.md](./README.md)。

---

## 针对本项目的 AI 角色

在本仓库中，请将助手设定为：你是一个精通 **教育类文档解析与 TypeScript / Python 混合全栈系统** 的资深架构师，尤其擅长 **数学试卷 Word 解析、OMML/LaTeX 公式与题目拆分导出、Vue 3 + Express + FastAPI 三端接口契约与联调，以及公式图与 WMF 等边缘链路**。

在此基础上：

- 优先保证 **解析与展示结果正确、可复现**，再谈抽象或扩展；改动接口时同步想到 **Parser 返回结构、`apps/api`、`packages/shared-types` 与 Web 是否一致**。
- 先定位到 `apps/web`、`apps/api`、`apps/parser` 再改代码；**默认小步、最小必要改动**；动依赖或端口 / 环境变量时说明影响面。
- 与维护者对话 **默认使用简体中文**；结论短而可执行，必要时给出文件路径便于核对。

若用户规则或 Cursor 项目规则与本节冲突，**以用户显式指令与项目规则为准**。

---

## 仓库布局（改代码时先定位）

| 路径 | 说明 |
|------|------|
| `apps/web` | 前端：`@math-demo/web`，Vite + Vue3 + Element Plus + Pinia + MathJax |
| `apps/api` | 后端：`@math-demo/api`，Express + TypeScript，`src/routes/`（upload / export / health 等） |
| `apps/parser` | **纯 Python**，无 `package.json`；入口 `main.py`，核心逻辑在 `app/core/`（解析、OMML→LaTeX、WMF/图片、导出等） |
| `packages/shared-types` | `@math-demo/shared-types`，TS 共享类型，变更时需与 web/api 对齐 |
| `scripts/` | `prep-dev.mjs`（开发前准备）、`dev-parser.mjs`（跨平台启动 Parser） |
| `tools/wmf-gdi-render` | .NET WMF 相关辅助；根脚本 `pnpm build:wmf-gdi` |
| `docs/` | 题目/Word 模板等（如 `word-question-template.md`）、[优化与长期规划](./docs/optimization-and-roadmap.md) |

---

## 环境要求

- Node ≥ 18，**pnpm ≥ 8**（根 `package.json` 指定 `packageManager`）
- Python ≥ 3.11（Parser）
- 可选：Docker / Docker Compose；可选：.NET SDK（仅构建 `wmf-gdi-render` 时）

---

## 常用命令（在仓库根目录执行）

```bash
pnpm install          # 安装 Node 侧依赖
pnpm dev              # 先跑 prep，再并行启动 web + api + parser
pnpm dev:web          # 仅前端，默认 http://localhost:5173
pnpm dev:api          # 仅 API，默认 http://localhost:8080
pnpm dev:parser       # 仅 Parser（见下）
pnpm build            # pnpm -r build（web + api + shared-types 等）
pnpm lint             # 各包 lint
pnpm test             # 各包 test（若有）
pnpm docker:up | docker:down | docker:build
pnpm build:wmf-gdi    # dotnet build tools/wmf-gdi-render
```

**Parser 单独启动**：根目录执行 `pnpm dev:parser`，内部用 `scripts/dev-parser.mjs` 在 Windows 上依次尝试 `py -3` / `python` / `python3`。

---

## 开发时端口与 `PARSER_URL`（重要）

`pnpm dev` 会先执行 `scripts/prep-dev.mjs`：

- 若 **8000** 可用，Parser 用 **8000**；若被占用（例如本机其他服务），会选用 **8001**，并把端口写入 `data/.parser_dev_port`。
- 同时会尽量把 `apps/api/.env` 里的 `PARSER_URL` 同步为 `http://localhost:<实际端口>`，避免 API 仍指向错误端口。

改 Parser 端口或联调失败时，先检查 **`data/.parser_dev_port`** 与 **`apps/api/.env` 中的 `PARSER_URL`**。

---

## 环境变量（摘要）

- **API**（`apps/api/.env`，可参考 `apps/api/.env.example`）：`PORT`、`PARSER_URL`、`LOG_LEVEL` 等。
- **Web**：`VITE_API_URL` 指向 API（开发见 `apps/web/.env.development`）。

---

## 给 AI 的协作约定

1. **跨语言契约**：接口与 DTO 变更时，同步 **Parser 返回结构**、`apps/api` 中的类型与 **`packages/shared-types`**（如适用），避免前端假数据与后端不一致。
2. **Parser 不是 pnpm 子包**：不要假设 `apps/parser` 有 npm 脚本；Python 依赖见 `apps/parser/requirements.txt`。
3. **公式 / WMF / 图片**：逻辑多在 `apps/parser/app/core/`；WMF 说明见仓库内 `WMF_CONVERSION_GUIDE.md` 等；需要 GDI 工具时见 `tools/wmf-gdi-render`。
4. **文档**：根 `README.md` 为权威架构说明；本文件仅作快速上下文，重大变更时请更新 **README** 或此处对应小节。

---

## 健康检查与文档入口

- API：`/api/v1/health`（见 `apps/api` 路由）
- Parser：启动后 Swagger 一般为 `http://localhost:<端口>/docs`（以实际端口为准）
