# 数学试卷解析系统 — 部署说明（小白向）

本文面向 **第一次部署** 的同学：尽量少术语，按顺序复制命令即可。若你已熟悉 Node / Docker，可直接看 [进阶：生产与 Docker](#进阶生产与-docker) 或目录跳转。

---

## 目录

1. [5 分钟看懂：这套系统在跑什么](#5-分钟看懂这套系统在跑什么)
2. [开始前：你需要安装什么](#开始前你需要安装什么)
3. [服务器要求与建议（正式部署）](#服务器要求与建议正式部署)（含 WMF 与 Linux / Windows）
4. [⭐ 推荐：本机开发一键跑起来](#-推荐本机开发一键跑起来)
5. [进阶：分开启动三个服务（调试用）](#进阶分开启动三个服务调试用)
6. [进阶：生产与 Docker](#进阶生产与-docker)
7. [环境变量说明（白话版）](#环境变量说明白话版)
8. [端口与联调（很重要）](#端口与联调很重要)
9. [怎么确认部署成功](#怎么确认部署成功)
10. [常见问题（按现象找）](#常见问题按现象找)
11. [附录：Nginx / PM2 / 备份](#附录nginx--pm2--备份)

---

## 5 分钟看懂：这套系统在跑什么

可以把系统想成 **三个互相打电话的程序**：

| 角色 | 是什么 | 平时干什么 |
|------|--------|------------|
| **前端（Web）** | 网页界面 | 你在浏览器里上传试卷、看题目、导出 |
| **API** | Node.js 中间层 | 收文件、转给解析服务、把结果返回给网页 |
| **Parser** | Python 解析服务 | 读 Word、拆题目、转公式等「重活」 |

**正常使用时**：你只打开浏览器访问前端；前端会把请求发给 API，API 再去找 Parser。**三个都要启动**，页面才能完整可用。

---

## 开始前：你需要安装什么

### 必装（本机开发）

| 软件 | 版本要求 | 用来干什么 | 下载提示 |
|------|----------|------------|----------|
| **Node.js** | ≥ 18 | 跑前端和 API | [nodejs.org](https://nodejs.org/) 选 LTS |
| **pnpm** | ≥ 8 | 安装前端/API 依赖（本仓库标准） | 装好 Node 后执行：`npm install -g pnpm` |
| **Python** | ≥ 3.11 | 跑 Parser | [python.org](https://www.python.org/)，安装时勾选 **Add to PATH**（Windows） |
| **Git** | 任意 | 拉代码 | [git-scm.com](https://git-scm.com/) |

### 可选

- **Docker Desktop**：想用容器一键起服务时再装（见 [Docker](#方式-bdocker适合想少装依赖的同学)）。
- **Pandoc**：部分「整页 HTML」能力依赖系统 Pandoc；没有时 Parser 可能降级或提示，按运行日志处理即可。

### Windows 特别提醒

- 打开 **PowerShell** 或 **cmd** 执行下文命令即可。
- Python 若提示 `python` 不可用，可尝试 **`py -3`**（安装器自带启动器）。
- 若 8000 端口被占用（常见打印控件等），本项目会自动改用 **8001**，并写配置文件（见下文 [端口与联调](#端口与联调很重要)）。

---

## 服务器要求与建议（正式部署）

本节说的是：**把系统部署到云主机、学校机房、公司服务器** 时，机器大概要什么配置、注意什么。个人电脑本地开发一般不用按「服务器标准」买机器，能跑 Node + Python 即可。

### 这套服务在服务器上要跑什么？

与 [前面说的三个程序](#5-分钟看懂这套系统在跑什么)一致：同一台机器上通常会同时跑 **API（Node）**、**Parser（Python）**，再加上 **Nginx** 提供网页静态文件和反向代理（推荐做法）。因此内存和磁盘要按「三个进程 + 临时文件」来留余量。

### 硬件：最低配置 vs 建议配置

下面按 **小规模使用**（例如教研组、几十人以内、偶发高峰）估算；若同时很多人上传大试卷，需要 **成比例加内存、磁盘**，并考虑把 Parser 单独放到更强的机器或做扩容（需自行架构）。

| 项目 | 最低可跑（体验一般） | 建议配置（更稳妥） | 说明 |
|------|----------------------|---------------------|------|
| **CPU** | 2 核 | 4 核及以上 | Word 解析、公式/图片处理偏 CPU；核多更不容易卡顿 |
| **内存** | 4 GB | **8 GB 及以上** | Parser 处理 docx、图片时占用明显；4G 易在并发或大单文件时吃紧 |
| **系统盘** | 40 GB（SSD 更佳） | 80 GB+ SSD | 系统、依赖、日志、临时文件 |
| **数据盘（可选）** | — | 单独挂载、100 GB+ | 上传的 Word、解析产生的图片/缓存、导出文件会持续增长，建议单独盘并做好备份 |
| **交换分区 / Swap** | 建议开启 | 2～4 GB | 内存紧张时可避免进程直接被系统杀掉的概率（速度会变慢） |

**磁盘空间心里要有数**：每份试卷解析可能产生 **原文件 + 图片/缓存**；若长期不清理，`data/`（或你挂载的上传目录）会越来越大，建议定期备份后做清理策略。

### 操作系统建议

| 场景 | 建议 |
|------|------|
| **生产环境（常用）** | **64 位 Linux**（如 **Ubuntu 22.04 LTS**、**Debian 12**）仍可作为默认推荐：Web/API/Parser 都可运行；**WMF 公式图**在 Linux 上走 **Inkscape / ImageMagick 等跨平台工具**（见下一小节），不是「没有 Windows 就不能用」。 |
| **Windows Server** | 适合希望 **统一用仓库里的 GDI+ 小工具**（`WmfGdiRender.exe`）处理 WMF/EMF、追求与 Word 更一致画质的场景；命令行与服务托管与 Linux 不同，文档示例以 Linux 为主。 |
| **macOS** | 适合开发，一般不作为机房长期服务器。 |

### WMF / 公式图：Windows 专用程序有没有考虑？

**有考虑。** 仓库里的 **`tools/wmf-gdi-render`**（`WmfGdiRender.exe`）是 **仅在 Windows 上** 用系统 **GDI+** 把部分 **WMF/EMF** 栅格成 PNG 的辅助程序，观感上往往 **更接近 Word**。

**Parser（Python）里的逻辑大致是：**

- 在 **Windows** 上：若能找到已编译的 `WmfGdiRender.exe`（或通过环境变量指定路径），会 **优先走这条 GDI+ 路径**。
- 在 **Linux / macOS** 上：**不会**使用 `WmfGdiRender.exe`，会自动尝试系统里的 **Inkscape**、**ImageMagick（`magick`/`convert`）** 等工具做 WMF 转换；检测顺序见 `apps/parser/app/core/wmf_converter.py`。

因此：

| 你的部署方式 | 公式图（WMF）方面怎么理解 |
|--------------|---------------------------|
| **Linux 服务器** | **完全可以部署**；请在服务器上 **安装至少一种** Parser 能调用的工具（常见：**Inkscape** 或 **ImageMagick**）。例如 Ubuntu：`sudo apt install -y inkscape` 或 `sudo apt install -y imagemagick`。若未安装任何转换工具，部分 WMF 可能无法转成预览图，日志里会有告警。 |
| **Windows 服务器** | 可 **`pnpm build:wmf-gdi`**（需 **.NET SDK**）编译 `WmfGdiRender.exe`，Parser/API 在 Windows 上会自动探测；也可在 **`apps/api/.env`** 中设置 **`WMF_GDI_RENDER_EXE`** 指向 exe 绝对路径（见 `apps/api/.env.example`）。 |
| **混合部署（进阶）** | 例如：**Web + API 在 Linux**，**Parser 单独跑在一台 Windows**（`PARSER_URL` 指向该机）。线上入口仍是 Linux，WMF 重活走 Windows GDI+；需自行处理网络与安全。 |

**API（Node）** 在 **Windows** 且存在上述 exe 时，对部分图片请求也会走 GDI；在 **Linux** 上不会调用该 exe。

更细的说明见 **`tools/wmf-gdi-render/CURSOR_HANDOFF.md`**。

### 网络与端口（防火墙）

- 对访客通常只开放 **80（HTTP）**、**443（HTTPS）**；由 Nginx 统一入口。
- **API、Parser** 建议只监听 **本机回环（127.0.0.1）或内网**，不要直接暴露到公网，由 Nginx 反代 `/api`。
- 解析、导出大文件时请求时间较长，反向代理要配 **较大的超时时间**（见 [附录](#附录nginx--pm2--备份) 中的 `proxy_read_timeout` 思路）。

### 用 Docker 部署时

- 除上述 CPU/内存外，Docker 自身有少量开销；**建议内存仍按 8 GB 档**规划更省心。
- 给容器挂载的 **数据卷** 所在磁盘要够大，并纳入备份。

### 什么时候该加配置或拆服务？

出现下面情况时，说明单机可能不够用了，需要 **加内存/磁盘**，或把 **Parser 单独部署**、多实例等（超出本文范围，需按访问量设计）：

- 多人同时上传，Parser 经常 **CPU 打满、内存飙高**；
- 磁盘 **告警** 或 `data` 目录暴涨；
- 接口频繁 **超时**（已与 `EXPORT_TIMEOUT_MS`、Nginx 超时核对过仍不够）。

---

## ⭐ 推荐：本机开发一键跑起来

下面所有命令都在 **项目根目录** `math-demo` 下执行（即包含 `package.json`、`apps` 文件夹的那一层）。

### 第 1 步：拿到代码

```bash
git clone <你的仓库地址>
cd math-demo
```

没有 Git 也可以下载 ZIP 解压，再 `cd` 进解压目录。

### 第 2 步：安装 Node 依赖（整仓一次即可）

```bash
pnpm install
```

若提示没有 `pnpm`，先执行：`npm install -g pnpm`。

### 第 3 步：配置 API 环境变量

1. 进入文件夹 `apps/api`。
2. 若有 `apps/api/.env.example`，**复制一份**改名为 `.env`（没有 `.example` 就自己新建 `.env`）。
3. 用记事本 / VS Code 打开 `apps/api/.env`，至少保证有类似内容（端口按你机器情况改）：

```env
PORT=8080
NODE_ENV=development
LOG_LEVEL=info
PARSER_URL=http://localhost:8000
```

> **说明**：`PARSER_URL` 必须和 Parser 实际监听地址一致。跑 `pnpm dev` 时脚本可能自动改成 `8001`，一般 **不用手改**；若联调失败再对照 [端口与联调](#端口与联调很重要)。

### 第 4 步：安装 Python 依赖（Parser）

在项目根目录执行：

```bash
cd apps/parser
pip install -r requirements.txt
```

若你习惯虚拟环境（推荐）：

```bash
cd apps/parser
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
# source .venv/bin/activate
pip install -r requirements.txt
```

### 第 5 步：一条命令启动三个服务

回到 **项目根目录** `math-demo`，执行：

```bash
pnpm dev
```

第一次运行前会自动执行 **端口检测**，并尽量把 `apps/api/.env` 里的 `PARSER_URL` 改成当前 Parser 端口（8000 或 8001）。

### 第 6 步：用浏览器打开

- 前端地址以终端里 Vite 输出为准；本仓库 **`apps/web/vite.config.ts` 默认开发端口为 `3000`**，一般为：  
  **http://localhost:3000**
- 在页面里试：**上传试卷** → 若能看到解析进度或题目列表，说明三条链路基本打通。

### 第 7 步：停掉服务

在运行 `pnpm dev` 的终端里按 **`Ctrl + C`**。

---

## 进阶：分开启动三个服务（调试用）

适合：**只想重启其中一个**、或 **`pnpm dev` 报错要排查**。

仍然在项目根目录装好依赖的前提下，开 **三个终端**，分别执行：

**终端 1 — Parser（Python）**

```bash
pnpm dev:parser
```

（内部会尝试 `py -3` / `python` / `python3`，见 `scripts/dev-parser.mjs`。）

**终端 2 — API**

```bash
pnpm dev:api
```

**终端 3 — 前端**

```bash
pnpm dev:web
```

然后浏览器访问 **http://localhost:3000**（以 Vite 提示为准）。

---

## 进阶：生产与 Docker

### 方式 A：本机构建后部署（概念流程）

1. 根目录执行 **`pnpm install`** → **`pnpm build`**（会构建 web、api 等包，见各子包脚本）。
2. **API**：运行 `apps/api` 构建产物（如 `node dist/index.js`），并配置环境变量 `NODE_ENV=production`、`PARSER_URL` 指向可访问的 Parser 地址。
3. **Parser**：在生产机用 `uvicorn` 或进程守护运行 `apps/parser/main.py`，保证 API 能访问到该地址。
4. **前端**：将 `apps/web/dist` 交给 **Nginx** 或其它静态服务器；接口路径需与生产环境一致（常见做法是同域反代 `/api`）。

具体 Nginx 示例、PM2、systemd 等见下文 [附录](#附录nginx--pm2--备份)。

### 方式 B：Docker（适合想少装依赖的同学）

**前提**：已安装 Docker，并支持 `docker compose`（或旧版 `docker-compose`）。

在项目根目录：

```bash
# 构建镜像
pnpm docker:build
# 或：docker compose build

# 后台启动
pnpm docker:up
# 或：docker compose up -d

# 看日志
docker compose logs -f

# 停止
pnpm docker:down
```

当前仓库的 `docker-compose.yml` 会映射 **8080（API）、8000（Parser）、3000（Web 容器内 Nginx）** 等，请以该文件为准。更新代码后常用：

```bash
git pull
pnpm docker:build
pnpm docker:up
```

> 容器内路径、数据卷与 **本机 `data/` 目录开发模式** 可能不一致；以 `docker-compose.yml` 与各 `Dockerfile` 为准。

---

## 环境变量说明（白话版）

### API（`apps/api/.env`）

| 变量 | 含义 | 小白注意 |
|------|------|----------|
| `PORT` | API 监听端口 | 默认 8080；改了要同步改前端代理或网关 |
| `PARSER_URL` | Parser 的根地址 | **必须**与 Parser 实际端口一致，例如 `http://localhost:8000` |
| `LOG_LEVEL` | 日志详细程度 | `info` / `debug` / `warn` 等 |
| `EXPORT_TIMEOUT_MS` | 导出等待 Parser 的最长时间（毫秒） | 大试卷可调大，见 `.env.example` 注释 |

### 前端开发（`apps/web/.env.development`）

一般由仓库自带；开发时通过 Vite 把 `/api` 代理到本机 API（见 `apps/web/vite.config.ts`）。若你改了 API 端口，需对照修改代理或 `VITE_API_URL`。

### Parser

常用环境变量：`PARSER_PORT`（若代码支持）、`PYTHONUNBUFFERED=1`（Docker 里方便看日志）。默认端口多为 **8000**。

---

## 端口与联调（很重要）

| 服务 | 开发时常见端口 | 说明 |
|------|----------------|------|
| 前端（Vite） | **3000** | 见 `apps/web/vite.config.ts` |
| API | **8080** | 见 `apps/api/.env` 的 `PORT` |
| Parser | **8000** 或 **8001** | 8000 被占用时会换 8001 |

本项目在 **`pnpm dev` 前** 会运行 `scripts/prep-dev.mjs`：

- 在仓库根目录写入 **`data/.parser_dev_port`**，里面是 Parser 实际使用的端口数字。
- 并尽量把 **`apps/api/.env` 里的 `PARSER_URL`** 改成 `http://localhost:该端口`。

**若页面上传后一直转圈 / 解析失败：**

1. 打开 `data/.parser_dev_port` 看是 `8000` 还是 `8001`。
2. 打开 `apps/api/.env`，确认 `PARSER_URL` 是否为 `http://localhost:同上端口`。
3. 修改 `.env` 后 **重启 API**（或整仓 `pnpm dev`）。

---

## 怎么确认部署成功

在三个服务都已启动的前提下：

1. **浏览器** 能打开前端首页与上传页。
2. **API 健康检查**（终端执行）：

```bash
curl http://localhost:8080/api/v1/health
```

应返回 JSON，且含成功状态（具体字段以实际响应为准）。

3. **Parser 健康检查**（端口换成你的，例如 8000）：

```bash
curl http://localhost:8000/health
```

4. **上传一份小 `.docx`**，看是否出现题目列表或 HTML 预览。

> 说明：仓库中 **没有** `test_monorepo.py` 集成脚本时，以上四步即为手动验收标准。

---

## 常见问题（按现象找）

### 1. `pnpm` 不是内部或外部命令

- 先安装 Node.js，再执行：`npm install -g pnpm`，**重新打开**终端。

### 2. `pnpm install` 报错或很慢

- 检查 Node 版本是否 ≥ 18。
- 公司网络若有限制，需配置 npm/pnpm 镜像（自行搜索「pnpm 镜像源」）。

### 3. Parser 起不来 / `ModuleNotFoundError`

- 确认在 `apps/parser` 下执行过：`pip install -r requirements.txt`。
- Windows 可尝试用 **`py -3`** 调用 Python。

### 4. 上传后没反应或立刻失败

- 对照 [端口与联调](#端口与联调很重要) 检查 `PARSER_URL` 与 `data/.parser_dev_port`。
- 看运行 `pnpm dev` 的终端里 **API、Parser** 是否报错。

### 5. 端口被占用（`EADDRINUSE`）

- Windows：`netstat -ano | findstr :8080` 查 PID，任务管理器结束对应进程。
- 或修改 `apps/api/.env` 的 `PORT`，并同步改前端代理 / 环境变量。

### 6. 生产环境上传大文件失败

- **Nginx** 增加：`client_max_body_size 100M;`（或按需更大）。
- 确认反向代理 **超时时间** 足够（解析大 Word 可能较慢）。

### 7. Docker 构建失败

- 确认 Docker 有足够磁盘空间。
- 查看 `docker compose build` 完整日志；有时是网络拉不下基础镜像，需重试或配置镜像加速。

---

## 附录：Nginx / PM2 / 备份

### 本机开发时数据写在哪里？

解析产生的上传文件、图片缓存等多在仓库根目录 **`data/`** 下（如 `data/uploads`、`data/images` 等，以实际代码为准）。**部署到服务器时请做好备份与磁盘空间规划。**

### Nginx 反向代理（示例思路）

- 静态资源：`root` 指向 `apps/web/dist`；`try_files` 支持 Vue Router。
- `/api`（或 `/api/v1`）反代到 `http://127.0.0.1:8080`，注意 **`proxy_read_timeout`** 宜放大（解析/导出较慢）。
- 与 `apps/web` 生产构建时的 `VITE_API_URL`、是否同域有关，需一套配置一致。

示例片段（**需按你的域名与路径改**）：

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        root /var/www/math-demo/apps/web/dist;
        try_files $uri $uri/ /index.html;
    }

    location /api/ {
        proxy_pass http://127.0.0.1:8080;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 600s;
        client_max_body_size 100M;
    }
}
```

### PM2 托管 API（示例）

```bash
npm install -g pm2
cd /var/www/math-demo/apps/api
pm2 start dist/index.js --name math-demo-api
pm2 save
pm2 startup
```

Parser 可用 **systemd** 或 Docker 单独跑；核心是 **API 的 `PARSER_URL` 能访问到 Parser**。

### 备份建议

定期备份：

- 仓库配置与 `.env`（勿泄露密钥）
- **`data/`** 或生产环境等价数据目录
- 若有独立上传目录、导出目录，一并纳入备份策略

---

## 安全与维护（简表）

| 建议 | 说明 |
|------|------|
| HTTPS | 公网访问务必配证书（如 Let’s Encrypt） |
| 防火墙 | 只开放 80/443（及 SSH），不要直接把数据库、Parser 暴露公网（除非你有专门方案） |
| 更新 | `git pull` → `pnpm install` → `pnpm build` → 重启进程/容器 |
| 日志 | API、Nginx、容器日志定期查看 |

---

## 更多文档

- 架构与目录：**[README.md](./README.md)**
- 给维护者/AI 的简报：**[AGENTS.md](./AGENTS.md)**
- Windows GDI+ 与 WMF/EMF 工具说明：**[tools/wmf-gdi-render/CURSOR_HANDOFF.md](./tools/wmf-gdi-render/CURSOR_HANDOFF.md)**

---

**最后更新**：2026-03-29（含服务器配置建议；Monorepo：`pnpm`、Vite 端口、`prep-dev` 与 `data/.parser_dev_port`）
