# 数学试卷解析系统 - 详细部署说明

## 目录

1. [环境要求](#环境要求)
2. [开发环境部署](#开发环境部署)
3. [生产环境部署](#生产环境部署)
4. [Docker 部署](#docker-部署)
5. [服务配置](#服务配置)
6. [常见问题](#常见问题)
7. [监控与日志](#监控与日志)

---

## 环境要求

### 必需软件

| 软件 | 版本 | 说明 |
|-----|------|------|
| Node.js | >= 18.0.0 | 前端和 API 服务运行环境 |
| Python | >= 3.11 | Parser 服务运行环境 |
| pnpm | >= 8.0.0 | 包管理器 |
| Git | 任意 | 代码版本控制 |

### 可选软件

| 软件 | 版本 | 说明 |
|-----|------|------|
| Docker | >= 20.0 | 容器化部署 |
| Docker Compose | >= 2.0 | 多容器编排 |
| Nginx | >= 1.20 | 反向代理 (生产环境) |

---

## 开发环境部署

### 1. 克隆代码

```bash
git clone <repository-url>
cd math-demo
```

### 2. 安装 pnpm

```bash
npm install -g pnpm
```

### 3. 安装依赖

#### 前端依赖
```bash
cd apps/web
npm install
```

#### API 服务依赖
```bash
cd apps/api
npm install
```

#### Parser 服务依赖
```bash
cd apps/parser
pip install -r requirements.txt
```

### 4. 配置环境变量

#### API 服务 (apps/api/.env)
```env
PORT=8080
NODE_ENV=development
LOG_LEVEL=info
PARSER_URL=http://localhost:8001
UPLOAD_DIR=uploads
EXPORT_DIR=exports
```

#### 前端 (apps/web/.env.development)
```env
VITE_API_URL=http://localhost:8080
VITE_PARSER_URL=http://localhost:8001
```

### 5. 启动服务

#### 方式一：分别启动

**终端 1 - 启动 Parser 服务**
```bash
cd apps/parser
python main.py
# 服务运行在 http://localhost:8001
```

**终端 2 - 启动 API 服务**
```bash
cd apps/api
npm run dev
# 服务运行在 http://localhost:8080
```

**终端 3 - 启动前端**
```bash
cd apps/web
npm run dev
# 服务运行在 http://localhost:3000
```

#### 方式二：使用 pnpm (推荐)

```bash
# 根目录
pnpm dev:parser  # 启动 Parser
pnpm dev:api     # 启动 API
pnpm dev:web     # 启动前端
```

### 6. 验证部署

```bash
# 运行集成测试
python test_monorepo.py
```

预期输出：
```
============================================================
数学试卷解析系统 - Monorepo 集成测试
============================================================

【测试1】前端服务 (Web)
  地址: http://localhost:3000
  ✓ 前端服务正常运行

【测试2】API 服务 (Node.js)
  地址: http://localhost:8080
  ✓ API 服务正常运行

【测试3】Parser 服务 (Python)
  地址: http://localhost:8001
  ✓ Parser 服务正常运行

【测试4】API 到 Parser 的连接
  ✓ 健康检查通过

============================================================
测试结果汇总
============================================================
  前端服务: ✓ 通过
  API 服务: ✓ 通过
  Parser 服务: ✓ 通过
  API-Parser 连接: ✓ 通过

✓ 所有测试通过！系统运行正常。
```

### 7. 访问应用

- **前端界面**: http://localhost:3000
- **API 文档**: http://localhost:8080/api/v1/health
- **Parser 文档**: http://localhost:8001/docs

---

## 生产环境部署

### 1. 构建应用

#### 构建前端
```bash
cd apps/web
npm run build
# 构建产物在 dist/ 目录
```

#### 构建 API 服务
```bash
cd apps/api
npm run build
# 构建产物在 dist/ 目录
```

### 2. 配置生产环境变量

#### API 服务 (apps/api/.env.production)
```env
PORT=8080
NODE_ENV=production
LOG_LEVEL=warn
PARSER_URL=http://parser:8000
UPLOAD_DIR=/app/uploads
EXPORT_DIR=/app/exports
```

#### 前端 (apps/web/.env.production)
```env
VITE_API_URL=/api
VITE_PARSER_URL=http://parser:8000
```

### 3. 使用 PM2 管理 Node.js 进程

```bash
# 安装 PM2
npm install -g pm2

# 创建 PM2 配置文件 (ecosystem.config.js)
module.exports = {
  apps: [
    {
      name: 'math-demo-api',
      script: './apps/api/dist/index.js',
      instances: 2,
      exec_mode: 'cluster',
      env: {
        NODE_ENV: 'production',
        PORT: 8080
      },
      log_file: './logs/api.log',
      error_file: './logs/api-error.log',
      out_file: './logs/api-out.log'
    }
  ]
};

# 启动服务
pm2 start ecosystem.config.js

# 保存配置
pm2 save

# 设置开机自启
pm2 startup
```

### 4. 使用 systemd 管理 Python 服务

创建 `/etc/systemd/system/math-demo-parser.service`:

```ini
[Unit]
Description=Math Demo Parser Service
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/var/www/math-demo/apps/parser
Environment="PATH=/var/www/math-demo/apps/parser/venv/bin"
Environment="PORT=8001"
ExecStart=/var/www/math-demo/apps/parser/venv/bin/python main.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

启动服务：
```bash
sudo systemctl daemon-reload
sudo systemctl enable math-demo-parser
sudo systemctl start math-demo-parser
sudo systemctl status math-demo-parser
```

### 5. Nginx 反向代理配置

```nginx
# /etc/nginx/sites-available/math-demo

upstream api_backend {
    server localhost:8080;
}

upstream parser_backend {
    server localhost:8001;
}

server {
    listen 80;
    server_name your-domain.com;

    # 前端静态文件
    location / {
        root /var/www/math-demo/apps/web/dist;
        index index.html;
        try_files $uri $uri/ /index.html;
    }

    # API 代理
    location /api/ {
        proxy_pass http://api_backend/;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_cache_bypass $http_upgrade;
        proxy_read_timeout 300s;
    }

    # 上传文件
    location /uploads/ {
        alias /var/www/math-demo/uploads/;
        expires 30d;
        add_header Cache-Control "public, immutable";
    }

    # 导出文件
    location /exports/ {
        alias /var/www/math-demo/exports/;
        expires 1d;
    }
}
```

启用配置：
```bash
sudo ln -s /etc/nginx/sites-available/math-demo /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

### 6. HTTPS 配置 (Let's Encrypt)

```bash
# 安装 Certbot
sudo apt install certbot python3-certbot-nginx

# 申请证书
sudo certbot --nginx -d your-domain.com

# 自动续期
sudo certbot renew --dry-run
```

---

## Docker 部署

### 1. 构建镜像

```bash
# 构建所有镜像
docker-compose build

# 或者分别构建
docker-compose build web
docker-compose build api
docker-compose build parser
```

### 2. 启动服务

```bash
# 后台启动
docker-compose up -d

# 查看日志
docker-compose logs -f

# 查看特定服务日志
docker-compose logs -f api
```

### 3. 停止服务

```bash
# 停止并删除容器
docker-compose down

# 停止并删除容器和数据卷
docker-compose down -v
```

### 4. 更新部署

```bash
# 拉取最新代码
git pull

# 重新构建并启动
docker-compose up -d --build

# 清理旧镜像
docker image prune -f
```

### 5. 生产环境 Docker 配置

创建 `docker-compose.prod.yml`:

```yaml
version: '3.8'

services:
  web:
    build:
      context: ./apps/web
      dockerfile: Dockerfile
    restart: always
    ports:
      - "3000:3000"
    environment:
      - NODE_ENV=production
    networks:
      - math-demo-network

  api:
    build:
      context: ./apps/api
      dockerfile: Dockerfile
    restart: always
    ports:
      - "8080:8080"
    environment:
      - NODE_ENV=production
      - PORT=8080
      - PARSER_URL=http://parser:8000
    volumes:
      - ./uploads:/app/uploads
      - ./exports:/app/exports
      - ./logs:/app/logs
    depends_on:
      - parser
    networks:
      - math-demo-network

  parser:
    build:
      context: ./apps/parser
      dockerfile: Dockerfile
    restart: always
    ports:
      - "8000:8000"
    environment:
      - PYTHONUNBUFFERED=1
    volumes:
      - ./uploads:/app/uploads
      - ./exports:/app/exports
    networks:
      - math-demo-network

networks:
  math-demo-network:
    driver: bridge

volumes:
  uploads:
  exports:
  logs:
```

启动生产环境：
```bash
docker-compose -f docker-compose.prod.yml up -d
```

---

## 服务配置

### 端口配置

| 服务 | 开发端口 | 生产端口 | 说明 |
|-----|---------|---------|------|
| Web | 3000 | 3000 | 前端服务 |
| API | 8080 | 8080 | Node.js API |
| Parser | 8001 | 8000 | Python Parser |

### 目录结构

```
/var/www/math-demo/
├── apps/
│   ├── web/
│   ├── api/
│   └── parser/
├── uploads/          # 上传文件
├── exports/          # 导出文件
├── logs/             # 日志文件
└── docker-compose.yml
```

### 环境变量说明

#### API 服务

| 变量名 | 默认值 | 说明 |
|-------|-------|------|
| PORT | 8080 | 服务端口 |
| NODE_ENV | development | 运行环境 |
| LOG_LEVEL | info | 日志级别 |
| PARSER_URL | http://localhost:8000 | Parser 服务地址 |
| UPLOAD_DIR | uploads | 上传目录 |
| EXPORT_DIR | exports | 导出目录 |

#### Parser 服务

| 变量名 | 默认值 | 说明 |
|-------|-------|------|
| PORT | 8000 | 服务端口 |
| PYTHONUNBUFFERED | 1 | Python 无缓冲输出 |

---

## 常见问题

### 1. 端口被占用

**问题**: `Error: listen EADDRINUSE: address already in use :::8080`

**解决**:
```bash
# 查找占用端口的进程
lsof -i :8080

# 或者使用 netstat
netstat -tlnp | grep 8080

# 杀死进程
kill -9 <PID>

# 或者修改环境变量使用其他端口
PORT=8081 npm run dev
```

### 2. Python 依赖安装失败

**问题**: `ERROR: Could not install packages due to an OSError`

**解决**:
```bash
# 使用 --user 选项
pip install --user -r requirements.txt

# 或者使用虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

### 3. Node.js 内存不足

**问题**: `JavaScript heap out of memory`

**解决**:
```bash
# 增加内存限制
export NODE_OPTIONS="--max-old-space-size=4096"
npm run build
```

### 4. CORS 错误

**问题**: `Access-Control-Allow-Origin` 错误

**解决**:
- 检查 API 服务的 CORS 配置
- 确保前端请求的 API 地址正确
- 生产环境使用相同的域名

### 5. 文件上传失败

**问题**: 上传大文件时失败

**解决**:
```nginx
# nginx.conf
client_max_body_size 100M;
```

```javascript
// API 服务
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));
```

---

## 监控与日志

### 日志位置

| 服务 | 日志位置 | 说明 |
|-----|---------|------|
| API | `apps/api/logs/` | Winston 日志 |
| Parser | 控制台输出 | 可重定向到文件 |
| Nginx | `/var/log/nginx/` | 访问和错误日志 |

### 查看日志

```bash
# API 服务日志
tail -f apps/api/logs/combined.log

# Parser 服务日志 (Docker)
docker-compose logs -f parser

# Nginx 日志
sudo tail -f /var/log/nginx/access.log
sudo tail -f /var/log/nginx/error.log
```

### 健康检查

```bash
# API 健康检查
curl http://localhost:8080/api/v1/health

# Parser 健康检查
curl http://localhost:8001/health

# 运行完整测试
python test_monorepo.py
```

### 性能监控

```bash
# 查看 Node.js 进程
pm2 monit

# 查看系统资源
htop

# Docker 容器状态
docker stats
```

---

## 备份与恢复

### 备份数据

```bash
#!/bin/bash
# backup.sh

BACKUP_DIR="/backup/math-demo/$(date +%Y%m%d)"
mkdir -p $BACKUP_DIR

# 备份上传文件
cp -r /var/www/math-demo/uploads $BACKUP_DIR/

# 备份导出文件
cp -r /var/www/math-demo/exports $BACKUP_DIR/

# 备份代码
cd /var/www/math-demo
git bundle create $BACKUP_DIR/repo.bundle --all

echo "Backup completed: $BACKUP_DIR"
```

### 恢复数据

```bash
#!/bin/bash
# restore.sh

BACKUP_DIR="/backup/math-demo/20240304"

# 恢复上传文件
cp -r $BACKUP_DIR/uploads /var/www/math-demo/

# 恢复导出文件
cp -r $BACKUP_DIR/exports /var/www/math-demo/

echo "Restore completed"
```

---

## 更新与维护

### 更新代码

```bash
# 拉取最新代码
git pull origin main

# 安装新依赖
pnpm install

# 重新构建
pnpm build

# 重启服务
pm2 restart all
# 或
docker-compose up -d --build
```

### 数据库迁移 (未来扩展)

```bash
# 如果有数据库，运行迁移
pnpm migrate
```

---

## 安全建议

1. **使用 HTTPS**: 生产环境必须启用 HTTPS
2. **防火墙配置**: 只开放必要的端口
3. **定期更新**: 及时更新依赖包和系统补丁
4. **文件权限**: 设置正确的文件权限
5. **日志审计**: 定期检查日志文件
6. **备份策略**: 定期备份重要数据

---

## 联系支持

如有问题，请：
1. 查看日志文件
2. 运行测试脚本
3. 检查环境变量配置
4. 参考本文档的常见问题部分

---

**最后更新**: 2026-03-04
