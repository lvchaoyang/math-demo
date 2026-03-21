FROM node:18-alpine

WORKDIR /app

# 安装 pnpm
RUN npm install -g pnpm

# 复制 package.json
COPY package.json ./
COPY pnpm-lock.yaml* ./

# 安装依赖
RUN pnpm install

# 复制源代码
COPY . .

# 构建
RUN pnpm build

# 创建日志目录
RUN mkdir -p logs

# 创建上传和导出目录
RUN mkdir -p uploads exports

EXPOSE 8080

CMD ["node", "dist/index.js"]
