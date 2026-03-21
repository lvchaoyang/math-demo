/**
 * Math Demo API Server
 * Node.js API 服务 - 业务逻辑层
 */

import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { uploadRouter } from './routes/upload.js';
import { exportRouter } from './routes/export.js';
import { healthRouter } from './routes/health.js';
import { imagesRouter } from './routes/images.js';
import { errorHandler } from './middleware/errorHandler.js';
import { logger } from './utils/logger.js';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 8080;
const PARSER_URL = process.env.PARSER_URL || 'http://localhost:8000';

// 中间件
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 日志中间件
app.use((req, res, next) => {
  // logger.info(`${req.method} ${req.path}`);
  next();
});

// 路由
app.use('/api/v1/upload', uploadRouter);
app.use('/api/v1/export', exportRouter);
app.use('/api/v1/health', healthRouter);
app.use('/api/v1/images', imagesRouter);

// 静态文件服务
app.use('/uploads', express.static('uploads'));
app.use('/exports', express.static('exports'));

// 根路径
app.get('/', (req, res) => {
  res.json({
    message: '数学试卷解析系统 API (Node.js)',
    version: '1.0.0',
    docs: '/api/v1/health',
    parser_service: PARSER_URL
  });
});

// 错误处理
app.use(errorHandler);

// 启动服务器
app.listen(PORT, () => {
  logger.info(`API Server running on http://localhost:${PORT}`);
  logger.info(`Parser Service: ${PARSER_URL}`);
});

export default app;
