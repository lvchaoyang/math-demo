import { Router } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import axios from 'axios';
import path from 'path';
import fs from 'fs';
import { logger } from '../utils/logger.js';
import type { UploadResponse, ParseProgress, Question } from '../types/index.js';

const router = Router();

// 存储解析进度

// 获取 Parser URL（在运行时读取环境变量）
function getParserUrl(): string {
  return process.env.PARSER_URL || 'http://localhost:8001';
}
const parseProgressStore = new Map<string, ParseProgress>();

// 确保 uploads 目录存在（使用统一的数据目录）
const UPLOAD_DIR = path.join(process.cwd(), '..', '..', 'data', 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  logger.info(`Created upload directory: ${UPLOAD_DIR}`);
}

// 配置 multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const fileId = uuidv4();
    cb(null, `${fileId}_${file.originalname}`);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        file.originalname.endsWith('.docx')) {
      cb(null, true);
    } else {
      cb(new Error('只支持 .docx 格式的文件'));
    }
  }
});

// 上传并解析
router.post('/', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: '没有上传文件'
      });
    }

    const fileId = path.basename(req.file.filename, path.extname(req.file.filename)).split('_')[0];
    const filePath = req.file.path;
    const mode = req.body.mode || 'questions'; // 默认为题目拆分模式

    // 初始化进度
    parseProgressStore.set(fileId, {
      file_id: fileId,
      status: 'parsing',
      progress: 0,
      message: '开始解析...',
      mode: mode
    });

    // 异步调用 Python Parser 服务
    parseDocument(fileId, filePath, req.file.originalname, mode);

    const response: UploadResponse = {
      success: true,
      file_id: fileId,
      filename: req.file.originalname,
      message: '文件上传成功，正在解析...',
      mode: mode
    };

    res.json(response);
  } catch (error) {
    logger.error(`Upload error: ${error}`);
    res.status(500).json({
      success: false,
      message: '上传失败'
    });
  }
});

// 异步解析文档
async function parseDocument(fileId: string, filePath: string, originalName: string, mode: string = 'questions') {
  try {
    parseProgressStore.get(fileId)!.progress = 20;
    parseProgressStore.get(fileId)!.message = '调用解析服务...';

    // 调用 Python Parser 服务
    const FormData = (await import('form-data')).default;
    const formData = new FormData();
    formData.append('file', fs.createReadStream(filePath), originalName);
    formData.append('mode', mode); // 添加模式参数

    const parserUrl = getParserUrl();
    logger.info(`Sending request to Parser: ${parserUrl}/parse with mode: ${mode}`);
    
    const response = await axios.post(`${parserUrl}/parse`, formData, {
      headers: formData.getHeaders(),
      timeout: 300000 // 5分钟超时
    });

    logger.info(`Parser response status: ${response.status}`);
    logger.info(`Parser response data: ${JSON.stringify(response.data).substring(0, 200)}...`);

    if (response.data.success) {
      if (mode === 'html') {
        // HTML 模式：返回完整的 HTML
        parseProgressStore.set(fileId, {
          file_id: fileId,
          status: 'completed',
          progress: 100,
          message: 'HTML 转换完成',
          mode: 'html',
          html: response.data.html
        });
        logger.info(`HTML conversion completed for file: ${fileId}`);
      } else {
        // 题目拆分模式
        console.log('response.data=>',response.data);
        parseProgressStore.set(fileId, {
          file_id: fileId,
          status: 'completed',
          progress: 100,
          message: '解析完成',
          mode: 'questions',
          questions: response.data.questions
        });
        logger.info(`Parse completed for file: ${fileId}`);
      }
    } else {
      throw new Error(response.data.message || '解析失败');
    }
  } catch (error: any) {
    logger.error(`Parse error for file ${fileId}: ${error.message}`);
    if (error.response) {
      logger.error(`Response status: ${error.response.status}`);
      logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
    }
    parseProgressStore.set(fileId, {
      file_id: fileId,
      status: 'error',
      progress: 0,
      message: `解析失败: ${error.message}`
    });
  }
}

// 获取解析进度
router.get('/progress/:fileId', (req, res) => {
  const { fileId } = req.params;
  const progress = parseProgressStore.get(fileId);
  
  if (!progress) {
    return res.status(404).json({
      success: false,
      message: '文件ID不存在'
    });
  }
  
  res.json(progress);
});

// 获取题目列表
router.get('/questions/:fileId', (req, res) => {
  const { fileId } = req.params;
  const progress = parseProgressStore.get(fileId);
  
  if (!progress) {
    return res.status(404).json({
      success: false,
      message: '文件ID不存在'
    });
  }
  
  if (progress.status !== 'completed') {
    return res.status(400).json({
      success: false,
      message: '文件尚未解析完成'
    });
  }
  
  res.json(progress.questions || []);
});

export { router as uploadRouter };
