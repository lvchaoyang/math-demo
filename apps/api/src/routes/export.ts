import { Router } from 'express';
import axios from 'axios';
import { logger } from '../utils/logger.js';
import type { ExportRequest } from '../types/index.js';

const router = Router();

// 获取 Parser URL（在运行时读取环境变量）
function getParserUrl(): string {
  return process.env.PARSER_URL || 'http://localhost:8000';
}

// 导出题目
router.post('/', async (req, res) => {
  try {
    const { file_id, question_ids, options, title }: ExportRequest = req.body;
    
    if (!file_id || !question_ids || !Array.isArray(question_ids)) {
      return res.status(400).json({
        success: false,
        message: '参数错误: 需要 file_id 和 question_ids'
      });
    }

    // 转发到 Python Export 服务
    const parserUrl = getParserUrl();
    const response = await axios.post(`${parserUrl}/export`, {
      file_id,
      question_ids,
      options,
      title
    }, {
      responseType: 'stream',
      timeout: 60000
    });

    // 设置响应头
    res.setHeader('Content-Type', response.headers['content-type'] || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', response.headers['content-disposition'] || 'attachment; filename=exported.docx');
    
    // 流式传输
    response.data.pipe(res);
    
    logger.info(`Export completed for file: ${file_id}, questions: ${question_ids.length}`);
  } catch (error) {
    logger.error(`Export error: ${error}`);
    res.status(500).json({
      success: false,
      message: '导出失败'
    });
  }
});

export { router as exportRouter };
