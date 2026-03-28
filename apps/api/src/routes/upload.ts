/**
 * 文件上传与解析路由
 * 
 * 解析方案说明：
 * - 默认使用 Pandoc 方案（/parse/v2），提供更高质量的数学公式解析
 * - Pandoc 方案支持复杂的数学公式（积分、矩阵、多行方程等）
 * - 当 Pandoc 不可用时会自动降级到原有方案
 * 
 * 解析模式：
 * - questions: 题目拆分模式（默认），将试卷拆分为独立题目
 * - html: HTML 模式，返回完整的 HTML 文档
 */

import { Router } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import axios from 'axios';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { logger } from '../utils/logger.js';
import type { UploadResponse, ParseProgress, Question } from '../types/index.js';

const router = Router();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.resolve(__dirname, '../../../../');

const parseProgressStore = new Map<string, ParseProgress>();

function getParserUrl(): string {
  return process.env.PARSER_URL || 'http://localhost:8000';
}

const UPLOAD_DIR = path.join(PROJECT_ROOT, 'data', 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  logger.info(`Created upload directory: ${UPLOAD_DIR}`);
}

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

/**
 * 上传并解析文档
 * 
 * @route POST /api/upload
 * @param {File} file - DOCX 格式的试卷文件
 * @param {string} mode - 解析模式（questions|html），默认 questions
 * @param {boolean} usePandoc - 是否使用 Pandoc 方案，默认 true
 * @returns {UploadResponse} 上传结果和文件 ID
 */
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
    const mode = req.body.mode || 'questions';
    const usePandoc = req.body.usePandoc !== 'false';

    parseProgressStore.set(fileId, {
      file_id: fileId,
      status: 'parsing',
      progress: 0,
      message: '开始解析...',
      mode: mode
    });

    parseDocument(fileId, filePath, req.file.originalname, mode, usePandoc);

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

/**
 * 异步解析文档
 * 
 * 解析流程：
 * 1. 优先调用 Pandoc 方案（/parse/v2）
 * 2. Pandoc 失败时自动降级到原有方案
 * 3. 支持题目拆分和 HTML 两种模式
 * 
 * @param fileId - 文件唯一标识
 * @param filePath - 文件路径
 * @param originalName - 原始文件名
 * @param mode - 解析模式（questions|html）
 * @param usePandoc - 是否使用 Pandoc 方案，默认 true
 */
async function parseDocument(
  fileId: string, 
  filePath: string, 
  originalName: string, 
  mode: string = 'questions',
  usePandoc: boolean = true
) {
  try {
    parseProgressStore.get(fileId)!.progress = 20;
    parseProgressStore.get(fileId)!.message = '调用解析服务...';

    const FormData = (await import('form-data')).default;
    const formData = new FormData();
    formData.append('file', fs.createReadStream(filePath), originalName);
    formData.append('mode', mode);
    formData.append('usePandoc', usePandoc.toString());

    const parserUrl = getParserUrl();
    
    const endpoint = usePandoc ? '/parse/v2' : '/parse';
    logger.info(`Sending request to Parser: ${parserUrl}${endpoint} with mode: ${mode}, usePandoc: ${usePandoc}`);
    
    const response = await axios.post(`${parserUrl}${endpoint}`, formData, {
      headers: formData.getHeaders(),
      timeout: 300000
    });

    logger.info(`Parser response status: ${response.status}`);
    logger.info(`Parser response data: ${JSON.stringify(response.data).substring(0, 200)}...`);

    if (response.data.success) {
      const parseMethod = response.data.method || 'unknown';
      
      if (mode === 'html') {
        parseProgressStore.set(fileId, {
          file_id: fileId,
          status: 'completed',
          progress: 100,
          message: `HTML 转换完成（使用 ${parseMethod} 方案）`,
          mode: 'html',
          html: response.data.html
        });
        logger.info(`HTML conversion completed for file: ${fileId}, method: ${parseMethod}`);
      } else {
        const formulaRenderSummary = response.data.formula_render_summary || undefined;
        const formulaAssetDebug = response.data.formula_asset_debug || undefined;
        const formulaRenderPlan = response.data.formula_render_plan || undefined;
        parseProgressStore.set(fileId, {
          file_id: fileId,
          status: 'completed',
          progress: 100,
          message: `解析完成（使用 ${parseMethod} 方案）`,
          mode: 'questions',
          questions: response.data.questions,
          formula_render_summary: formulaRenderSummary,
          formula_asset_debug: formulaAssetDebug,
          formula_render_plan: formulaRenderPlan,
        });
        logger.info(`Parse completed for file: ${fileId}, method: ${parseMethod}`);
        logger.info(`Formulas extracted: ${response.data.formulas_count || 0}, Images: ${response.data.images_count || 0}`);
        if (formulaRenderSummary) {
          logger.info(
            `Formula render summary: total=${formulaRenderSummary.total}, rendered=${formulaRenderSummary.rendered}, source_only=${formulaRenderSummary.source_only}`
          );
        }
        if (formulaAssetDebug) {
          logger.info(
            `Formula asset debug: paragraphs=${formulaAssetDebug.paragraphs}, content_items=${formulaAssetDebug.total_content_items}, images=${formulaAssetDebug.image_count}`
          );
        }
        if (formulaRenderPlan) {
          logger.info(`Formula render plan entries: ${Array.isArray(formulaRenderPlan) ? formulaRenderPlan.length : 0}`);
        }
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

/**
 * 获取解析进度
 * 
 * @route GET /api/upload/progress/:fileId
 * @param fileId - 文件唯一标识
 * @returns {ParseProgress} 解析进度信息
 */
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

/**
 * 获取题目列表
 * 
 * @route GET /api/upload/questions/:fileId
 * @param fileId - 文件唯一标识
 * @returns {Question[]} 题目列表
 */
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
