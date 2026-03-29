import { Router } from 'express';
import axios from 'axios';
import { logger } from '../utils/logger.js';
import type { ExportRequest, Question } from '../types/index.js';
import { getParseProgressSnapshot } from './upload.js';

const router = Router();

// 获取 Parser URL（在运行时读取环境变量）
function getParserUrl(): string {
  return process.env.PARSER_URL || 'http://localhost:8000';
}

/** 导出会重新执行完整 parse_docx + 拆题，大试卷可能远超默认 60s */
const EXPORT_TIMEOUT_MS = Number(process.env.EXPORT_TIMEOUT_MS || 600000);

function fastApiDetailMessage(data: unknown): string | null {
  if (!data || typeof data !== 'object') return null;
  const d = (data as { detail?: unknown }).detail;
  if (typeof d === 'string') return d;
  if (Array.isArray(d)) return d.map((x) => (typeof x === 'object' ? JSON.stringify(x) : String(x))).join('；');
  return null;
}

// 导出题目
router.post('/', async (req, res) => {
  try {
    const body = req.body as ExportRequest;
    const { file_id, question_ids, assembly, title } = body;
    const options =
      body.options ?? { include_answer: false, include_analysis: false };

    const parserUrl = getParserUrl();
    let exportFileId: string;
    let exportQuestionIds: string[];
    let exportBody: Record<string, unknown>;

    if (Array.isArray(assembly) && assembly.length > 0) {
      const ordered: Question[] = [];
      for (let i = 0; i < assembly.length; i++) {
        const item = assembly[i];
        const fid = item?.file_id;
        const qid = item?.question_id;
        if (!fid || !qid) {
          return res.status(400).json({
            success: false,
            message: '参数错误: assembly 中每项需要 file_id 与 question_id'
          });
        }
        const snap = getParseProgressSnapshot(fid);
        if (
          snap?.status !== 'completed' ||
          snap.mode !== 'questions' ||
          !Array.isArray(snap.questions)
        ) {
          return res.status(400).json({
            success: false,
            message: `试卷尚未解析完成或已失效，请重新上传：${fid}`
          });
        }
        const found = snap.questions.find((q) => q.id === qid);
        if (!found) {
          return res.status(400).json({
            success: false,
            message: `未找到题目 ${qid}，请重新解析后再导出`
          });
        }
        ordered.push({ ...found, number: i + 1 });
      }
      exportFileId = assembly[0]!.file_id;
      exportQuestionIds = ordered.map((q) => q.id);
      exportBody = {
        file_id: exportFileId,
        question_ids: exportQuestionIds,
        options,
        title,
        questions: ordered,
        use_questions_payload_order: true
      };
      logger.info(
        `Export start (assembly): files=${new Set(assembly.map((a) => a.file_id)).size}, questions=${ordered.length}, timeoutMs=${EXPORT_TIMEOUT_MS}`
      );
    } else if (file_id && question_ids && Array.isArray(question_ids) && question_ids.length > 0) {
      exportFileId = file_id;
      exportQuestionIds = question_ids;
      const snap = getParseProgressSnapshot(file_id);
      const cachedQuestions =
        snap?.status === 'completed' &&
        snap.mode === 'questions' &&
        Array.isArray(snap.questions) &&
        snap.questions.length > 0
          ? snap.questions
          : undefined;

      logger.info(
        `Export start: file_id=${file_id}, questions=${question_ids.length}, timeoutMs=${EXPORT_TIMEOUT_MS}, useCached=${Boolean(cachedQuestions)}`
      );
      exportBody = {
        file_id,
        question_ids,
        options,
        title,
        ...(cachedQuestions ? { questions: cachedQuestions } : {})
      };
    } else {
      return res.status(400).json({
        success: false,
        message: '参数错误: 需要 assembly，或 file_id 与 question_ids'
      });
    }

    const response = await axios.post(`${parserUrl}/export`, exportBody, {
      responseType: 'stream',
      timeout: EXPORT_TIMEOUT_MS,
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });

    // 设置响应头
    res.setHeader('Content-Type', response.headers['content-type'] || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', response.headers['content-disposition'] || 'attachment; filename=exported.docx');
    
    // 流式传输
    response.data.pipe(res);
    
    logger.info(
      `Export completed for file: ${exportFileId}, questions: ${exportQuestionIds.length}`
    );
  } catch (error: unknown) {
    logger.error(`Export error: ${error}`);
    let message = '导出失败';
    if (axios.isAxiosError(error)) {
      if (error.code === 'ECONNABORTED' || String(error.message || '').toLowerCase().includes('timeout')) {
        message =
          '导出超时：服务端需重新解析整卷文档，大文件耗时较长。可稍后重试、少选几题导出，或在环境变量 EXPORT_TIMEOUT_MS 中加大时限（默认 600000）。';
      } else if (error.response?.data) {
        const parsed = fastApiDetailMessage(error.response.data);
        if (parsed) message = parsed;
      }
    }
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        message
      });
    }
  }
});

export { router as exportRouter };
