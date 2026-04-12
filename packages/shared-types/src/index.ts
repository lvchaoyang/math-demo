/**
 * 数学试卷解析系统 - 共享类型定义
 * 用于前后端类型统一
 */

// 题目类型枚举
export enum QuestionType {
  SINGLE_CHOICE = 'single_choice',
  MULTIPLE_CHOICE = 'multiple_choice',
  FILL_BLANK = 'fill_blank',
  ANSWER = 'answer',
  PROOF = 'proof',
  UNKNOWN = 'unknown'
}

// 选项
export interface Option {
  label: string;
  content: string;
  is_latex: boolean;
}

/** Parser 题干导出片段 */
export interface ContentExportSegment {
  kind: string;
  text?: string;
  latex?: string;
  filename?: string;
  display?: string;
  source_image?: string;
}

// 题目
export interface Question {
  id: string;
  number: number;
  type: QuestionType;
  type_name: string;
  content: string;
  content_html: string;
  file_id?: string;
  content_export_segments?: ContentExportSegment[];
  options: Option[];
  answer?: string;
  analysis?: string;
  score?: number;
  difficulty?: string;
  images: string[];
  latex_formulas: string[];
}

// 解析进度
export interface ParseProgress {
  file_id: string;
  status: 'parsing' | 'completed' | 'error';
  progress: number;
  message: string;
  questions?: Question[];
}

// 上传响应
export interface UploadResponse {
  success: boolean;
  file_id: string;
  filename: string;
  message: string;
}

/** 跨卷组卷：按顺序导出 */
export interface ExportAssemblyItem {
  file_id: string;
  question_id: string;
  /**
   * 若带上且 id 与 question_id 一致：导出以该完整题目为准（与页面展示一致），含 options/content_html/segments 等。
   * 仍校验快照中存在该题；file_id、导出题号由 assembly 项强制覆盖。
   */
  question?: Question;
  /** 有则覆盖本题 content_export_segments：按段导出题干与选项，LaTeX 在 Parser 侧转 MathType */
  content_export_segments?: ContentExportSegment[];
}

// 导出请求
export interface ExportRequest {
  file_id?: string;
  question_ids?: string[];
  assembly?: ExportAssemblyItem[];
  title?: string;
  options?: {
    include_answer: boolean;
    include_analysis: boolean;
    watermark?: string;
  };
}

// 导出响应
export interface ExportResponse {
  success: boolean;
  download_url: string;
  filename: string;
}

// 段落内容项
export interface ContentItem {
  type: 'text' | 'latex' | 'latex_block' | 'image' | 'formula_image' | 'break';
  content: string | { filename: string; width?: number; height?: number };
  latex?: string;
}

// 段落
export interface Paragraph {
  text: string;
  style: string | null;
  is_bold: boolean;
  is_formula: boolean;
  latex_content: string | null;
  content_items: ContentItem[];
}

// 解析结果
export interface ParseResult {
  paragraphs: Paragraph[];
  images: string[];
  formulas: string[];
}
