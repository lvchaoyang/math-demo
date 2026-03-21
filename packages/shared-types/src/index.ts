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

// 题目
export interface Question {
  id: string;
  number: number;
  type: QuestionType;
  type_name: string;
  content: string;
  content_html: string;
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

// 导出请求
export interface ExportRequest {
  file_id: string;
  question_ids: string[];
  options: {
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
