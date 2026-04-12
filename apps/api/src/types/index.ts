/** Parser 题干导出片段，与 Python Question.content_export_segments 对齐 */
export interface ContentExportSegment {
  kind: string;
  text?: string;
  latex?: string;
  filename?: string;
  display?: string;
  source_image?: string;
}

export interface Question {
  id: string;
  number: number;
  type: string;
  type_name: string;
  content: string;
  content_html: string;
  /** 本题资源所在上传目录，跨卷导出时 Parser 按题切换 data/images/{file_id}/ */
  file_id?: string;
  /** 有则 Word 导出优先按片段顺序写入，避免再解析 content_html */
  content_export_segments?: ContentExportSegment[];
  options: Option[];
  answer?: string;
  analysis?: string;
  score?: number;
  difficulty?: string;
  images: string[];
  latex_formulas: string[];
}

export interface Option {
  label: string;
  content: string;
  content_html?: string;
  is_latex: boolean;
  images?: string[];
}

export interface ParseProgress {
  file_id: string;
  status: 'parsing' | 'completed' | 'error';
  progress: number;
  message: string;
  mode?: 'questions' | 'html';
  questions?: Question[];
  html?: string;
  formula_render_summary?: {
    total: number;
    rendered: number;
    source_only: number;
    planned: number;
    skip: number;
    by_source_type?: Record<string, number>;
    by_action?: Record<string, number>;
    by_note?: Record<string, number>;
  };
  formula_asset_debug?: {
    paragraphs: number;
    paragraphs_with_content_items: number;
    total_content_items: number;
    item_type_count?: Record<string, number>;
    image_count: number;
    image_ext_count?: Record<string, number>;
    image_meta_type_count?: Record<string, number>;
  };
  formula_render_plan?: Array<{
    asset_id?: string;
    source_type?: string;
    action?: string;
    cache_key?: string;
    source_filename?: string | null;
    source_path?: string | null;
    target_path?: string;
    status?: string;
    note?: string;
    rendered_image?: string | null;
  }>;
}

export interface UploadResponse {
  success: boolean;
  file_id: string;
  filename: string;
  message: string;
  mode?: 'questions' | 'html';
}

/** 跨卷组卷：按数组顺序导出 */
export interface ExportAssemblyItem {
  file_id: string;
  question_id: string;
  /** 完整题目（与前端当前一致）；优先于内存快照 */
  question?: Question;
  /** 有则覆盖本题 content_export_segments：Word 按段导出（含选择题选项段），LaTeX 走 MathType/OMML */
  content_export_segments?: ContentExportSegment[];
}

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
