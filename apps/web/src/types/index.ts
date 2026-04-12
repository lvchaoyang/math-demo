// 题目类型
export type QuestionType = 
  | 'unknown'
  | 'single_choice'
  | 'multiple_choice'
  | 'fill_blank'
  | 'true_false'
  | 'short_answer'
  | 'calculation'
  | 'proof'
  | 'comprehensive'

// 选项
export interface Option {
  label: string
  content: string
  is_latex: boolean
  content_html?: string
}

// 题目
export interface Question {
  id: string
  number: number
  type: QuestionType
  type_name: string
  content: string
  content_html: string
  /** 本题图片目录（与 data/images/{file_id} 一致），跨卷组卷导出依赖 */
  file_id?: string
  options: Option[]
  answer?: string
  analysis?: string
  score?: number
  difficulty?: string
  images: string[]
  latex_formulas: string[]
  /** 与 Parser 对齐；导出时可随 assembly 传给 API，题干按段写入 Word */
  content_export_segments?: Array<{
    kind: string
    text?: string
    latex?: string
    filename?: string
    display?: string
    source_image?: string
  }>
  confidence_score?: number
  low_confidence?: boolean
  low_confidence_reasons?: string[]
}

// 上传响应
export interface UploadResponse {
  success: boolean
  file_id: string
  filename: string
  question_count: number
  questions: Question[]
  message: string
}

export interface ExportAssemblyItem {
  file_id: string
  question_id: string
  /** 完整题目，导出以页面数据为准 */
  question?: Question
  /** 与 question.content_export_segments 一致时可显式带上；API 会用于覆盖，Word 按段写题干与选项 */
  content_export_segments?: Question['content_export_segments']
}

// 导出请求
export interface ExportRequest {
  file_id?: string
  question_ids?: string[]
  assembly?: ExportAssemblyItem[]
  title?: string
  watermark?: string
  include_answer: boolean
  include_analysis: boolean
  paper_size: string
}

// 导出响应
export interface ExportResponse {
  success: boolean
  download_url: string
  filename: string
  message: string
}

// 解析进度
export interface ParseProgress {
  file_id: string
  status: 'pending' | 'parsing' | 'completed' | 'error'
  progress: number
  message: string
  questions: Question[]
}
