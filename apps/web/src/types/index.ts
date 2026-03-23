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
  options: Option[]
  answer?: string
  analysis?: string
  score?: number
  difficulty?: string
  images: string[]
  latex_formulas: string[]
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

// 导出请求
export interface ExportRequest {
  question_ids: string[]
  title: string
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
