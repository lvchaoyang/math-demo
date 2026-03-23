export interface Question {
  id: string;
  number: number;
  type: string;
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

export interface Option {
  label: string;
  content: string;
  is_latex: boolean;
}

export interface ParseProgress {
  file_id: string;
  status: 'parsing' | 'completed' | 'error';
  progress: number;
  message: string;
  mode?: 'questions' | 'html';
  questions?: Question[];
  html?: string;
}

export interface UploadResponse {
  success: boolean;
  file_id: string;
  filename: string;
  message: string;
  mode?: 'questions' | 'html';
}

export interface ExportRequest {
  file_id: string;
  question_ids: string[];
  title?: string;
  options: {
    include_answer: boolean;
    include_analysis: boolean;
    watermark?: string;
  };
}
