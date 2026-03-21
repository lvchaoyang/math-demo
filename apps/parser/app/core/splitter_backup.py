"""
题目拆分引擎
将解析后的文档内容按题目拆分为独立单元
"""

import re
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum


class QuestionType(Enum):
    """题目类型"""
    UNKNOWN = "unknown"
    SINGLE_CHOICE = "single_choice"      # 单选题
    MULTIPLE_CHOICE = "multiple_choice"  # 多选题
    FILL_BLANK = "fill_blank"            # 填空题
    TRUE_FALSE = "true_false"            # 判断题
    SHORT_ANSWER = "short_answer"        # 简答题
    CALCULATION = "calculation"          # 计算题
    PROOF = "proof"                      # 证明题
    COMPREHENSIVE = "comprehensive"      # 综合题


@dataclass
class Option:
    """选择题选项"""
    label: str           # A, B, C, D
    content: str         # 选项内容
    content_html: str = ""  # 选项内容（HTML格式）
    is_latex: bool = False  # 是否包含公式


@dataclass
class Question:
    """题目数据结构"""
    id: str                      # 唯一标识
    number: int                  # 题号
    type: QuestionType           # 题目类型
    type_name: str               # 题型名称（如"选择题"）
    content: str                 # 题目内容（纯文本）
    content_html: str            # 题目内容（HTML格式，含公式）
    options: List[Option] = field(default_factory=list)  # 选项（选择题）
    answer: Optional[str] = None # 答案（如果有）
    analysis: Optional[str] = None  # 解析（如果有）
    score: Optional[float] = None   # 分值
    difficulty: Optional[str] = None  # 难度
    images: List[str] = field(default_factory=list)  # 关联的图片
    latex_formulas: List[str] = field(default_factory=list)  # 包含的LaTeX公式
    raw_paragraphs: List[Dict] = field(default_factory=list)  # 原始段落数据


class QuestionSplitter:
    """题目拆分器"""

    TYPE_PATTERNS = [
        (r'^[一二三四五六七八九十]+[、\.]\s*选择题', QuestionType.SINGLE_CHOICE, '选择题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*单选题', QuestionType.SINGLE_CHOICE, '单选题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*多选题', QuestionType.MULTIPLE_CHOICE, '多选题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*填空题', QuestionType.FILL_BLANK, '填空题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*判断题', QuestionType.TRUE_FALSE, '判断题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*简答题', QuestionType.SHORT_ANSWER, '简答题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*计算题', QuestionType.CALCULATION, '计算题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*证明题', QuestionType.PROOF, '证明题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*解答题', QuestionType.CALCULATION, '解答题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*综合题', QuestionType.COMPREHENSIVE, '综合题'),
        (r'^[一二三四五六七八九十]+[、\.]\s*应用题', QuestionType.CALCULATION, '应用题'),
        (r'^Part\s*[A-D]\s*[:：]?\s*选择题', QuestionType.SINGLE_CHOICE, '选择题'),
        (r'^Part\s*[A-D]\s*[:：]?\s*单选题', QuestionType.SINGLE_CHOICE, '单选题'),
        (r'^Part\s*[A-D]\s*[:：]?\s*多选题', QuestionType.MULTIPLE_CHOICE, '多选题'),
        (r'^Part\s*[A-D]\s*[:：]?\s*填空题', QuestionType.FILL_BLANK, '填空题'),
        (r'^选择题\s*[\(（][^）\)]*[\)）]', QuestionType.SINGLE_CHOICE, '选择题'),
        (r'^单选题\s*[\(（][^）\)]*[\)）]', QuestionType.SINGLE_CHOICE, '单选题'),
        (r'^多选题\s*[\(（][^）\)]*[\)）]', QuestionType.MULTIPLE_CHOICE, '多选题'),
        (r'^填空题\s*[\(（][^）\)]*[\)）]', QuestionType.FILL_BLANK, '填空题'),
    ]
    
    QUESTION_NUMBER_PATTERNS = [
        r'^\s*(\d+)\s*[\.、．,，:：]\s*',
        r'^\s*[\(（](\d+)[\)）]\s*',
        r'^\s*(\d+)\s*\)\s*',
        r'^\s*第\s*(\d+)\s*题',
        r'^\s*Question\s*(\d+)',
        r'^\s*Q(\d+)',
    ]
    
    OPTION_PATTERNS = [
        r'([A-H])\s*[\.、．,，:：]\s*',
        r'([A-H])\s*[\.、．]\s*',
        r'[\(（]([A-H])[\)）]\s*',
    ]
    
    ANSWER_PATTERNS = [
        r'【答案】\s*(.+?)(?=【|$)',
        r'答案[：:]\s*(.+?)(?=【|$)',
        r'参考答案[：:]\s*(.+?)(?=【|$)',
        r'正确答案[：:]\s*(.+?)(?=【|$)',
    ]
    
    ANALYSIS_PATTERNS = [
        r'【解析】\s*(.+?)(?=【|$)',
        r'解析[：:]\s*(.+?)(?=【|$)',
        r'【分析】\s*(.+?)(?=【|$)',
        r'分析[：:]\s*(.+?)(?=【|$)',
        r'【解答】\s*(.+?)(?=【|$)',
    ]
    
    SCORE_PATTERNS = [
        r'[\(（](\d+)\s*分[\)）]',
        r'(\d+)\s*分',
        r'共\s*(\d+)\s*分',
    ]
    
    def __init__(self, file_id: str = None):
        self.current_type = QuestionType.UNKNOWN
        self.current_type_name = ""
        self.questions: List[Question] = []
        self.file_id = file_id
        self.pending_options: List[Tuple[str, str]] = []
        self.in_answer_section = False
        self.current_answer_question = None
        
    def split(self, paragraphs: List[Dict[str, Any]]) -> List[Question]:
        """
        将段落列表拆分为题目
        
        Args:
            paragraphs: 解析后的段落列表
            
        Returns:
            题目列表
        """
        self.questions = []
        self.current_type = QuestionType.UNKNOWN
        self.current_type_name = ""
        self.pending_options = []
        self.in_answer_section = False
        
        current_question = None
        current_paragraphs = []
        current_option_label = None
        current_option_content = []
        current_option_paras = []
        
        for para in paragraphs:
            text = para.get('text', '').strip()
            if not text:
                continue
            
            type_match = self._match_question_type(text)
            if type_match:
                if current_question:
                    if current_option_label and current_option_content:
                        self._add_option_to_question(
                            current_question, 
                            current_option_label, 
                            ' '.join(current_option_content),
                            current_option_paras
                        )
                        current_option_label = None
                        current_option_content = []
                        current_option_paras = []
                    self._finalize_question(current_question, current_paragraphs)
                    
                self.current_type = type_match[0]
                self.current_type_name = type_match[1]
                current_question = None
                current_paragraphs = []
                continue
            
            answer_match = self._match_answer(text)
            if answer_match:
                self.in_answer_section = True
                if current_question:
                    if current_option_label and current_option_content:
                        self._add_option_to_question(
                            current_question, 
                            current_option_label, 
                            ' '.join(current_option_content),
                            current_option_paras
                        )
                        current_option_label = None
                        current_option_content = []
                        current_option_paras = []
                    current_question.answer = answer_match.group(1).strip()
                continue
                
            analysis_match = self._match_analysis(text)
            if analysis_match:
                if current_question:
                    if current_option_label and current_option_content:
                        self._add_option_to_question(
                            current_question, 
                            current_option_label, 
                            ' '.join(current_option_content),
                            current_option_paras
                        )
                        current_option_label = None
                        current_option_content = []
                        current_option_paras = []
                    current_question.analysis = analysis_match.group(1).strip()
                continue
            
            score_match = self._match_score(text)
            if score_match and current_question:
                current_question.score = float(score_match.group(1))
            
            number_match = self._match_question_number(text)
            if number_match:
                if current_question:
                    if current_option_label and current_option_content:
                        self._add_option_to_question(
                            current_question, 
                            current_option_label, 
                            ' '.join(current_option_content),
                            current_option_paras
                        )
                        current_option_label = None
                        current_option_content = []
                        current_option_paras = []
                    self._finalize_question(current_question, current_paragraphs)
                    
                question_number = int(number_match.group(1))
                question_text = text[number_match.end():].strip()
                
                current_question = Question(
                    id=f"q_{len(self.questions)}",
                    number=question_number,
                    type=self.current_type,
                    type_name=self.current_type_name,
                    content="",
                    content_html=""
                )
                current_paragraphs = [para]
                self.in_answer_section = False
                
                options, remaining_text, pending_opt = self._extract_options_with_remainder(question_text)
                if options:
                    if self.current_type == QuestionType.UNKNOWN:
                        current_question.type = QuestionType.SINGLE_CHOICE
                        current_question.type_name = "选择题"
                    else:
                        current_question.type_name = self.current_type_name or "选择题"
                    current_question.options = options
                    current_question.content = remaining_text.strip()
                    if pending_opt:
                        current_option_label = pending_opt[0]
                        current_option_content = [pending_opt[1]]
                else:
                    current_question.content = question_text
                    current_question.type_name = self.current_type_name or "未知题型"
                    
            elif current_question:
                current_paragraphs.append(para)
                
                new_option = self._check_new_option(text)
                
                if new_option:
                    if current_option_label and current_option_content:
                        self._add_option_to_question(
                            current_question, 
                            current_option_label, 
                            ' '.join(current_option_content),
                            current_option_paras
                        )
                    
                    current_option_label = new_option[0]
                    current_option_content = [new_option[1]]
                    current_option_paras = [para]
                elif current_option_label:
                    current_option_content.append(text)
                    current_option_paras.append(para)
                else:
                    options, remaining_text, pending_opt = self._extract_options_with_remainder(text)
                    if options:
                        if current_question.type == QuestionType.UNKNOWN:
                            current_question.type = QuestionType.SINGLE_CHOICE
                            current_question.type_name = "选择题"
                        current_question.options.extend(options)
                        if pending_opt:
                            current_option_label = pending_opt[0]
                            current_option_content = [pending_opt[1]]
                    else:
                        current_question.content += " " + text
                    
        if current_question:
            if current_option_label and current_option_content:
                self._add_option_to_question(
                    current_question, 
                    current_option_label, 
                    ' '.join(current_option_content),
                    current_option_paras
                )
            self._finalize_question(current_question, current_paragraphs)
            
        return self.questions
        
    def _match_question_type(self, text: str) -> Optional[Tuple[QuestionType, str]]:
        """匹配题型标题"""
        for pattern, q_type, type_name in self.TYPE_PATTERNS:
            if re.match(pattern, text):
                return (q_type, type_name)
        return None
        
    def _match_question_number(self, text: str) -> Optional[re.Match]:
        """匹配题号"""
        for pattern in self.QUESTION_NUMBER_PATTERNS:
            match = re.match(pattern, text)
            if match:
                return match
        return None
        
    def _extract_options(self, text: str) -> List[Option]:
        """提取选择题选项"""
        options = []
        
        for pattern in self.OPTION_PATTERNS:
            matches = list(re.finditer(pattern, text))
            if len(matches) >= 2:
                for i, match in enumerate(matches):
                    label = match.group(1)
                    start = match.end()
                    end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                    content = text[start:end].strip()
                    
                    is_latex = '$' in content or '\\' in content
                    
                    options.append(Option(
                        label=label,
                        content=content,
                        is_latex=is_latex
                    ))
                    
                break
                
        return options
    
    def _extract_options_with_remainder(self, text: str) -> Tuple[List[Option], str, Optional[Tuple[str, str]]]:
        """
        提取选择题选项，并返回剩余文本和待处理的选项
        
        Returns:
            (选项列表, 剩余文本, 待处理选项(label, content)或None)
        """
        options = []
        first_option_pos = len(text)
        pending_option = None
        
        for pattern in self.OPTION_PATTERNS:
            matches = list(re.finditer(pattern, text))
            if matches:
                first_match = matches[0]
                first_option_pos = first_match.start()
                
                for i, match in enumerate(matches):
                    label = match.group(1)
                    start = match.end()
                    
                    if i + 1 < len(matches):
                        end = matches[i + 1].start()
                        content = text[start:end].strip()
                        is_latex = '$' in content or '\\' in content
                        options.append(Option(label=label, content=content, is_latex=is_latex))
                    else:
                        content = text[start:].strip()
                        if content:
                            pending_option = (label, content)
                
                break
        
        remainder = text[:first_option_pos].strip()
        return options, remainder, pending_option
    
    def _check_new_option(self, text: str) -> Optional[Tuple[str, str]]:
        """
        检查文本是否是新选项的开始
        
        Returns:
            (选项标签, 选项内容) 或 None
        """
        for pattern in self.OPTION_PATTERNS:
            match = re.match(pattern, text)
            if match:
                label = match.group(1)
                content = text[match.end():].strip()
                return (label, content)
        return None
    
    def _add_option_to_question(self, question: Question, label: str, content: str, paras: List[Dict] = None):
        """添加选项到题目"""
        is_latex = '$' in content or '\\' in content
        content_html = f'<span class="option-label">{label}.</span> '
        
        if paras:
            for para in paras:
                para_html = self._paragraph_to_html(para)
                if para_html:
                    option_match = re.match(rf'^{label}\s*[\.、．:：]\s*', para_html)
                    if option_match:
                        content_html += para_html[option_match.end():]
                    else:
                        content_html += para_html
        else:
            content_html += self._escape_html(content)
        
        question.options.append(Option(
            label=label,
            content=content,
            content_html=content_html,
            is_latex=is_latex
        ))
    
    def _match_answer(self, text: str) -> Optional[re.Match]:
        """匹配答案"""
        for pattern in self.ANSWER_PATTERNS:
            match = re.search(pattern, text, re.DOTALL)
            if match:
                return match
        return None
    
    def _match_analysis(self, text: str) -> Optional[re.Match]:
        """匹配解析"""
        for pattern in self.ANALYSIS_PATTERNS:
            match = re.search(pattern, text, re.DOTALL)
            if match:
                return match
        return None
    
    def _match_score(self, text: str) -> Optional[re.Match]:
        """匹配分值"""
        for pattern in self.SCORE_PATTERNS:
            match = re.search(pattern, text)
            if match:
                return match
        return None
        
    def _finalize_question(self, question: Question, paragraphs: List[Dict]):
        """完成题目处理"""
        question.raw_paragraphs = paragraphs
        
        content_parts = []
        latex_formulas = []
        
        for para in paragraphs:
            html_part = self._paragraph_to_html(para)
            if html_part:
                content_parts.append(html_part)
                
            for item in para.get('content_items', []):
                if item['type'] in ['latex', 'latex_block']:
                    latex_formulas.append(item['content'])
                    
        question.content_html = ' '.join(content_parts)
        question.latex_formulas = latex_formulas
        
        for para in paragraphs:
            for item in para.get('content_items', []):
                if item['type'] == 'image':
                    img_info = item['content']
                    if isinstance(img_info, dict) and 'filename' in img_info:
                        question.images.append(img_info['filename'])
        
        self._clean_question_content(question)
        self.questions.append(question)
    
    def _clean_question_content(self, question: Question):
        """清理题目内容，移除答案和解析标记"""
        answer_keywords = ['【答案】', '答案：', '答案:', '参考答案：', '参考答案:', '正确答案：', '正确答案:']
        analysis_keywords = ['【解析】', '解析：', '解析:', '【分析】', '分析：', '分析:', '【解答】', '解答：', '解答:']
        
        for keyword in answer_keywords + analysis_keywords:
            if keyword in question.content:
                question.content = question.content.split(keyword)[0].strip()
        
        question.content = re.sub(r'\s+', ' ', question.content).strip()
        
    def _paragraph_to_html(self, para: Dict) -> str:
        """将段落转换为 HTML"""
        parts = []
        
        for item in para.get('content_items', []):
            item_type = item['type']
            content = item['content']
            
            if item_type == 'text':
                content = self._escape_html(content)
                parts.append(content)
                
            elif item_type == 'latex':
                parts.append(self._format_inline_formula(content))
                
            elif item_type == 'latex_block':
                parts.append(self._format_block_formula(content))
                
            elif item_type == 'image':
                if isinstance(content, dict):
                    img_src = content.get('filename', '')
                    parts.append(self._format_image(img_src, 'question-image'))
                    
        return ''.join(parts)
    
    def _escape_html(self, text: str) -> str:
        """转义 HTML 特殊字符"""
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        return text
    
    def _format_inline_formula(self, latex: str) -> str:
        """格式化行内公式"""
        latex = latex.strip()
        return f'<span class="math-inline" data-latex="{self._escape_html(latex)}">${latex}$</span>'
    
    def _format_block_formula(self, latex: str) -> str:
        """格式化块级公式"""
        latex = latex.strip()
        return f'<div class="math-block" data-latex="{self._escape_html(latex)}">$${latex}$$</div>'
    
    def _format_image(self, filename: str, css_class: str, alt: str = '') -> str:
        """格式化图片标签"""
        if self.file_id:
            img_src = f"http://localhost:3000/api/v1/images/{self.file_id}/{filename}"
        else:
            img_src = f"http://localhost:3000/api/v1/images/{filename}"
        
        alt_attr = f' alt="{alt}"' if alt else ''
        return f'<img src="{img_src}" class="{css_class}"{alt_attr} />'


def split_questions(paragraphs: List[Dict[str, Any]], file_id: str = None) -> List[Question]:
    """
    便捷函数：拆分题目

    Args:
        paragraphs: 解析后的段落列表
        file_id: 文件ID，用于生成正确的图片URL

    Returns:
        题目列表
    """
    splitter = QuestionSplitter(file_id=file_id)
    return splitter.split(paragraphs)
