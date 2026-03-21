"""
题目拆分引擎 - 优化版
将解析后的文档内容按题目拆分为独立单元
改进：更准确的题号识别、选项解析、答案提取
"""

import re
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum


class QuestionType(Enum):
    """题目类型"""
    UNKNOWN = "unknown"
    SINGLE_CHOICE = "single_choice"
    MULTIPLE_CHOICE = "multiple_choice"
    FILL_BLANK = "fill_blank"
    TRUE_FALSE = "true_false"
    SHORT_ANSWER = "short_answer"
    CALCULATION = "calculation"
    PROOF = "proof"
    COMPREHENSIVE = "comprehensive"


@dataclass
class Option:
    """选择题选项"""
    label: str
    content: str
    content_html: str = ""
    is_latex: bool = False
    images: List[str] = field(default_factory=list)


@dataclass
class Question:
    """题目数据结构"""
    id: str
    number: int
    type: QuestionType
    type_name: str
    content: str
    content_html: str
    options: List[Option] = field(default_factory=list)
    answer: Optional[str] = None
    analysis: Optional[str] = None
    score: Optional[float] = None
    difficulty: Optional[str] = None
    images: List[str] = field(default_factory=list)
    latex_formulas: List[str] = field(default_factory=list)
    raw_paragraphs: List[Dict] = field(default_factory=list)


class QuestionSplitter:
    """题目拆分器 - 优化版"""

    TYPE_PATTERNS = [
        (r'^[一二三四五六七八九十]+[、\.．]\s*选择题', QuestionType.SINGLE_CHOICE, '选择题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*单选题', QuestionType.SINGLE_CHOICE, '单选题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*多选题', QuestionType.MULTIPLE_CHOICE, '多选题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*填空题', QuestionType.FILL_BLANK, '填空题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*判断题', QuestionType.TRUE_FALSE, '判断题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*简答题', QuestionType.SHORT_ANSWER, '简答题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*计算题', QuestionType.CALCULATION, '计算题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*证明题', QuestionType.PROOF, '证明题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*解答题', QuestionType.CALCULATION, '解答题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*综合题', QuestionType.COMPREHENSIVE, '综合题'),
        (r'^[一二三四五六七八九十]+[、\.．]\s*应用题', QuestionType.CALCULATION, '应用题'),
    ]
    
    ANSWER_SECTION_MARKERS = [
        r'^参考答案',
        r'^答案与解析',
        r'^答案解析',
        r'^试题解析',
        r'^详细解析',
        r'^【参考答案】',
        r'^【答案与解析】',
        r'.*参考答案.*$',
        r'.*答案与解析.*$',
    ]
    
    QUESTION_NUMBER_PATTERN = r'^\s*(\d+)\s*[\.、．]\s*'
    
    OPTION_PATTERN = r'([A-H])\s*[\.、．]\s*'
    
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
        self.in_answer_section = False
        
    def split(self, paragraphs: List[Dict[str, Any]]) -> List[Question]:
        """将段落列表拆分为题目"""
        self.questions = []
        self.current_type = QuestionType.UNKNOWN
        self.current_type_name = ""
        self.in_answer_section = False
        
        current_question = None
        current_paragraphs = []
        current_options_data = []
        
        i = 0
        while i < len(paragraphs):
            para = paragraphs[i]
            text = para.get('text', '').strip()
            
            if not text and not self._has_content(para):
                i += 1
                continue
            
            type_match = self._match_question_type(text)
            if type_match:
                if current_question:
                    self._finalize_question(current_question, current_paragraphs, current_options_data)
                    
                self.current_type = type_match[0]
                self.current_type_name = type_match[1]
                current_question = None
                current_paragraphs = []
                current_options_data = []
                i += 1
                continue
            
            answer_section_match = self._match_answer_section(text)
            if answer_section_match:
                self.in_answer_section = True
                i += 1
                continue
            
            answer_match = self._match_answer(text)
            if answer_match:
                if current_question:
                    current_question.answer = answer_match.group(1).strip()
                i += 1
                continue
                
            analysis_match = self._match_analysis(text)
            if analysis_match:
                if current_question:
                    current_question.analysis = analysis_match.group(1).strip()
                i += 1
                continue
            
            score_match = self._match_score(text)
            if score_match and current_question:
                current_question.score = float(score_match.group(1))
                i += 1
                continue
            
            number_match = re.match(self.QUESTION_NUMBER_PATTERN, text)
            if number_match:
                if current_question:
                    self._finalize_question(current_question, current_paragraphs, current_options_data)
                
                question_number = int(number_match.group(1))
                question_text = text[number_match.end():].strip()
                
                if self.in_answer_section:
                    existing_q = self._find_question_by_number(question_number)
                    if existing_q:
                        answer_content = self._extract_answer_content(question_text, paragraphs[i+1:])
                        if answer_content:
                            if not existing_q.answer:
                                existing_q.answer = answer_content.get('answer', '')
                            if not existing_q.analysis:
                                existing_q.analysis = answer_content.get('analysis', '')
                        current_question = None
                        current_paragraphs = []
                        current_options_data = []
                        i += 1
                        continue
                
                current_question = Question(
                    id=f"q_{len(self.questions)}",
                    number=question_number,
                    type=self.current_type,
                    type_name=self.current_type_name,
                    content="",
                    content_html=""
                )
                current_paragraphs = [para]
                current_options_data = []
                self.in_answer_section = False
                
                options, remaining_text = self._extract_options_from_text(question_text, para)
                if options:
                    if self.current_type == QuestionType.UNKNOWN:
                        current_question.type = QuestionType.SINGLE_CHOICE
                        current_question.type_name = "选择题"
                    current_options_data = options
                    current_question.content = remaining_text.strip()
                else:
                    current_question.content = question_text
                    
                i += 1
                continue
            
            if current_question:
                options = self._extract_options_from_text(text, para)[0]
                
                if options:
                    if current_question.type == QuestionType.UNKNOWN:
                        current_question.type = QuestionType.SINGLE_CHOICE
                        current_question.type_name = "选择题"
                    current_options_data.extend(options)
                else:
                    current_paragraphs.append(para)
                    current_question.content += " " + text
            
            i += 1
        
        if current_question:
            self._finalize_question(current_question, current_paragraphs, current_options_data)
            
        return self.questions
    
    def _has_content(self, para: Dict) -> bool:
        """检查段落是否有实际内容（图片或公式）"""
        for item in para.get('content_items', []):
            if item.get('type') in ['image', 'latex', 'latex_block']:
                return True
        return False
    
    def _match_question_type(self, text: str) -> Optional[Tuple[QuestionType, str]]:
        """匹配题型标题"""
        for pattern, q_type, type_name in self.TYPE_PATTERNS:
            if re.match(pattern, text):
                return (q_type, type_name)
        return None
    
    def _match_answer_section(self, text: str) -> Optional[re.Match]:
        """匹配答案部分开始标记"""
        for pattern in self.ANSWER_SECTION_MARKERS:
            match = re.match(pattern, text)
            if match:
                return match
        return None
    
    def _find_question_by_number(self, number: int) -> Optional[Question]:
        """根据题号查找已存在的题目"""
        for q in self.questions:
            if q.number == number:
                return q
        return None
    
    def _extract_answer_content(self, text: str, remaining_paragraphs: List[Dict]) -> Dict[str, str]:
        """从答案部分提取答案和解析内容"""
        result = {'answer': '', 'analysis': ''}
        
        answer_patterns = [
            r'^([A-H])\s*[\[【]',
            r'^([A-H])\s*$',
        ]
        
        for pattern in answer_patterns:
            match = re.match(pattern, text)
            if match:
                result['answer'] = match.group(1)
                break
        
        if remaining_paragraphs and len(remaining_paragraphs) > 0:
            analysis_parts = []
            for para in remaining_paragraphs[:10]:
                next_text = para.get('text', '').strip()
                
                if re.match(r'^\d+\s*[\.、．]', next_text):
                    break
                
                if next_text.startswith('【'):
                    if '【分析】' in next_text or '【解析】' in next_text or '【详解】' in next_text:
                        match = re.search(r'【分析】|【解析】|【详解】', next_text)
                        if match:
                            analysis_parts.append(next_text[match.end():].strip())
                elif next_text and not next_text.startswith('故选'):
                    analysis_parts.append(next_text)
            
            if analysis_parts:
                result['analysis'] = ' '.join(analysis_parts)
        
        return result
    
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
    
    def _extract_options_from_text(self, text: str, para: Dict) -> Tuple[List[Tuple[str, str, Dict]], str]:
        """
        从文本中提取选项
        
        Returns:
            (选项列表[(label, content, para)], 剩余文本)
        """
        options = []
        matches = list(re.finditer(self.OPTION_PATTERN, text))
        
        if not matches:
            return [], text
        
        first_option_pos = matches[0].start()
        
        for i, match in enumerate(matches):
            label = match.group(1)
            start = match.end()
            
            if i + 1 < len(matches):
                end = matches[i + 1].start()
            else:
                end = len(text)
            
            content = text[start:end].strip()
            options.append((label, content, para))
        
        remainder = text[:first_option_pos].strip()
        return options, remainder
    
    def _finalize_question(self, question: Question, paragraphs: List[Dict], 
                          options_data: List[Tuple[str, str, Dict]]):
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
                elif item['type'] == 'image':
                    img_info = item['content']
                    if isinstance(img_info, dict) and 'filename' in img_info:
                        question.images.append(img_info['filename'])
        
        question.content_html = ' '.join(content_parts)
        question.latex_formulas = latex_formulas
        
        for label, content, para in options_data:
            option_html = f'<span class="option-label">{label}.</span> '
            option_images = []
            
            para_html = self._paragraph_to_html(para)
            if para_html:
                option_match = re.match(rf'^{label}\s*[\.、．]\s*', para_html)
                if option_match:
                    option_html += para_html[option_match.end():]
                else:
                    option_html += self._escape_html(content)
            else:
                option_html += self._escape_html(content)
            
            for item in para.get('content_items', []):
                if item['type'] == 'image':
                    img_info = item['content']
                    if isinstance(img_info, dict) and 'filename' in img_info:
                        option_images.append(img_info['filename'])
                        question.images.append(img_info['filename'])
            
            is_latex = '$' in content or '\\' in content or bool(option_images)
            
            question.options.append(Option(
                label=label,
                content=content,
                content_html=option_html,
                is_latex=is_latex,
                images=option_images
            ))
        
        self._clean_question_content(question)
        self.questions.append(question)
    
    def _clean_question_content(self, question: Question):
        """清理题目内容"""
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
                    parts.append(self._format_image(img_src, 'formula-image'))
        
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
    """便捷函数：拆分题目"""
    splitter = QuestionSplitter(file_id=file_id)
    return splitter.split(paragraphs)
