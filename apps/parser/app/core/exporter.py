"""
Word 导出器
将选中的题目导出为 Word 文档，支持水印
"""

from typing import List, Dict, Any, Optional
from pathlib import Path
import io

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement, parse_xml

from .splitter import Question


class WatermarkType:
    """水印类型"""
    TEXT = "text"
    IMAGE = "image"


class WordExporter:
    """Word 文档导出器"""
    
    def __init__(self):
        self.document = None
        self.image_dir: Optional[Path] = None
        
    def create_document(self, title: str = "", watermark_text: str = None, image_dir: Optional[str] = None) -> Document:
        """
        创建新文档
        
        Args:
            title: 文档标题
            watermark_text: 水印文字
            
        Returns:
            Document 对象
        """
        self.document = Document()
        self.image_dir = Path(image_dir) if image_dir else None
        
        # 设置默认字体
        self._set_default_font()
        
        # 添加标题
        if title:
            self._add_title(title)
            
        # 添加水印
        if watermark_text:
            self._add_watermark(watermark_text)
            
        return self.document
        
    def _set_default_font(self):
        """设置默认字体"""
        # 设置正文样式
        style = self.document.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        
        # 设置中文字体
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        
    def _add_title(self, title: str):
        """添加文档标题"""
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(title)
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        
        # 添加空行
        self.document.add_paragraph()
        
    def _add_watermark(self, text: str):
        """
        添加文字水印
        
        通过修改文档的 header 来实现水印效果
        """
        # 获取或创建 header
        section = self.document.sections[0]
        header = section.header
        
        # 创建水印的 XML
        watermark_xml = f'''
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:w10="urn:schemas-microsoft-com:office:word">
            <w:pict>
                <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" 
                    path="m@7,l@8,m@5,21600l@6,21600e">
                    <v:formulas>
                        <v:f eqn="sum #0 0 10800"/>
                        <v:f eqn="prod #0 2 1"/>
                        <v:f eqn="sum 21600 0 @1"/>
                        <v:f eqn="sum 0 0 @2"/>
                        <v:f eqn="sum 21600 0 @3"/>
                        <v:f eqn="if @0 @3 0"/>
                        <v:f eqn="if @0 21600 @1"/>
                        <v:f eqn="if @0 0 @2"/>
                        <v:f eqn="if @0 @4 21600"/>
                        <v:f eqn="mid @5 @6"/>
                        <v:f eqn="mid @8 @5"/>
                        <v:f eqn="mid @7 @8"/>
                        <v:f eqn="mid @6 @7"/>
                        <v:f eqn="sum @6 0 @5"/>
                    </v:formulas>
                    <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
                    <v:textpath on="t" fitshape="t"/>
                    <v:handles>
                        <v:h position="#0,bottomRight" xrange="6629,14971"/>
                    </v:handles>
                    <o:lock v:ext="edit" text="t" shapetype="t"/>
                </v:shapetype>
                <v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" 
                    style="position:absolute;margin-left:0;margin-top:0;width:500pt;height:100pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                    o:allowincell="f" fillcolor="silver" stroked="f">
                    <v:fill opacity=".5"/>
                    <v:textpath style="font-family:&quot;宋体&quot;;font-size:60pt" string="{text}"/>
                    <w10:wrap anchorx="margin" anchory="margin"/>
                </v:shape>
            </w:pict>
        </w:r>
        '''
        
        # 添加水印到 header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        run = paragraph.add_run()
        
        # 解析并插入水印 XML
        # python-docx 没有 OxmlElement.fromstring，这里使用 parse_xml。
        try:
            paragraph._p.append(parse_xml(watermark_xml))
        except Exception:
            # 水印失败不应阻断导出主流程
            pass
        
    def add_question(self, question: Question, include_answer: bool = False, 
                     include_analysis: bool = False):
        """
        添加题目到文档
        
        Args:
            question: 题目对象
            include_answer: 是否包含答案
            include_analysis: 是否包含解析
        """
        # 题号
        para = self.document.add_paragraph()
        run = para.add_run(f"{question.number}. ")
        run.font.bold = True
        run.font.size = Pt(12)
        
        # 题目内容
        self._add_formatted_text(para, question.content)
        self._add_question_images(question.images)
        
        # 选项（选择题）
        if question.options:
            self._add_options(question.options)
            
        # 答案
        if include_answer and question.answer:
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.3)
            run = para.add_run(f"【答案】{question.answer}")
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 128, 0)
            
        # 解析
        if include_analysis and question.analysis:
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.3)
            run = para.add_run(f"【解析】")
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 0, 255)
            self._add_formatted_text(para, question.analysis)
            
        # 添加空行
        self.document.add_paragraph()

    def _resolve_image_path(self, filename: str) -> Optional[Path]:
        """解析图片在本地磁盘中的真实路径"""
        if not self.image_dir:
            return None

        # 支持 filename 中可能带有 media/ 前缀
        candidates = [
            self.image_dir / filename,
            self.image_dir / "media" / filename,
        ]
        if filename.startswith("media/"):
            candidates.append(self.image_dir / filename.replace("media/", "", 1))

        for p in candidates:
            if p.exists() and p.is_file():
                return p
        return None

    def _add_question_images(self, images: List[str]):
        """将题目关联图片插入到文档"""
        if not images:
            return

        # 去重但保序，避免同图重复写入
        deduped = list(dict.fromkeys(images))
        for img in deduped:
            img_path = self._resolve_image_path(img)
            if not img_path:
                continue
            try:
                p = self.document.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(img_path), width=Inches(5.8))
            except Exception:
                # 单张图失败不影响整份文档导出
                continue
        
    def _add_formatted_text(self, paragraph, text: str):
        """添加格式化文本（处理 LaTeX 公式）"""
        # 简单的公式处理：将 $...$ 标记为斜体
        parts = text.split('$')
        
        for i, part in enumerate(parts):
            if i % 2 == 0:  # 普通文本
                if part:
                    run = paragraph.add_run(part)
                    run.font.size = Pt(12)
            else:  # 公式
                # 公式使用斜体
                run = paragraph.add_run(f" {part} ")
                run.italic = True
                run.font.size = Pt(12)
                
    def _add_options(self, options: List[Any]):
        """添加选项"""
        # 每行两个选项
        for i in range(0, len(options), 2):
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.5)
            
            # 第一个选项
            opt1 = options[i]
            run = para.add_run(f"{opt1.label}. ")
            run.font.bold = True
            self._add_formatted_text(para, opt1.content)
            
            # 第二个选项（如果有）
            if i + 1 < len(options):
                opt2 = options[i + 1]
                # 添加制表符分隔
                para.add_run('\t\t')
                run = para.add_run(f"{opt2.label}. ")
                run.font.bold = True
                self._add_formatted_text(para, opt2.content)
                
    def add_section_header(self, text: str):
        """添加章节标题（如"一、选择题"）"""
        para = self.document.add_paragraph()
        run = para.add_run(text)
        run.font.bold = True
        run.font.size = Pt(14)
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        para.paragraph_format.space_before = Pt(12)
        para.paragraph_format.space_after = Pt(6)
        
    def save(self, output_path: str):
        """保存文档"""
        self.document.save(output_path)
        
    def to_bytes(self) -> bytes:
        """导出为字节流"""
        buffer = io.BytesIO()
        self.document.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()


def export_questions(questions: List[Question], output_path: str, 
                     title: str = "", watermark: str = None,
                     image_dir: Optional[str] = None,
                     include_answer: bool = False, 
                     include_analysis: bool = False) -> str:
    """
    便捷函数：导出题目到 Word
    
    Args:
        questions: 题目列表
        output_path: 输出路径
        title: 文档标题
        watermark: 水印文字
        include_answer: 是否包含答案
        include_analysis: 是否包含解析
        
    Returns:
        输出文件路径
    """
    exporter = WordExporter()
    exporter.create_document(title=title, watermark_text=watermark, image_dir=image_dir)
    
    # 按题型分组
    current_type = None
    for question in questions:
        # 添加题型标题
        if question.type_name and question.type_name != current_type:
            exporter.add_section_header(f"{question.type_name}")
            current_type = question.type_name
            
        exporter.add_question(question, include_answer, include_analysis)
        
    exporter.save(output_path)
    return output_path
