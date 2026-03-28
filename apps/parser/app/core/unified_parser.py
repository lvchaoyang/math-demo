"""
统一文档解析器
整合 Pandoc 主方案和图片降级处理
"""

import zipfile
import tempfile
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import logging
import re

from .pandoc_converter import EnhancedPandocConverter
from .image_fallback import ImageFallbackProcessor
from .docx_to_html import convert_docx_to_html

logger = logging.getLogger(__name__)


class UnifiedDocxParser:
    """统一 DOCX 解析器 - Pandoc 优先 + 图片降级"""
    
    def __init__(self):
        self.pandoc = EnhancedPandocConverter()
        self.image_processor = ImageFallbackProcessor()
        
    def parse(
        self,
        docx_path: str,
        extract_images: bool = True,
        image_output_dir: Optional[str] = None,
        file_id: Optional[str] = None
    ) -> Tuple[bool, Dict[str, Any], str]:
        """
        解析 DOCX 文档
        
        Args:
            docx_path: DOCX 文件路径
            extract_images: 是否提取图片
            image_output_dir: 图片输出目录
            file_id: 文件 ID
            
        Returns:
            (成功标志, 解析结果, 错误信息)
        """
        docx_path = Path(docx_path)
        if not docx_path.exists():
            return False, {}, f"文件不存在: {docx_path}"
        
        if file_id is None:
            file_id = docx_path.stem
        
        result = {
            'file_id': file_id,
            'paragraphs': [],
            'images': {},
            'formulas': [],
            'html_content': '',
            'method': 'unknown'
        }
        
        if self.pandoc.is_available():
            success, data, error = self._parse_with_pandoc(
                docx_path, extract_images, image_output_dir, file_id
            )
            if success:
                result.update(data)
                result['method'] = 'pandoc'
                return True, result, ""
            else:
                logger.warning(f"Pandoc 解析失败，尝试降级: {error}")
        
        success, data, error = self._parse_with_fallback(
            docx_path, extract_images, image_output_dir, file_id
        )
        if success:
            result.update(data)
            result['method'] = 'fallback'
            return True, result, ""
        
        return False, {}, error
    
    def _parse_with_pandoc(
        self,
        docx_path: Path,
        extract_images: bool,
        image_output_dir: Optional[str],
        file_id: str
    ) -> Tuple[bool, Dict[str, Any], str]:
        """使用 Pandoc 解析"""
        media_dir = image_output_dir if extract_images else None
        
        success, html_content, metadata = self.pandoc.convert_with_strategy(
            str(docx_path),
            strategy='html_mathjax',
            extract_media=extract_images,
            media_dir=media_dir
        )
        
        if not success:
            return False, {}, html_content
        
        try:
            # 替换 HTML 中的图片路径为 API 路径
            if extract_images and metadata.get('media_dir'):
                media_path = Path(metadata['media_dir'])
                if media_path.exists():
                    # 替换所有图片路径
                    for img_file in media_path.rglob('*'):
                        if img_file.is_file():
                            # 替换绝对路径为 API 路径
                            # 例如：/Users/.../data/images/file_id/media/image1.png -> /api/v1/images/file_id/image1.png
                            old_path = str(img_file)
                            new_path = f"/api/v1/images/{file_id}/{img_file.name}"
                            html_content = html_content.replace(old_path, new_path)

            # 移除 pandoc 生成的图片宽高样式，避免公式被压缩变形
            html_content = re.sub(
                r'<img([^>]*?)\sstyle="[^"]*?(?:width|height)\s*:[^"]*?"([^>]*?)>',
                r'<img\1\2>',
                html_content,
                flags=re.IGNORECASE
            )
            html_content = re.sub(
                r'\s(?:width|height)="[^"]*"',
                '',
                html_content,
                flags=re.IGNORECASE
            )
            
            if not (html_content or '').strip():
                logger.warning('Pandoc 输出 HTML 为空，回退 docx_to_html')
                try:
                    html_content = convert_docx_to_html(str(docx_path))
                except Exception as e:
                    logger.error(f'docx_to_html 回退失败: {e}')

            paragraphs = self._parse_html_to_paragraphs(html_content)
            formulas = self._extract_formulas_from_html(html_content)
            
            images = {}
            if extract_images and metadata.get('media_dir'):
                media_path = Path(metadata['media_dir'])
                if media_path.exists():
                    for img_file in media_path.rglob('*'):
                        if img_file.is_file():
                            # 返回文件名，前端会拼接成 /images/{file_id}/{filename}
                            images[img_file.name] = img_file.name
            
            return True, {
                'paragraphs': paragraphs,
                'images': images,
                'formulas': formulas,
                'html_content': html_content
            }, ""
            
        except Exception as e:
            return False, {}, f"解析 HTML 失败: {str(e)}"
    
    def _parse_with_fallback(
        self,
        docx_path: Path,
        extract_images: bool,
        image_output_dir: Optional[str],
        file_id: str
    ) -> Tuple[bool, Dict[str, Any], str]:
        """降级解析方案"""
        try:
            paragraphs = []
            images = {}
            formulas = []
            
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                document_xml = docx_zip.read('word/document.xml')
                
                paragraphs = self._parse_document_xml(document_xml)
                
                if extract_images and image_output_dir:
                    success, img_map, error = self.image_processor.process_docx_images(
                        str(docx_path),
                        image_output_dir,
                        file_id
                    )
                    if success:
                        images = img_map
            
            # Pandoc 不可用时此前返回空 HTML，导致「整体 HTML」模式前端无内容
            html_content = ''
            try:
                html_content = convert_docx_to_html(str(docx_path))
            except Exception as conv_err:
                logger.warning(f"降级路径 docx_to_html 失败，改用段落拼接: {conv_err}")
            if not (html_content or '').strip():
                body = '\n'.join(
                    p.get('content_html') or f"<p>{p.get('text', '')}</p>"
                    for p in paragraphs
                )
                html_content = self._minimal_preview_html(body)

            return True, {
                'paragraphs': paragraphs,
                'images': images,
                'formulas': formulas,
                'html_content': html_content
            }, ""
            
        except Exception as e:
            return False, {}, f"降级解析失败: {str(e)}"
    
    def _parse_html_to_paragraphs(self, html_content: str) -> List[Dict[str, Any]]:
        """解析 HTML 为段落列表"""
        paragraphs = []
        
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(html_content, 'html.parser')
            
            for elem in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'li']):
                text = elem.get_text(strip=True)
                
                latex_formulas = self._extract_inline_formulas(str(elem))
                
                images = []
                for img in elem.find_all('img'):
                    src = img.get('src', '')
                    if src:
                        images.append(src)
                
                if text or latex_formulas or images:
                    paragraphs.append({
                        'text': text,
                        'content_html': str(elem),
                        'latex_formulas': [f['latex'] for f in latex_formulas],
                        'images': images,
                        'type': elem.name
                    })
                    
        except ImportError:
            logger.warning("BeautifulSoup 未安装，使用简单解析")
            paragraphs = self._simple_html_parse(html_content)
        
        return paragraphs
    
    def _simple_html_parse(self, html_content: str) -> List[Dict[str, Any]]:
        """简单 HTML 解析"""
        paragraphs = []
        
        p_pattern = r'<(p|h[1-6]|div)[^>]*>(.*?)</\1>'
        for match in re.finditer(p_pattern, html_content, re.DOTALL | re.IGNORECASE):
            content = match.group(2)
            text = re.sub(r'<[^>]+>', '', content).strip()
            
            if text:
                paragraphs.append({
                    'text': text,
                    'content_html': match.group(0),
                    'latex_formulas': [],
                    'images': [],
                    'type': match.group(1).lower()
                })
        
        return paragraphs
    
    def _extract_formulas_from_html(self, html_content: str) -> List[Dict[str, str]]:
        """从 HTML 中提取公式"""
        formulas = []
        
        inline_pattern = r'<script[^>]*type=["\']math/tex["\'][^>]*>(.*?)</script>'
        for match in re.finditer(inline_pattern, html_content, re.DOTALL):
            formulas.append({
                'type': 'inline',
                'latex': match.group(1).strip()
            })
        
        display_pattern = r'<script[^>]*type=["\']math/tex;\s*mode=display["\'][^>]*>(.*?)</script>'
        for match in re.finditer(display_pattern, html_content, re.DOTALL):
            formulas.append({
                'type': 'display',
                'latex': match.group(1).strip()
            })
        
        return formulas
    
    def _extract_inline_formulas(self, html: str) -> List[Dict[str, str]]:
        """提取行内公式"""
        formulas = []
        
        pattern = r'<script[^>]*type=["\']math/tex[^"\']*["\'][^>]*>(.*?)</script>'
        for match in re.finditer(pattern, html, re.DOTALL):
            formulas.append({
                'type': 'inline',
                'latex': match.group(1).strip()
            })
        
        return formulas
    
    def _parse_document_xml(self, xml_content: bytes) -> List[Dict[str, Any]]:
        """解析 Word document.xml"""
        paragraphs = []
        
        try:
            from xml.etree import ElementTree as ET
            root = ET.fromstring(xml_content)
            
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            for para in root.findall('.//w:p', ns):
                text_parts = []
                for t in para.findall('.//w:t', ns):
                    if t.text:
                        text_parts.append(t.text)
                
                text = ''.join(text_parts).strip()
                if text:
                    paragraphs.append({
                        'text': text,
                        'content_html': f'<p>{text}</p>',
                        'latex_formulas': [],
                        'images': [],
                        'type': 'p'
                    })
                    
        except Exception as e:
            logger.error(f"解析 document.xml 失败: {e}")
        
        return paragraphs
    
    def _minimal_preview_html(self, body: str) -> str:
        """无 Pandoc / docx_to_html 失败时的最小可预览 HTML"""
        b = (body or '').strip()
        if not b:
            b = '<p>（未能从文档中提取可见内容）</p>'
        return (
            '<!DOCTYPE html>\n<html><head><meta charset="UTF-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1">'
            '<style>body{font-family:"Times New Roman","SimSun",serif;font-size:12pt;'
            'line-height:1.6;margin:40px;color:#333;}</style></head><body>'
            f'{b}</body></html>'
        )
    
    def _generate_optimized_html(self, formula_result: dict, file_id: str) -> str:
        """根据公式解析结果生成优化的 HTML"""
        html_parts = ['<!DOCTYPE html>', '<html>', '<head>', '<meta charset="UTF-8">']
        
        # 添加 MathJax 支持
        html_parts.append('''
        <script>
        window.MathJax = {
            tex: {
                inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
                displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']]
            },
            svg: {
                fontCache: 'global'
            },
            startup: {
                typeset: true
            }
        };
        </script>
        <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
        ''')
        
        html_parts.append('<style>')
        html_parts.append(self._get_optimized_styles())
        html_parts.append('</style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        # 生成公式内容
        formulas = formula_result.get('formulas', [])
        for formula in formulas:
            status = formula.get('status', 'unknown')
            latex = formula.get('latex', '')
            image_path = formula.get('rendered_image') or formula.get('image_path')
            
            if image_path and Path(image_path).exists():
                img_filename = Path(image_path).name
                html_parts.append(f'<div class="formula-container" data-status="{status}">')
                html_parts.append(f'<img src="/api/v1/images/{file_id}/{img_filename}" class="formula-image" alt="formula" />')
                html_parts.append('</div>')
            elif latex:
                if formula.get('type') in ['block', 'environment']:
                    html_parts.append(f'<div class="math-display">$$${latex}$$$</div>')
                else:
                    html_parts.append(f'<span class="math-inline">$$${latex}$$$</span>')
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        return '\n'.join(html_parts)
    
    def _get_optimized_styles(self) -> str:
        """获取优化的样式表"""
        return '''
        body {
            font-family: "Times New Roman", "SimSun", serif;
            font-size: 12pt;
            line-height: 1.6;
            margin: 40px;
            color: #333;
            text-rendering: optimizeLegibility;
        }
        .formula-container {
            margin: 1em 0;
            text-align: center;
        }
        .formula-image {
            max-height: 1.5em;
            vertical-align: -0.1em;
            margin: 0 2px;
        }
        .math-inline {
            display: inline-block;
            vertical-align: -0.1em;
        }
        .math-display {
            display: block;
            text-align: center;
            margin: 1.5em 0;
            overflow-x: auto;
        }
        '''


def parse_docx_unified(
    docx_path: str,
    extract_images: bool = True,
    image_output_dir: Optional[str] = None,
    file_id: Optional[str] = None
) -> Tuple[bool, Dict[str, Any], str]:
    """
    统一解析 DOCX 的便捷函数
    
    Args:
        docx_path: DOCX 文件路径
        extract_images: 是否提取图片
        image_output_dir: 图片输出目录
        file_id: 文件 ID
        
    Returns:
        (成功标志, 解析结果, 错误信息)
    """
    parser = UnifiedDocxParser()
    return parser.parse(docx_path, extract_images, image_output_dir, file_id)
