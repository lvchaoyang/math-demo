"""
Word 文档转 HTML 解析器
将 docx 文件解析为完整的 HTML，图片转换为 base64 内嵌
"""
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import base64
import tempfile
import subprocess
import re
import logging

from .omml2latex import convert_omml_to_latex
from .wmf_converter import WMFConverter

logger = logging.getLogger(__name__)


class DocxToHtmlConverter:
    """Word 文档转 HTML 转换器"""
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'v': 'urn:schemas-microsoft-com:vml',
        'o': 'urn:schemas-microsoft-com:office:office',
    }
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.docx_zip = None
        self.relationships = {}
        self.media_cache = {}
        self.wmf_converter = WMFConverter()
        
    def __enter__(self):
        self.docx_zip = zipfile.ZipFile(self.file_path, 'r')
        self._load_relationships()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.docx_zip:
            self.docx_zip.close()
    
    def _load_relationships(self):
        """加载文档关系"""
        try:
            rels_xml = self.docx_zip.read('word/_rels/document.xml.rels')
            rels_root = ET.fromstring(rels_xml)
            
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type')
                self.relationships[rel_id] = {
                    'target': target,
                    'type': rel_type
                }
            
            print(f"[DEBUG] 加载了 {len(self.relationships)} 个关系")
            
        except Exception as e:
            logger.warning(f"加载关系失败: {e}")
    
    def convert_to_html(self) -> str:
        """将整个文档转换为 HTML"""
        document_xml = self.docx_zip.read('word/document.xml')
        root = ET.fromstring(document_xml)
        
        html_parts = ['<!DOCTYPE html>', '<html>', '<head>', '<meta charset="UTF-8">']
        html_parts.append('<style>')
        html_parts.append(self._get_default_styles())
        html_parts.append('</style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        body = root.find('.//w:body', self.NAMESPACES)
        if body is not None:
            for child in body:
                tag = self._get_tag_name(child.tag)
                
                if tag == 'p':
                    para_html = self._convert_paragraph_to_html(child)
                    if para_html.strip():
                        html_parts.append(para_html)
                elif tag == 'tbl':
                    table_html = self._convert_table_to_html(child)
                    html_parts.append(table_html)
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        # 打印标签统计
        if hasattr(self, '_tag_stats'):
            print(f"[DEBUG] 标签统计:")
            for tag, count in sorted(self._tag_stats.items(), key=lambda x: x[1], reverse=True)[:20]:
                print(f"  {tag}: {count}")
        
        return '\n'.join(html_parts)
    
    def _get_default_styles(self) -> str:
        """获取默认样式"""
        return '''
        body {
            font-family: "Times New Roman", "SimSun", serif;
            font-size: 12pt;
            line-height: 1.6;
            margin: 40px;
            color: #333;
            text-rendering: optimizeLegibility;
        }
        p {
            margin: 0.5em 0;
        }
        
        /* 行内公式图片：小尺寸，基线对齐 */
        .formula-image-inline {
            display: inline-block;
            vertical-align: -0.1em;
            max-height: 1.2em;
            margin: 0 2px;
        }
        
        /* 块级公式图片：独立显示，保持宽高比 */
        .formula-image-block {
            display: block;
            margin: 10px auto;
            max-width: 100%;
            height: auto;
        }
        
        /* 通用公式图片样式（兜底）*/
        .formula-image {
            display: inline-block;
            vertical-align: middle;
            max-width: 100%;
            height: auto;
            margin: 0 2px;
        }
        
        /* 行内公式：严格基线对齐，防止撑大行高 */
        .math-inline {
            display: inline-block;
            vertical-align: -0.1em;
            max-height: 1.2em;
            margin: 0 2px;
        }
        
        /* 块级公式：独立段落，居中，增加上下间距 */
        .math-display {
            display: block;
            text-align: center;
            margin: 1.5em 0;
            overflow-x: auto;
            overflow-y: hidden;
        }
        
        /* 针对 WMF 转换后的图片的特殊处理 */
        .formula-image-large {
            display: block;
            margin: 10px auto;
            max-width: 100%;
            height: auto;
        }
        
        h1, h2, h3, h4, h5, h6 {
            margin: 1em 0 0.5em;
            font-weight: bold;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }
        td, th {
            border: 1px solid #ccc;
            padding: 8px;
            text-align: left;
        }
        '''
    
    def _convert_paragraph_to_html(self, para_elem: ET.Element) -> str:
        """将段落转换为 HTML"""
        style = self._get_paragraph_style(para_elem)
        
        content_parts = []
        
        # 递归处理所有子元素
        for child in para_elem:
            content = self._convert_element_recursive(child)
            if content:
                content_parts.append(content)
        
        content_html = ''.join(content_parts)
        
        if not content_html.strip():
            return ''
        
        if style['is_heading']:
            tag = f'h{style["heading_level"]}'
            return f'<{tag}>{content_html}</{tag}>\n'
        else:
            align = style.get('alignment', 'left')
            align_style = f' text-align: {align};' if align != 'left' else ''
            return f'<p style="margin: 0.5em 0;{align_style}">{content_html}</p>\n'
    
    def _convert_element_recursive(self, elem: ET.Element, is_inside_formula=False) -> str:
        """递归转换元素及其子元素"""
        tag = self._get_tag_name(elem.tag)
        
        # 统计各种标签
        if not hasattr(self, '_tag_stats'):
            self._tag_stats = {}
        self._tag_stats[tag] = self._tag_stats.get(tag, 0) + 1
        
        # 【关键修复】如果已经在公式内部，忽略内部的 drawing/pict/object，避免重复渲染
        if is_inside_formula and tag in ['drawing', 'pict', 'object', 'imagedata']:
            return ""
        
        # 处理特定标签
        if tag == 'r':
            return self._convert_run_to_html(elem, is_inside_formula=is_inside_formula)
        elif tag == 'oMath' or tag == 'oMathPara':
            latex = self._parse_omml(elem)
            if tag == 'oMathPara':
                return f'<div class="math-display" data-latex="{self._escape_html(latex)}">$${latex}$$</div>'
            else:
                return f'<span class="math-inline" data-latex="{self._escape_html(latex)}">${latex}$</span>'
        elif tag == 'drawing':
            # 只有不在公式内才尝试转为图片
            return self._convert_drawing_to_html(elem)
        elif tag == 'pict':
            return self._convert_pict_to_html(elem)
        elif tag == 'object':
            return self._convert_ole_to_html(elem)
        elif tag == 'imagedata':
            return self._convert_imagedata_to_html(elem)
        elif tag == 't':
            text = elem.text or ''
            return self._escape_html(text)
        elif tag == 'tab':
            return '&nbsp;&nbsp;&nbsp;&nbsp;'
        elif tag == 'br':
            return '<br/>'
        
        # 对于其他元素，递归处理子元素
        content_parts = []
        for child in elem:
            # 如果当前是 oMath 的子元素，传递 True
            next_level_formula = is_inside_formula or (tag in ['oMath', 'oMathPara'])
            child_content = self._convert_element_recursive(child, is_inside_formula=next_level_formula)
            if child_content:
                content_parts.append(child_content)
        
        return ''.join(content_parts)
    
    def _convert_run_to_html(self, run_elem: ET.Element, is_inside_formula=False) -> str:
        """将 Run 元素转换为 HTML"""
        content_parts = []
        has_formula_image = False
        
        for child in run_elem:
            tag = self._get_tag_name(child.tag)
            
            if tag == 't':
                text = child.text or ''
                content_parts.append(('text', self._escape_html(text)))
            elif tag == 'tab':
                content_parts.append(('text', '&nbsp;&nbsp;&nbsp;&nbsp;'))
            elif tag == 'br':
                content_parts.append(('text', '<br/>'))
            elif tag == 'drawing':
                if not is_inside_formula:
                    img_html = self._convert_drawing_to_html(child)
                    if img_html:
                        content_parts.append(('image', img_html))
                        has_formula_image = True
            elif tag == 'pict':
                if not is_inside_formula:
                    img_html = self._convert_pict_to_html(child)
                    if img_html:
                        content_parts.append(('image', img_html))
                        has_formula_image = True
            elif tag == 'object':
                if not is_inside_formula:
                    obj_html = self._convert_ole_to_html(child)
                    if obj_html:
                        content_parts.append(('image', obj_html))
                        has_formula_image = True
            elif tag == 'imagedata':
                if not is_inside_formula:
                    img_html = self._convert_imagedata_to_html(child)
                    if img_html:
                        content_parts.append(('image', img_html))
                        has_formula_image = True
            else:
                # 对于其他标签，递归处理
                next_level_formula = is_inside_formula or (tag in ['oMath', 'oMathPara'])
                child_content = self._convert_element_recursive(child, is_inside_formula=next_level_formula)
                if child_content:
                    content_parts.append(('other', child_content))
        
        if not content_parts:
            return ''
        
        # 如果包含公式图片，不要应用 run 样式到图片上
        if has_formula_image:
            result_parts = []
            for item_type, content in content_parts:
                if item_type == 'image':
                    # 图片直接输出，不包裹 span
                    result_parts.append(content)
                elif item_type == 'text':
                    result_parts.append(content)
                else:
                    result_parts.append(content)
            
            content = ''.join(result_parts)
            
            # 如果还有文本内容，才应用样式
            r_pr = run_elem.find('w:rPr', self.NAMESPACES)
            if r_pr is not None:
                styles = []
                
                if r_pr.find('w:b', self.NAMESPACES) is not None:
                    styles.append('font-weight: bold')
                if r_pr.find('w:i', self.NAMESPACES) is not None:
                    styles.append('font-style: italic')
                if r_pr.find('w:u', self.NAMESPACES) is not None:
                    styles.append('text-decoration: underline')
                
                if styles:
                    # 只包裹文本内容，不包裹图片
                    text_content = ''.join([c for t, c in content_parts if t == 'text'])
                    image_content = ''.join([c for t, c in content_parts if t == 'image'])
                    other_content = ''.join([c for t, c in content_parts if t == 'other'])
                    
                    if text_content or other_content:
                        return f'{image_content}<span style="{"; ".join(styles)}">{text_content}{other_content}</span>'
                    else:
                        return image_content
        
        # 没有图片，正常处理
        content = ''.join([c for t, c in content_parts])
        
        r_pr = run_elem.find('w:rPr', self.NAMESPACES)
        if r_pr is not None:
            styles = []
            
            if r_pr.find('w:b', self.NAMESPACES) is not None:
                styles.append('font-weight: bold')
            if r_pr.find('w:i', self.NAMESPACES) is not None:
                styles.append('font-style: italic')
            if r_pr.find('w:u', self.NAMESPACES) is not None:
                styles.append('text-decoration: underline')
            
            if styles:
                return f'<span style="{"; ".join(styles)}">{content}</span>'
        
        return content
    
    def _convert_drawing_to_html(self, drawing_elem: ET.Element) -> str:
        """将 Drawing 元素转换为 HTML"""
        try:
            blip = drawing_elem.find('.//a:blip', self.NAMESPACES)
            if blip is None:
                blip = drawing_elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            
            if blip is not None:
                embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed_id and embed_id in self.relationships:
                    target = self.relationships[embed_id]['target']
                    return self._get_image_html(target)
        
        except Exception as e:
            logger.debug(f"解析 drawing 失败: {e}")
        
        return ''
    
    def _convert_pict_to_html(self, pict_elem: ET.Element) -> str:
        """将 Pict 元素（VML）转换为 HTML"""
        try:
            # 查找 imagedata 标签
            for elem in pict_elem.iter():
                tag = self._get_tag_name(elem.tag)
                if tag == 'imagedata':
                    # 获取 r:id 属性
                    ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
                    r_id = elem.get(f'{ns_r}id')
                    if not r_id:
                        # 尝试其他命名空间
                        for attr_name, attr_value in elem.attrib.items():
                            if 'id' in attr_name.lower():
                                r_id = attr_value
                                break
                    
                    if r_id and r_id in self.relationships:
                        target = self.relationships[r_id]['target']
                        return self._get_image_html(target)
                    else:
                        # 打印前几个警告
                        if not hasattr(self, '_pict_warning_count'):
                            self._pict_warning_count = 0
                        if self._pict_warning_count < 5:
                            print(f"[WARNING] pict 中未找到关系: r_id={r_id}")
                            self._pict_warning_count += 1
        
        except Exception as e:
            logger.error(f"解析 pict 失败: {e}")
        
        return ''
    
    def _convert_ole_to_html(self, ole_elem: ET.Element) -> str:
        """将 OLE 对象转换为 HTML"""
        try:
            # 首先查找 imagedata 标签（用于 MathType 等公式）
            for elem in ole_elem.iter():
                tag = self._get_tag_name(elem.tag)
                if tag == 'imagedata':
                    # 获取 r:id 属性
                    ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
                    r_id = elem.get(f'{ns_r}id')
                    if not r_id:
                        # 尝试其他属性
                        for attr_name, attr_value in elem.attrib.items():
                            if 'id' in attr_name.lower():
                                r_id = attr_value
                                break
                    
                    if r_id and r_id in self.relationships:
                        target = self.relationships[r_id]['target']
                        return self._get_image_html(target)
            
            # 如果没有 imagedata，尝试查找 OLEObject
            ole_obj = ole_elem.find('.//o:OLEObject', self.NAMESPACES)
            if ole_obj is None:
                ole_obj = ole_elem.find('.//{urn:schemas-microsoft-com:office:office}OLEObject')
            
            if ole_obj is not None:
                r_id = ole_obj.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if r_id and r_id in self.relationships:
                    target = self.relationships[r_id]['target']
                    return self._get_image_html(target)
        
        except Exception as e:
            logger.error(f"解析 OLE 失败: {e}")
        
        return ''
    
    def _convert_imagedata_to_html(self, imagedata_elem: ET.Element) -> str:
        """将 imagedata 元素转换为 HTML"""
        try:
            # 获取 r:id 属性
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            r_id = imagedata_elem.get(f'{ns_r}id')
            
            if not r_id:
                # 尝试其他属性
                for attr_name, attr_value in imagedata_elem.attrib.items():
                    if 'id' in attr_name.lower():
                        r_id = attr_value
                        break
            
            if r_id and r_id in self.relationships:
                target = self.relationships[r_id]['target']
                return self._get_image_html(target)
            else:
                # 只打印前几个警告
                if not hasattr(self, '_imagedata_warning_count'):
                    self._imagedata_warning_count = 0
                if self._imagedata_warning_count < 5:
                    print(f"[WARNING] 未找到关系: r_id={r_id}")
                    self._imagedata_warning_count += 1
        
        except Exception as e:
            logger.error(f"解析 imagedata 失败: {e}")
        
        return ''
    
    def _get_image_html(self, target: str) -> str:
        """获取图片的 HTML（base64 内嵌）"""
        if target in self.media_cache:
            return self.media_cache[target]
        
        try:
            if target.startswith('media/'):
                media_path = f'word/{target}'
            elif target.startswith('embeddings/'):
                media_path = f'word/{target}'
            else:
                media_path = f'word/{target}'
            
            if media_path not in self.docx_zip.namelist():
                logger.warning(f"媒体文件不存在: {media_path}")
                return ''
            
            content = self.docx_zip.read(media_path)
            
            file_ext = Path(target).suffix.lower()
            
            if file_ext in {'.wmf', '.emf'}:
                base64_data = self._convert_wmf_to_base64(content)
                if not base64_data:
                    # WMF 转换失败，使用原始 WMF 作为兜底
                    logger.warning(f"WMF 转换失败，使用原始 WMF: {target}")
                    base64_wmf = base64.b64encode(content).decode('utf-8')
                    base64_data = f"data:image/wmf;base64,{base64_wmf}"
                # WMF 图片使用块级样式
                img_class = "formula-image-block"
            else:
                mime_type = self._get_mime_type(file_ext)
                base64_data = f"data:{mime_type};base64,{base64.b64encode(content).decode('utf-8')}"
                img_class = "formula-image"
            
            img_html = f'<img src="{base64_data}" class="{img_class}" alt="formula" />'
            self.media_cache[target] = img_html
            return img_html
            
        except Exception as e:
            logger.error(f"获取图片失败: {e}")
            return ''
    
    def _convert_wmf_to_base64(self, wmf_content: bytes) -> str:
        """将 WMF 内容转换为 base64 PNG"""
        try:
            with tempfile.NamedTemporaryFile(suffix='.wmf', delete=False) as temp_wmf:
                temp_wmf.write(wmf_content)
                temp_wmf_path = temp_wmf.name
            
            temp_png_path = temp_wmf_path.replace('.wmf', '.png')
            
            success, result = self.wmf_converter.convert(temp_wmf_path, temp_png_path)
            
            Path(temp_wmf_path).unlink()
            
            if success and Path(temp_png_path).exists():
                with open(temp_png_path, 'rb') as f:
                    png_content = f.read()
                
                Path(temp_png_path).unlink()
                
                base64_str = base64.b64encode(png_content).decode('utf-8')
                return f"data:image/png;base64,{base64_str}"
            else:
                logger.warning(f"WMF 转换失败: {result}")
                return ''
                
        except Exception as e:
            logger.error(f"WMF 转 base64 失败: {e}")
            return ''
    
    def _get_mime_type(self, ext: str) -> str:
        """获取 MIME 类型"""
        mime_types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.tiff': 'image/tiff',
        }
        return mime_types.get(ext, 'image/png')
    
    def _convert_table_to_html(self, table_elem: ET.Element) -> str:
        """将表格转换为 HTML"""
        html_parts = ['<table>']
        
        for row in table_elem.findall('.//w:tr', self.NAMESPACES):
            html_parts.append('<tr>')
            
            for cell in row.findall('w:tc', self.NAMESPACES):
                cell_content = []
                
                for para in cell.findall('w:p', self.NAMESPACES):
                    para_html = self._convert_paragraph_to_html(para)
                    cell_content.append(para_html)
                
                html_parts.append(f'<td>{"".join(cell_content)}</td>')
            
            html_parts.append('</tr>')
        
        html_parts.append('</table>')
        return '\n'.join(html_parts)
    
    def _get_paragraph_style(self, para_elem: ET.Element) -> Dict[str, Any]:
        """获取段落样式"""
        style = {
            'style_id': None,
            'is_heading': False,
            'heading_level': 0,
            'alignment': 'left',
        }
        
        p_pr = para_elem.find('w:pPr', self.NAMESPACES)
        if p_pr is not None:
            p_style = p_pr.find('w:pStyle', self.NAMESPACES)
            if p_style is not None:
                style['style_id'] = p_style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                
                if style['style_id'] and style['style_id'].startswith('Heading'):
                    style['is_heading'] = True
                    try:
                        style['heading_level'] = int(style['style_id'].replace('Heading', ''))
                    except:
                        style['heading_level'] = 1
            
            jc = p_pr.find('w:jc', self.NAMESPACES)
            if jc is not None:
                style['alignment'] = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        
        return style
    
    def _parse_omml(self, omml_elem: ET.Element) -> str:
        """解析 OMML 公式"""
        try:
            return convert_omml_to_latex(omml_elem)
        except Exception as e:
            logger.debug(f"OMML 解析失败: {e}")
            return ""
    
    def _get_tag_name(self, full_tag: str) -> str:
        """从完整标签名中提取标签名"""
        if '}' in full_tag:
            return full_tag.split('}')[1]
        return full_tag
    
    def _escape_html(self, text: str) -> str:
        """转义 HTML 特殊字符"""
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        return text


def convert_docx_to_html(file_path: str) -> str:
    """便捷函数：将 docx 转换为 HTML"""
    with DocxToHtmlConverter(file_path) as converter:
        return converter.convert_to_html()
