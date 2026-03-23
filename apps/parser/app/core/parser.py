"""
Word 文档解析器
解析 docx 文件，提取文本、公式、图片等内容
支持：OMML公式、旧版公式编辑器（图片形式）、普通图片
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import re

from .omml2latex import convert_omml_to_latex
from .image_converter import ImageConverter
from .wmf_converter import WMFConverter
from .mathtype_parser import MathTypeParser


class DocxParser:
    """Word 文档解析器"""
    
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
        'w10': 'urn:schemas-microsoft-com:office:word',
    }
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.docx_zip = None
        self.relationships = {}
        self.media_map = {}
        self._has_omml = False  # 标记是否找到OMML公式
        self._has_vml_shapes = False  # 标记是否找到VML图形
        
    def __enter__(self):
        self.docx_zip = zipfile.ZipFile(self.file_path, 'r')
        self._load_relationships()
        self._detect_formula_type()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.docx_zip:
            self.docx_zip.close()
    
    def _detect_formula_type(self):
        """检测文档中的公式类型"""
        try:
            doc_xml = self.docx_zip.read('word/document.xml')
            root = ET.fromstring(doc_xml)
            
            # 检查是否有OMML公式
            omaths = root.findall('.//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath')
            self._has_omml = len(omaths) > 0
            
            # 检查是否有VML shape（可能是旧版公式编辑器）
            ns_v = '{urn:schemas-microsoft-com:vml}'
            shapes = root.findall(f'.//{ns_v}shape')
            self._has_vml_shapes = len(shapes) > 0
            
            print(f"文档分析: OMML公式={self._has_omml}, VML图形={self._has_vml_shapes}")
        except Exception as e:
            print(f"检测公式类型失败: {e}")
            
    def _load_relationships(self):
        """加载文档关系"""
        try:
            rels_content = self.docx_zip.read('word/_rels/document.xml.rels')
            root = ET.fromstring(rels_content)
            
            for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                rel_type = rel.get('Type', '').split('/')[-1]
                rel_target = rel.get('Target')
                self.relationships[rel_id] = {
                    'type': rel_type,
                    'target': rel_target
                }
        except Exception as e:
            print(f"解析关系文件失败: {e}")
            
    def extract_media(self, output_dir: str, file_id: str = None) -> Dict[str, str]:
        """提取文档中的媒体文件，自动转换 WMF 格式和 OLE 对象为 PNG"""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        if file_id is None:
            file_id = output_path.name

        self.media_map = {}

        converter = ImageConverter()
        wmf_converter = WMFConverter()
        mathtype_parser = MathTypeParser()

        # 处理 OLE 对象（MathType 公式）
        for item in self.docx_zip.namelist():
            if item.startswith('word/embeddings/') and item.endswith('.bin'):
                original_filename = Path(item).name
                new_filename = f"{file_id}_{Path(original_filename).stem}.png"
                output_file = output_path / new_filename

                try:
                    with self.docx_zip.open(item) as src:
                        ole_content = src.read()

                    temp_ole_path = output_path / f"temp_{file_id}_{original_filename}"
                    with open(temp_ole_path, 'wb') as f:
                        f.write(ole_content)

                    success, result = mathtype_parser.convert_to_image(
                        str(temp_ole_path),
                        str(output_file),
                        dpi=300
                    )

                    if success:
                        print(f"MathType OLE 转换成功: {original_filename} -> {new_filename}")
                        self.media_map[original_filename] = new_filename
                    else:
                        print(f"MathType OLE 转换失败: {result}")

                    if temp_ole_path.exists():
                        temp_ole_path.unlink()

                except Exception as e:
                    print(f"提取 OLE 对象失败 {original_filename}: {e}")

        # 处理媒体文件（WMF 和其他图片）
        for item in self.docx_zip.namelist():
            if item.startswith('word/media/'):
                original_filename = Path(item).name

                file_ext = Path(original_filename).suffix.lower()
                is_wmf = file_ext in {'.wmf', '.emf'}

                if is_wmf:
                    # 解析阶段不再强制转换 WMF，直接保留原始文件。
                    # 原因：批量转换失败率高且耗时长，最终由 API 图片路由按需转换更稳定。
                    new_filename = f"{file_id}_{original_filename}"
                else:
                    new_filename = f"{file_id}_{original_filename}"

                output_file = output_path / new_filename

                try:
                    with self.docx_zip.open(item) as src:
                        content = src.read()

                    if is_wmf:
                        with open(output_file, 'wb') as dst:
                            dst.write(content)
                    else:
                        with open(output_file, 'wb') as dst:
                            dst.write(content)

                    self.media_map[original_filename] = new_filename

                except Exception as e:
                    print(f"提取图片失败 {original_filename}: {e}")

        return self.media_map
        
    def parse_document(self) -> List[Dict[str, Any]]:
        """解析整个文档"""
        document_xml = self.docx_zip.read('word/document.xml')
        root = ET.fromstring(document_xml)
        
        paragraphs = []
        
        body = root.find('.//w:body', self.NAMESPACES)
        if body is not None:
            for para in body.findall('w:p', self.NAMESPACES):
                para_data = self._parse_paragraph(para)
                if para_data:
                    paragraphs.append(para_data)
        else:
            for para in root.findall('.//w:p', self.NAMESPACES):
                para_data = self._parse_paragraph(para)
                if para_data:
                    paragraphs.append(para_data)
                
        return paragraphs
        
    def _parse_paragraph(self, para_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析单个段落"""
        p_style = self._get_paragraph_style(para_elem)
        
        content_items = []
        full_text = ""
        
        # 递归解析段落中的所有内容
        items, text = self._parse_element_recursive(para_elem)
        content_items.extend(items)
        full_text += text
                    
        if not content_items:
            return None

        latex_formulas: List[str] = []
        inline_formula_count = 0
        block_formula_count = 0
        has_formula = False

        for item in content_items:
            item_type = item.get('type')
            if item_type in ('latex', 'latex_block'):
                has_formula = True
                latex_content = item.get('content')
                if isinstance(latex_content, str):
                    latex_formulas.append(latex_content)
                if item_type == 'latex':
                    inline_formula_count += 1
                else:
                    block_formula_count += 1
        
        return {
            'style': p_style,
            'text': full_text.strip(),
            'content_items': content_items,
            'has_formula': has_formula,
            'inline_formula_count': inline_formula_count,
            'block_formula_count': block_formula_count,
            'latex_formulas': latex_formulas,
        }
    
    def _parse_element_recursive(self, elem: ET.Element) -> Tuple[List[Dict], str]:
        """递归解析元素及其子元素"""
        items = []
        text = ""
        
        tag = self._get_tag_name(elem.tag)
        
        # 处理特定标签
        if tag == 'oMath':
            latex = self._parse_omml(elem)
            items.append({'type': 'latex', 'content': latex})
            text += f"${latex}$"
            return items, text
            
        elif tag == 'oMathPara':
            latex = self._parse_omml(elem)
            items.append({'type': 'latex_block', 'content': latex})
            text += f"$${latex}$$"
            return items, text
            
        elif tag == 't':
            content = elem.text or ''
            if content:
                items.append({'type': 'text', 'content': content})
                text += content
            return items, text
            
        elif tag == 'tab':
            items.append({'type': 'text', 'content': '\t'})
            text += '\t'
            return items, text
            
        elif tag == 'br':
            items.append({'type': 'text', 'content': '\n'})
            text += '\n'
            return items, text
            
        elif tag == 'drawing':
            image_info = self._parse_drawing(elem)
            if image_info:
                items.append({'type': 'image', 'content': image_info})
            return items, text
            
        elif tag == 'pict':
            image_info = self._parse_vml(elem)
            if image_info:
                items.append({'type': 'image', 'content': image_info})
            return items, text
            
        elif tag == 'object':
            image_info = self._parse_ole_object(elem)
            if image_info:
                items.append({'type': 'image', 'content': image_info})
            return items, text
            
        elif tag == 'shape':
            image_info = self._parse_vml_shape(elem)
            if image_info:
                items.append({'type': 'image', 'content': image_info})
            return items, text
            
        elif tag == 'r':
            # Run 元素
            for child in elem:
                child_items, child_text = self._parse_element_recursive(child)
                items.extend(child_items)
                text += child_text
            return items, text
        
        # 对于其他元素，递归处理子元素
        for child in elem:
            child_items, child_text = self._parse_element_recursive(child)
            items.extend(child_items)
            text += child_text
            
        return items, text
    
    def _parse_formula_shape(self, pict_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析公式编辑器创建的 shape（在 pict 内）- 现在统一作为图片处理"""
        try:
            ns_v = '{urn:schemas-microsoft-com:vml}'
            ns_o = '{urn:schemas-microsoft-com:office:office}'
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            
            shape = None
            for child in pict_elem:
                tag = self._get_tag_name(child.tag)
                if tag == 'shape':
                    shape = child
                    break
                    
            if shape is None:
                shape = pict_elem.find(f'.//{ns_v}shape')
                
            if shape is None:
                return None
                
            imagedata = None
            for child in shape:
                tag = self._get_tag_name(child.tag)
                if tag == 'imagedata':
                    imagedata = child
                    break
                    
            if imagedata is None:
                imagedata = shape.find(f'{ns_v}imagedata')
                
            if imagedata is not None:
                rel_id = imagedata.get('id')
                if not rel_id:
                    rel_id = imagedata.get(f'{ns_o}relid')
                if not rel_id:
                    rel_id = imagedata.get(f'{ns_r}id')
                    
                if rel_id and rel_id in self.relationships:
                    rel_info = self.relationships[rel_id]
                    target = rel_info['target']
                    original_filename = Path(target).name
                    full_filename = self.media_map.get(original_filename, original_filename)
                    return {
                        'type': 'image',
                        'rel_id': rel_id,
                        'filename': full_filename,
                        'original_filename': original_filename,
                    }
                    
        except Exception as e:
            print(f"解析 shape 失败: {e}")
        return None
        
    def _parse_vml_shape(self, shape_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析独立的 VML shape - 现在统一作为图片处理"""
        try:
            ns_v = '{urn:schemas-microsoft-com:vml}'
            ns_o = '{urn:schemas-microsoft-com:office:office}'
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            
            imagedata = shape_elem.find(f'{ns_v}imagedata')
            if imagedata is not None:
                rel_id = imagedata.get(f'{ns_o}relid') or imagedata.get(f'{ns_r}id')
                if rel_id and rel_id in self.relationships:
                    rel_info = self.relationships[rel_id]
                    target = rel_info['target']
                    original_filename = Path(target).name
                    full_filename = self.media_map.get(original_filename, original_filename)
                    return {
                        'type': 'image',
                        'rel_id': rel_id,
                        'filename': full_filename,
                        'original_filename': original_filename,
                    }
                    
        except Exception as e:
            print(f"解析 VML shape 失败: {e}")
        return None
        
    def _parse_ole_object(self, obj_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析 OLE 对象（MathType 公式）"""
        try:
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            
            ole = None
            for child in obj_elem:
                tag = self._get_tag_name(child.tag)
                if tag == 'OLEObject':
                    ole = child
                    break
                    
            if ole is None:
                ole = obj_elem.find(f'.//{ns_r}OLEObject')
                
            if ole is not None:
                r_id = ole.get(f'{{{self.NAMESPACES["r"]}}}id')
                if r_id and r_id in self.relationships:
                    rel_info = self.relationships[r_id]
                    target = rel_info.get('target', '')
                    
                    if 'embeddings' in target:
                        ole_filename = Path(target).name
                        if ole_filename in self.media_map:
                            return {
                                'filename': self.media_map[ole_filename],
                                'type': 'mathtype_ole'
                            }
                    
                shape_info = self._parse_formula_shape(obj_elem)
                if shape_info:
                    return shape_info
                    
        except Exception as e:
            print(f"解析 OLE 对象失败: {e}")
        return None
        
    def _get_paragraph_style(self, para_elem: ET.Element) -> Dict[str, Any]:
        """获取段落样式信息"""
        style = {
            'style_id': None,
            'is_heading': False,
            'heading_level': 0,
            'alignment': 'left',
            'indent': 0,
        }
        
        pPr = para_elem.find('w:pPr', self.NAMESPACES)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', self.NAMESPACES)
            if pStyle is not None:
                style_id = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                style['style_id'] = style_id
                
                if style_id:
                    if style_id.startswith('Heading'):
                        style['is_heading'] = True
                        try:
                            style['heading_level'] = int(style_id.replace('Heading', ''))
                        except ValueError:
                            pass
                    elif style_id.startswith('标题'):
                        style['is_heading'] = True
                        try:
                            style['heading_level'] = int(style_id.replace('标题', ''))
                        except ValueError:
                            pass
                        
            jc = pPr.find('w:jc', self.NAMESPACES)
            if jc is not None:
                align = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                style['alignment'] = align or 'left'
                
        return style
        
    def _parse_omml(self, omath_elem: ET.Element) -> str:
        """解析 OMML 公式为 LaTeX"""
        try:
            omml_xml = ET.tostring(omath_elem, encoding='unicode')
            return convert_omml_to_latex(omml_xml)
        except Exception as e:
            print(f"OMML 解析失败: {e}")
            return "[公式]"
        
    def _parse_drawing(self, drawing_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析 DrawingML 图片"""
        try:
            ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            
            blip = drawing_elem.find(f'.//{ns_a}blip')
            if blip is not None:
                embed = blip.get(f'{ns_r}embed')
                if embed and embed in self.relationships:
                    rel_info = self.relationships[embed]
                    target = rel_info['target']
                    original_filename = Path(target).name
                    # 获取完整的文件名（包含 file_id 前缀）
                    full_filename = self.media_map.get(original_filename, original_filename)
                    
                    return {
                        'type': 'image',
                        'rel_id': embed,
                        'filename': full_filename,
                        'original_filename': original_filename,
                    }
                        
        except Exception as e:
            print(f"解析图片出错: {e}")
            
        return None
        
    def _parse_vml(self, pict_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """解析 VML 图形"""
        try:
            ns_v = '{urn:schemas-microsoft-com:vml}'
            ns_o = '{urn:schemas-microsoft-com:office:office}'
            ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            
            imagedata = pict_elem.find(f'.//{ns_v}imagedata')
            if imagedata is not None:
                rel_id = imagedata.get(f'{ns_r}id')
                if rel_id and rel_id in self.relationships:
                    rel_info = self.relationships[rel_id]
                    target = rel_info['target']
                    original_filename = Path(target).name
                    # 获取完整的文件名（包含 file_id 前缀）
                    full_filename = self.media_map.get(original_filename, original_filename)
                    
                    return {
                        'type': 'image',
                        'rel_id': rel_id,
                        'filename': full_filename,
                        'original_filename': original_filename,
                    }
                    
        except Exception as e:
            print(f"解析 VML 失败: {e}")
        return None
        
    def _get_tag_name(self, tag: str) -> str:
        """获取标签名（去掉命名空间）"""
        if '}' in tag:
            return tag.split('}')[1]
        return tag


def parse_docx(file_path: str, extract_images: bool = True, image_output_dir: str = None, file_id: str = None) -> Dict[str, Any]:
    """便捷函数：解析 Word 文档"""
    result = {
        'paragraphs': [],
        'images': {},
        'metadata': {}
    }
    
    with DocxParser(file_path) as parser:
        if extract_images:
            if image_output_dir is None:
                image_output_dir = Path(file_path).parent / 'images'
            # 如果没有提供 file_id，从 output_dir 中提取
            if file_id is None:
                file_id = Path(image_output_dir).name
            result['images'] = parser.extract_media(image_output_dir, file_id)
            parser.media_map = result['images']
            
        paragraphs = parser.parse_document()
        result['paragraphs'] = paragraphs

        # 文档级元数据：为复杂数学试卷提供整体统计信息
        total_formulas = 0
        total_formula_paragraphs = 0
        for para in paragraphs:
            if isinstance(para, dict):
                latex_list = para.get('latex_formulas') or []
                if latex_list:
                    total_formulas += len(latex_list)
                    total_formula_paragraphs += 1

        result['metadata'] = {
            'has_omml': parser._has_omml,
            'has_vml_shapes': parser._has_vml_shapes,
            'total_paragraphs': len(paragraphs),
            'total_formula_paragraphs': total_formula_paragraphs,
            'total_formulas': total_formulas,
        }

    return result
