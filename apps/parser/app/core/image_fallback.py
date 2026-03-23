"""
图片降级处理器
当 Pandoc 无法处理某些图片格式时，使用此模块进行降级处理
支持：WMF/EMF 转换、MathType OLE 提取、图片格式转换
"""

import zipfile
import subprocess
import shutil
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class ImageFallbackProcessor:
    """图片降级处理器"""
    
    def __init__(self):
        self.wmf_converter = self._detect_wmf_converter()
        self.ole_tools = self._detect_ole_tools()
        
    def _detect_wmf_converter(self) -> Optional[str]:
        """检测可用的 WMF 转换工具"""
        converters = ['inkscape', 'libreoffice', 'soffice', 'magick', 'convert']
        for converter in converters:
            if shutil.which(converter):
                logger.info(f"检测到 WMF 转换工具: {converter}")
                return converter
        return None
    
    def _detect_ole_tools(self) -> bool:
        """检测 OLE 处理工具"""
        try:
            import olefile
            logger.info("olefile 可用")
            return True
        except ImportError:
            logger.warning("olefile 未安装，OLE 对象处理受限")
            return False
    
    def process_docx_images(
        self,
        docx_path: str,
        output_dir: str,
        file_id: str = None
    ) -> Tuple[bool, Dict[str, str], str]:
        """
        处理 DOCX 中的所有图片（包括嵌入对象）
        
        Args:
            docx_path: DOCX 文件路径
            output_dir: 输出目录
            file_id: 文件 ID（用于命名）
            
        Returns:
            (成功标志, 图片映射字典, 错误信息)
        """
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        if file_id is None:
            file_id = output_path.name
        
        media_map = {}
        errors = []
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                media_map.update(self._process_media_files(docx_zip, output_path, file_id))
                
                ole_map, ole_errors = self._process_ole_objects(docx_zip, output_path, file_id)
                media_map.update(ole_map)
                errors.extend(ole_errors)
            
            logger.info(f"图片处理完成，共 {len(media_map)} 个文件")
            return True, media_map, "; ".join(errors) if errors else ""
            
        except Exception as e:
            logger.error(f"处理 DOCX 图片失败: {e}")
            return False, {}, str(e)
    
    def _process_media_files(
        self,
        docx_zip: zipfile.ZipFile,
        output_path: Path,
        file_id: str
    ) -> Dict[str, str]:
        """处理 word/media/ 目录下的文件"""
        media_map = {}
        
        for item in docx_zip.namelist():
            if not item.startswith('word/media/'):
                continue
            
            original_filename = Path(item).name
            file_ext = Path(original_filename).suffix.lower()
            
            if file_ext in {'.wmf', '.emf'}:
                success, new_name = self._convert_wmf_file(
                    docx_zip, item, output_path, file_id, original_filename
                )
                if success:
                    media_map[original_filename] = new_name
                else:
                    wmf_output = output_path / f"{file_id}_{original_filename}"
                    with open(wmf_output, 'wb') as f:
                        f.write(docx_zip.read(item))
                    media_map[original_filename] = f"{file_id}_{original_filename}"
            else:
                new_filename = f"{file_id}_{original_filename}"
                output_file = output_path / new_filename
                
                with open(output_file, 'wb') as f:
                    f.write(docx_zip.read(item))
                
                media_map[original_filename] = new_filename
        
        return media_map
    
    def _convert_wmf_file(
        self,
        docx_zip: zipfile.ZipFile,
        item: str,
        output_path: Path,
        file_id: str,
        original_filename: str
    ) -> Tuple[bool, str]:
        """转换单个 WMF/EMF 文件"""
        if not self.wmf_converter:
            return False, ""
        
        try:
            temp_wmf = output_path / f"temp_{original_filename}"
            with open(temp_wmf, 'wb') as f:
                f.write(docx_zip.read(item))
            
            new_filename = f"{file_id}_{Path(original_filename).stem}.png"
            output_file = output_path / new_filename
            
            success = self._convert_with_tool(str(temp_wmf), str(output_file))
            
            if temp_wmf.exists():
                temp_wmf.unlink()
            
            return success, new_filename if success else ""
            
        except Exception as e:
            logger.error(f"WMF 转换失败 {original_filename}: {e}")
            return False, ""
    
    def _convert_with_tool(self, input_path: str, output_path: str) -> bool:
        """使用检测到的工具进行转换"""
        converter = self.wmf_converter
        
        try:
            if converter == 'inkscape':
                cmd = ['inkscape', input_path, '--export-filename', output_path, '--export-dpi', '300']
            elif converter in ['libreoffice', 'soffice']:
                output_dir = str(Path(output_path).parent)
                cmd = [converter, '--headless', '--convert-to', 'png', '--outdir', output_dir, input_path]
            elif converter in ['magick', 'convert']:
                cmd = [converter, '-density', '300', input_path, '-quality', '90', output_path]
            else:
                return False
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and Path(output_path).exists():
                logger.info(f"{converter} 转换成功: {Path(input_path).name}")
                
                if converter in ['libreoffice', 'soffice']:
                    expected = Path(output_dir) / f"{Path(input_path).stem}.png"
                    if expected.exists() and str(expected) != output_path:
                        expected.rename(output_path)
                
                return True
            else:
                logger.error(f"{converter} 转换失败: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"{converter} 转换超时")
            return False
        except Exception as e:
            logger.error(f"{converter} 转换异常: {e}")
            return False
    
    def _process_ole_objects(
        self,
        docx_zip: zipfile.ZipFile,
        output_path: Path,
        file_id: str
    ) -> Tuple[Dict[str, str], List[str]]:
        """处理 OLE 嵌入对象（MathType 等）"""
        ole_map = {}
        errors = []
        
        for item in docx_zip.namelist():
            if not item.startswith('word/embeddings/') or not item.endswith('.bin'):
                continue
            
            original_filename = Path(item).name
            
            try:
                success, result = self._process_ole_file(
                    docx_zip, item, output_path, file_id, original_filename
                )
                
                if success:
                    ole_map[original_filename] = result
                else:
                    errors.append(f"OLE 处理失败 {original_filename}: {result}")
                    
            except Exception as e:
                errors.append(f"OLE 处理异常 {original_filename}: {str(e)}")
        
        return ole_map, errors
    
    def _process_ole_file(
        self,
        docx_zip: zipfile.ZipFile,
        item: str,
        output_path: Path,
        file_id: str,
        original_filename: str
    ) -> Tuple[bool, str]:
        """处理单个 OLE 文件"""
        try:
            import olefile
            
            temp_ole = output_path / f"temp_{original_filename}"
            with open(temp_ole, 'wb') as f:
                f.write(docx_zip.read(item))
            
            ole = olefile.OleFileIO(str(temp_ole))
            
            if ole.exists('Equation Native'):
                data = ole.openstream('Equation Native').read()
                
                new_filename = f"{file_id}_{Path(original_filename).stem}.png"
                output_file = output_path / new_filename
                
                success, msg = self._convert_mathtype_ole(data, str(output_file))
                
                ole.close()
                if temp_ole.exists():
                    temp_ole.unlink()
                
                return success, new_filename if success else msg
            else:
                ole.close()
                if temp_ole.exists():
                    temp_ole.unlink()
                return False, "非 MathType OLE 对象"
                
        except ImportError:
            return False, "olefile 未安装"
        except Exception as e:
            return False, str(e)
    
    def _convert_mathtype_ole(self, data: bytes, output_path: str) -> Tuple[bool, str]:
        """转换 MathType OLE 数据为图片"""
        try:
            if shutil.which('soffice') or shutil.which('libreoffice'):
                return self._convert_ole_with_libreoffice(data, output_path)
            
            preview = self._extract_wmf_preview(data)
            if preview:
                temp_wmf = Path(output_path).parent / "temp_preview.wmf"
                with open(temp_wmf, 'wb') as f:
                    f.write(preview)
                
                success = self._convert_with_tool(str(temp_wmf), output_path)
                if temp_wmf.exists():
                    temp_wmf.unlink()
                
                return success, "转换成功" if success else "WMF 转换失败"
            
            return False, "无法提取预览图"
            
        except Exception as e:
            return False, str(e)
    
    def _convert_ole_with_libreoffice(self, data: bytes, output_path: str) -> Tuple[bool, str]:
        """使用 LibreOffice 转换 OLE 对象"""
        try:
            temp_dir = tempfile.mkdtemp()
            temp_ole = Path(temp_dir) / "equation.bin"
            
            with open(temp_ole, 'wb') as f:
                f.write(data)
            
            libreoffice = shutil.which('soffice') or shutil.which('libreoffice')
            cmd = [
                libreoffice,
                '--headless',
                '--convert-to', 'png',
                '--outdir', temp_dir,
                str(temp_ole)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            expected_output = Path(temp_dir) / "equation.png"
            
            if result.returncode == 0 and expected_output.exists():
                shutil.move(str(expected_output), output_path)
                shutil.rmtree(temp_dir)
                return True, "转换成功"
            else:
                shutil.rmtree(temp_dir)
                return False, result.stderr or "LibreOffice 转换失败"
                
        except Exception as e:
            return False, str(e)
    
    def _extract_wmf_preview(self, ole_data: bytes) -> Optional[bytes]:
        """从 OLE 数据中提取 WMF 预览图"""
        try:
            wmf_signature = b'\xd7\xcd\xc6\x9a'
            pos = ole_data.find(wmf_signature)
            
            if pos > 0:
                return ole_data[pos:]
            
            return None
            
        except Exception as e:
            logger.error(f"提取 WMF 预览失败: {e}")
            return None


def process_docx_images(
    docx_path: str,
    output_dir: str,
    file_id: str = None
) -> Tuple[bool, Dict[str, str], str]:
    """
    处理 DOCX 图片的便捷函数
    
    Args:
        docx_path: DOCX 文件路径
        output_dir: 输出目录
        file_id: 文件 ID
        
    Returns:
        (成功标志, 图片映射, 错误信息)
    """
    processor = ImageFallbackProcessor()
    return processor.process_docx_images(docx_path, output_dir, file_id)
