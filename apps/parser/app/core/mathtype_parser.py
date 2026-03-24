"""
MathType OLE 对象解析器
用于从 Word 文档中提取 MathType 公式并转换为高质量图片
"""
import olefile
import struct
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Tuple, Optional, Dict, Any
import logging

logger = logging.getLogger(__name__)


class MathTypeParser:
    """MathType 公式解析器"""
    
    MTEF_SIGNATURE = b'DSMT'
    WMF_SIGNATURE = b'\xd7\xcd\xc6\x9a'
    EMF_SIGNATURE = b'\x20\x45\x4d\x46'
    
    def __init__(self):
        self.temp_dir = None
    
    def parse_ole_file(self, ole_path: str) -> Optional[Dict[str, Any]]:
        """
        解析 MathType OLE 文件
        
        Args:
            ole_path: OLE 文件路径
            
        Returns:
            包含公式信息的字典，或 None 如果解析失败
        """
        try:
            ole = olefile.OleFileIO(ole_path)
            
            if not ole.exists('Equation Native'):
                logger.warning(f"OLE 文件中没有 Equation Native 流: {ole_path}")
                return None
            
            data = ole.openstream('Equation Native').read()
            ole.close()
            
            return self._parse_mtef(data)
            
        except Exception as e:
            logger.error(f"解析 OLE 文件失败: {e}")
            return None
    
    def _parse_mtef(self, data: bytes) -> Dict[str, Any]:
        """
        解析 MTEF (MathType Equation Format) 数据
        
        Args:
            data: MTEF 原始数据
            
        Returns:
            解析后的公式信息
        """
        result = {
            'mtef_version': None,
            'fonts': [],
            'equation_data': None,
            'raw_data': data
        }
        
        try:
            # MTEF 头部结构
            # Offset 0-3: MTEF 版本和选项
            # Offset 4-27: 字体表信息
            
            if len(data) < 28:
                logger.warning("MTEF 数据太短")
                return result
            
            # 解析头部
            header_size = struct.unpack('<I', data[0:4])[0]
            result['mtef_version'] = data[4]
            
            # 查找字体表
            offset = 28
            fonts = []
            
            # 跳过头部，查找字体信息
            # MTEF 格式中，字体信息通常在特定位置
            font_section_start = data.find(b'Times New Roman')
            if font_section_start > 0:
                # 提取字体名称
                font_names = []
                i = font_section_start
                while i < len(data) - 10:
                    # 查找以 null 结尾的字符串
                    end = data.find(b'\x00', i)
                    if end > i and end - i < 50:  # 合理的字体名称长度
                        font_name = data[i:end].decode('latin-1', errors='ignore')
                        if font_name and len(font_name) > 2:
                            font_names.append(font_name)
                        i = end + 1
                    else:
                        i += 1
                
                result['fonts'] = font_names[:5]  # 最多保留 5 个字体
            
            # 保存公式数据（用于后续转换）
            result['equation_data'] = data
            
            return result
            
        except Exception as e:
            logger.error(f"解析 MTEF 数据失败: {e}")
            return result
    
    def convert_to_image(self, ole_path: str, output_path: str, dpi: int = 300) -> Tuple[bool, str]:
        """
        将 MathType OLE 对象转换为高质量图片
        
        Args:
            ole_path: OLE 文件路径
            output_path: 输出图片路径
            dpi: 输出 DPI（默认 300）
            
        Returns:
            (成功标志, 输出路径或错误信息)
        """
        try:
            # 方法 1: 使用 LibreOffice 转换（如果可用）
            if shutil.which('soffice') or shutil.which('libreoffice'):
                return self._convert_with_libreoffice(ole_path, output_path, dpi)
            
            # 方法 2: 提取 WMF 预览图并转换（质量较低）
            return self._extract_wmf_preview(ole_path, output_path, dpi)
            
        except Exception as e:
            logger.error(f"转换 MathType 失败: {e}")
            return False, str(e)
    
    def _convert_with_libreoffice(self, ole_path: str, output_path: str, dpi: int) -> Tuple[bool, str]:
        """使用 LibreOffice 转换 MathType 对象"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            
            # LibreOffice 可以直接打开 OLE 对象
            cmd = [
                shutil.which('soffice') or shutil.which('libreoffice'),
                '--headless',
                '--convert-to', 'png',
                '--outdir', temp_dir,
                ole_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                # 查找生成的 PNG 文件
                png_files = list(Path(temp_dir).glob('*.png'))
                if png_files:
                    shutil.move(str(png_files[0]), output_path)
                    shutil.rmtree(temp_dir)
                    logger.info(f"LibreOffice 转换成功: {Path(ole_path).name}")
                    return True, output_path
            
            shutil.rmtree(temp_dir)
            return False, "LibreOffice 转换失败"
            
        except Exception as e:
            return False, f"LibreOffice 转换异常: {e}"
    
    def _extract_wmf_preview(self, ole_path: str, output_path: str, dpi: int) -> Tuple[bool, str]:
        """从 OLE 对象中提取 WMF 预览图"""
        try:
            ole = olefile.OleFileIO(ole_path)
            
            # 查找 WMF/EMF 数据
            metafile_data = None
            metafile_ext = '.wmf'
            for stream in ole.listdir():
                stream_name = '/'.join(stream)
                try:
                    stream_data = ole.openstream(stream).read()
                except Exception:
                    continue
                if 'WMF' in stream_name.upper():
                    metafile_data = stream_data
                    metafile_ext = '.wmf'
                    break
                if 'EMF' in stream_name.upper():
                    metafile_data = stream_data
                    metafile_ext = '.emf'
                    break
                if 'PICTURE' in stream_name.upper() or 'OLEPRES' in stream_name.upper():
                    found_data, found_ext = self._extract_metafile_from_bytes(stream_data)
                    if found_data:
                        metafile_data = found_data
                        metafile_ext = found_ext
                        break

            if not metafile_data:
                for stream in ole.listdir():
                    try:
                        stream_data = ole.openstream(stream).read()
                    except Exception:
                        continue
                    found_data, found_ext = self._extract_metafile_from_bytes(stream_data)
                    if found_data:
                        metafile_data = found_data
                        metafile_ext = found_ext
                        break
            
            ole.close()
            
            if not metafile_data:
                return False, "OLE 对象中没有找到 WMF/EMF 预览图"
            
            # 保存 WMF/EMF 文件
            temp_meta = output_path.replace('.png', metafile_ext)
            with open(temp_meta, 'wb') as f:
                f.write(metafile_data)
            
            # 使用 Inkscape 转换
            if shutil.which('inkscape'):
                cmd = [
                    'inkscape',
                    temp_meta,
                    '--export-filename', output_path,
                    '--export-type', 'png',
                    '--export-dpi', str(dpi)
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                
                Path(temp_meta).unlink()
                
                if result.returncode == 0 and Path(output_path).exists():
                    logger.info(f"{metafile_ext.upper()} 预览图转换成功: {Path(ole_path).name}")
                    return True, output_path
            
            Path(temp_meta).unlink()
            return False, "WMF/EMF 转换失败"
            
        except Exception as e:
            return False, f"提取 WMF 预览图失败: {e}"

    def _extract_metafile_from_bytes(self, data: bytes) -> Tuple[Optional[bytes], str]:
        """从二进制流里定位 WMF/EMF 数据"""
        if not data:
            return None, '.wmf'
        wmf_pos = data.find(self.WMF_SIGNATURE)
        if wmf_pos >= 0:
            return data[wmf_pos:], '.wmf'
        emf_pos = data.find(self.EMF_SIGNATURE)
        if emf_pos >= 0:
            return data[emf_pos:], '.emf'
        return None, '.wmf'
    
    def extract_latex(self, ole_path: str) -> Optional[str]:
        """
        从 MathType 对象中提取 LaTeX 代码（如果可能）
        
        Args:
            ole_path: OLE 文件路径
            
        Returns:
            LaTeX 字符串，或 None 如果无法提取
        """
        try:
            ole = olefile.OleFileIO(ole_path)
            
            if not ole.exists('Equation Native'):
                return None
            
            data = ole.openstream('Equation Native').read()
            ole.close()
            
            # TODO: 实现 MTEF 到 LaTeX 的转换
            # 这需要完整的 MTEF 解析器
            
            return None
            
        except Exception as e:
            logger.error(f"提取 LaTeX 失败: {e}")
            return None
