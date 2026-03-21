"""
WMF 图片转换器
使用多种方法转换 WMF 文件为 PNG 格式
"""

import subprocess
import shutil
from pathlib import Path
from typing import Optional, Tuple
import logging
import os

logger = logging.getLogger(__name__)


class WMFConverter:
    """WMF 文件转换器"""
    
    def __init__(self):
        self.converter = self._detect_converter()
    
    def _detect_converter(self) -> Optional[str]:
        """检测可用的转换工具"""
        converters = [
            ('inkscape', self._convert_with_inkscape),
            ('libreoffice', self._convert_with_libreoffice),
            ('soffice', self._convert_with_libreoffice),
            ('magick', self._convert_with_imagemagick),
            ('convert', self._convert_with_imagemagick),
        ]
        
        for name, func in converters:
            if shutil.which(name):
                logger.info(f"检测到 WMF 转换工具: {name}")
                return name
        
        logger.warning("未检测到可用的 WMF 转换工具")
        return None
    
    def convert(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        转换 WMF 文件为 PNG
        
        Args:
            input_path: 输入 WMF 文件路径
            output_path: 输出 PNG 文件路径
            
        Returns:
            (成功标志, 输出文件路径或错误信息)
        """
        input_path = Path(input_path)
        output_path = Path(output_path)
        
        if not input_path.exists():
            return False, f"输入文件不存在: {input_path}"
        
        if not self.converter:
            return False, "没有可用的 WMF 转换工具"
        
        converters = {
            'inkscape': self._convert_with_inkscape,
            'libreoffice': self._convert_with_libreoffice,
            'soffice': self._convert_with_libreoffice,
            'magick': self._convert_with_imagemagick,
            'convert': self._convert_with_imagemagick,
        }
        
        converter_func = converters.get(self.converter)
        if converter_func:
            return converter_func(str(input_path), str(output_path))
        
        return False, "未知的转换工具"
    
    def _convert_with_inkscape(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 Inkscape 转换"""
        try:
            cmd = [
                'inkscape',
                input_path,
                '--export-filename', output_path,
                '--export-type', 'png',
                '--export-dpi', '300'
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and Path(output_path).exists():
                logger.info(f"Inkscape 转换成功 (300 DPI): {Path(input_path).name}")
                return True, output_path
            else:
                error = result.stderr or result.stdout
                logger.error(f"Inkscape 转换失败: {error}")
                return False, f"Inkscape 转换失败: {error}"
                
        except subprocess.TimeoutExpired:
            return False, "Inkscape 转换超时"
        except Exception as e:
            logger.error(f"Inkscape 转换异常: {e}")
            return False, f"Inkscape 转换异常: {e}"
    
    def _convert_with_libreoffice(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 LibreOffice 转换"""
        try:
            output_dir = str(Path(output_path).parent)
            
            cmd = [
                self.converter,
                '--headless',
                '--convert-to', 'png',
                '--outdir', output_dir,
                input_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            expected_output = Path(output_dir) / f"{Path(input_path).stem}.png"
            
            if result.returncode == 0 and expected_output.exists():
                if str(expected_output) != output_path:
                    expected_output.rename(output_path)
                logger.info(f"LibreOffice 转换成功: {Path(input_path).name}")
                return True, output_path
            else:
                error = result.stderr or result.stdout
                logger.error(f"LibreOffice 转换失败: {error}")
                return False, f"LibreOffice 转换失败: {error}"
                
        except subprocess.TimeoutExpired:
            return False, "LibreOffice 转换超时"
        except Exception as e:
            logger.error(f"LibreOffice 转换异常: {e}")
            return False, f"LibreOffice 转换异常: {e}"
    
    def _convert_with_imagemagick(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 ImageMagick 转换"""
        try:
            cmd_name = 'magick' if self.converter == 'magick' else 'convert'
            cmd = [
                cmd_name,
                '-density', '300',
                input_path,
                '-quality', '90',
                output_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and Path(output_path).exists():
                logger.info(f"ImageMagick 转换成功: {Path(input_path).name}")
                return True, output_path
            else:
                error = result.stderr or result.stdout
                if 'libreoffice' in error.lower():
                    return False, "ImageMagick 需要 LibreOffice 支持"
                logger.error(f"ImageMagick 转换失败: {error}")
                return False, f"ImageMagick 转换失败: {error}"
                
        except subprocess.TimeoutExpired:
            return False, "ImageMagick 转换超时"
        except Exception as e:
            logger.error(f"ImageMagick 转换异常: {e}")
            return False, f"ImageMagick 转换异常: {e}"


def convert_wmf_to_png(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    便捷函数：转换 WMF 文件为 PNG
    
    Args:
        input_path: 输入 WMF 文件路径
        output_path: 输出 PNG 文件路径
        
    Returns:
        (成功标志, 输出文件路径或错误信息)
    """
    converter = WMFConverter()
    return converter.convert(input_path, output_path)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    
    converter = WMFConverter()
    print(f"检测到的转换工具: {converter.converter}")
