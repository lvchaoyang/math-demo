"""
图片格式转换器 - 增强版
支持多种方式转换 WMF/EMF 格式
"""

import subprocess
import shutil
from pathlib import Path
from typing import Optional, Tuple
import os
import logging

logger = logging.getLogger(__name__)


class ImageConverter:
    """图片格式转换器"""

    SUPPORTED_INPUT_FORMATS = {'.wmf', '.emf', '.svg', '.eps'}
    SUPPORTED_OUTPUT_FORMATS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}

    def __init__(self):
        self.converter_tool = self._detect_converter()

    def _detect_converter(self) -> Optional[str]:
        """检测可用的图片转换工具"""
        tools = ['inkscape', 'magick', 'convert', 'ffmpeg']

        for tool in tools:
            if shutil.which(tool):
                logger.info(f"检测到图片转换工具: {tool}")
                return tool

        logger.warning("未检测到可用的图片转换工具")
        return None

    def convert(self, input_path: str, output_path: Optional[str] = None,
                output_format: str = '.png') -> Tuple[bool, str]:
        """
        转换图片格式

        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径（可选，默认与输入相同目录）
            output_format: 输出格式（默认 .png）

        Returns:
            (成功标志, 输出文件路径或错误信息)
        """
        input_path = Path(input_path)

        if not input_path.exists():
            return False, f"输入文件不存在: {input_path}"

        input_ext = input_path.suffix.lower()
        
        if input_ext not in self.SUPPORTED_INPUT_FORMATS:
            if input_ext in self.SUPPORTED_OUTPUT_FORMATS:
                return True, str(input_path)
            return False, f"不支持的输入格式: {input_ext}"

        if output_path is None:
            output_path = input_path.with_suffix(output_format)
        else:
            output_path = Path(output_path)

        if self.converter_tool == 'inkscape':
            return self._convert_with_inkscape(input_path, output_path)
        elif self.converter_tool == 'magick':
            return self._convert_with_magick(input_path, output_path)
        elif self.converter_tool == 'convert':
            return self._convert_with_imagemagick(input_path, output_path)
        elif self.converter_tool == 'ffmpeg':
            return self._convert_with_ffmpeg(input_path, output_path)
        else:
            return False, "没有可用的图片转换工具（需要 Inkscape、ImageMagick 或 FFmpeg）"

    def _convert_with_inkscape(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """使用 Inkscape 转换"""
        try:
            cmd = [
                'inkscape',
                str(input_path),
                '--export-filename', str(output_path),
                '--export-type', output_path.suffix.lstrip('.')
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            if result.returncode == 0 and output_path.exists():
                logger.info(f"Inkscape 转换成功: {input_path.name} -> {output_path.name}")
                return True, str(output_path)
            else:
                error_msg = result.stderr or result.stdout
                logger.error(f"Inkscape 转换失败: {error_msg}")
                return False, f"Inkscape 转换失败: {error_msg}"

        except subprocess.TimeoutExpired:
            logger.error("Inkscape 转换超时")
            return False, "Inkscape 转换超时"
        except Exception as e:
            logger.error(f"Inkscape 转换异常: {str(e)}")
            return False, f"Inkscape 转换异常: {str(e)}"
    
    def _convert_with_magick(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """使用 ImageMagick v7 (magick) 转换"""
        try:
            cmd = [
                'magick',
                '-density', '300',
                str(input_path),
                '-quality', '90',
                str(output_path)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            if result.returncode == 0 and output_path.exists():
                logger.info(f"ImageMagick v7 转换成功: {input_path.name} -> {output_path.name}")
                return True, str(output_path)
            else:
                error_msg = result.stderr or result.stdout
                logger.error(f"ImageMagick v7 转换失败: {error_msg}")
                return False, f"ImageMagick v7 转换失败: {error_msg}"

        except subprocess.TimeoutExpired:
            logger.error("ImageMagick v7 转换超时")
            return False, "ImageMagick v7 转换超时"
        except Exception as e:
            logger.error(f"ImageMagick v7 转换异常: {str(e)}")
            return False, f"ImageMagick v7 转换异常: {str(e)}"

    def _convert_with_imagemagick(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """使用 ImageMagick v6 (convert) 转换"""
        try:
            cmd = [
                'convert',
                '-density', '300',
                str(input_path),
                '-quality', '90',
                str(output_path)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            if result.returncode == 0 and output_path.exists():
                logger.info(f"ImageMagick v6 转换成功: {input_path.name} -> {output_path.name}")
                return True, str(output_path)
            else:
                error_msg = result.stderr or result.stdout
                logger.error(f"ImageMagick v6 转换失败: {error_msg}")
                return False, f"ImageMagick v6 转换失败: {error_msg}"

        except subprocess.TimeoutExpired:
            logger.error("ImageMagick v6 转换超时")
            return False, "ImageMagick v6 转换超时"
        except Exception as e:
            logger.error(f"ImageMagick v6 转换异常: {str(e)}")
            return False, f"ImageMagick v6 转换异常: {str(e)}"

    def _convert_with_ffmpeg(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """使用 FFmpeg 转换（作为最后的备选）"""
        try:
            cmd = [
                'ffmpeg',
                '-i', str(input_path),
                '-y',
                str(output_path)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            if result.returncode == 0 and output_path.exists():
                logger.info(f"FFmpeg 转换成功: {input_path.name} -> {output_path.name}")
                return True, str(output_path)
            else:
                error_msg = result.stderr or result.stdout
                logger.error(f"FFmpeg 转换失败: {error_msg}")
                return False, f"FFmpeg 转换失败: {error_msg}"

        except subprocess.TimeoutExpired:
            logger.error("FFmpeg 转换超时")
            return False, "FFmpeg 转换超时"
        except Exception as e:
            logger.error(f"FFmpeg 转换异常: {str(e)}")
            return False, f"FFmpeg 转换异常: {str(e)}"


def convert_wmf_to_png(input_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
    """
    便捷函数：将 WMF 文件转换为 PNG

    Args:
        input_path: WMF 文件路径
        output_path: 输出 PNG 文件路径（可选）

    Returns:
        (成功标志, 输出文件路径或错误信息)
    """
    converter = ImageConverter()
    return converter.convert(input_path, output_path, '.png')


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    
    converter = ImageConverter()
    print(f"检测到的转换工具: {converter.converter_tool}")
    print(f"支持的输入格式: {converter.SUPPORTED_INPUT_FORMATS}")
    print(f"支持的输出格式: {converter.SUPPORTED_OUTPUT_FORMATS}")
