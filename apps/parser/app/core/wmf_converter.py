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
        self.gdi_exe: Optional[str] = None
        self.converter = self._detect_converter()
    
    @staticmethod
    def _find_gdi_render_exe() -> Optional[Path]:
        """Windows：仓库内或环境变量中的 WmfGdiRender.exe（与 Node 图片路由一致）"""
        env = (os.environ.get('WMF_GDI_RENDER_EXE') or '').strip()
        if env:
            p = Path(env)
            if p.is_file():
                return p
        if os.name != 'nt':
            return None
        here = Path(__file__).resolve()
        rels = (
            'tools/wmf-gdi-render/bin/Release/net8.0-windows/WmfGdiRender.exe',
            'tools/wmf-gdi-render/bin/Debug/net8.0-windows/WmfGdiRender.exe',
        )
        for root in here.parents:
            for rel in rels:
                cand = root / rel
                if cand.is_file():
                    return cand
        return None
    
    def _detect_converter(self) -> Optional[str]:
        """检测可用的转换工具"""
        # Windows 优先 GDI+（浏览器无法显示内嵌 data:image/wmf）
        if os.name == 'nt':
            gdi = self._find_gdi_render_exe()
            if gdi:
                self.gdi_exe = str(gdi)
                logger.info('检测到 WmfGdiRender.exe (GDI+)')
                return 'gdi_render'
        
        # 优先使用 Inkscape（跨平台，效果好）
        if shutil.which('inkscape'):
            logger.info("检测到 Inkscape 工具")
            return 'inkscape'
        
        # 其次使用 ImageMagick
        converters = [
            ('magick', self._convert_with_imagemagick),
            ('convert', self._convert_with_imagemagick),
        ]
        
        for name, func in converters:
            if shutil.which(name):
                logger.info(f"检测到 ImageMagick 工具：{name}")
                return name
        
        # macOS 特有工具（最后选择）
        if shutil.which('sips'):
            logger.info("检测到 macOS sips 工具")
            return 'sips'
        
        # LibreOffice
        if shutil.which('soffice'):
            logger.info("检测到 LibreOffice 工具")
            return 'soffice'
        
        logger.warning("未检测到可用的 WMF 转换工具")
        return None
    
    def convert(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        转换 WMF 文件为 PNG
        
        Args:
            input_path: 输入 WMF 文件路径
            output_path: 输出 PNG 文件路径
            
        Returns:
            (成功标志，输出文件路径或错误信息)
        """
        input_path = Path(input_path)
        output_path = Path(output_path)
        
        if not input_path.exists():
            return False, f"输入文件不存在：{input_path}"
        
        if not self.converter:
            return False, "没有可用的 WMF 转换工具"
        
        # sips 特殊处理
        if self.converter == 'sips':
            return self._convert_with_sips(str(input_path), str(output_path))
        
        if self.converter == 'gdi_render' and self.gdi_exe:
            return self._convert_with_gdi_render(str(input_path), str(output_path))
        
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
    
    def _convert_with_gdi_render(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 tools/wmf-gdi-render 的 WmfGdiRender.exe（仅 Windows）"""
        try:
            exe = self.gdi_exe or ''
            if not exe or not Path(exe).is_file():
                return False, 'WmfGdiRender.exe 不可用'
            cmd = [exe, input_path, output_path, '--dpi', '220']
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=45)
            if result.returncode == 0 and Path(output_path).exists():
                logger.info(f'GDI+ 转换成功: {Path(input_path).name}')
                return True, output_path
            err = (result.stderr or result.stdout or '').strip()
            return False, err or 'GDI+ 转换失败'
        except subprocess.TimeoutExpired:
            return False, 'GDI+ 转换超时'
        except Exception as e:
            logger.error(f'GDI+ 转换异常: {e}')
            return False, str(e)
    
    def _convert_with_sips(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 macOS sips 命令转换（通过预览）"""
        try:
            # sips 不直接支持 WMF，需要先通过 qlmanage 生成预览
            # 使用 qlmanage 生成 PDF 预览
            temp_pdf = str(Path(output_path).with_suffix('.pdf'))
            cmd_quicklook = [
                'qlmanage',
                '-t',
                '-f', '100',
                '-o', str(Path(output_path).parent),
                input_path
            ]
            
            # qlmanage 输出到临时目录，然后复制
            import tempfile
            with tempfile.TemporaryDirectory() as temp_dir:
                cmd_quicklook = [
                    'qlmanage',
                    '-t',
                    '-f', '100',
                    '-o', temp_dir,
                    input_path
                ]
                result = subprocess.run(cmd_quicklook, capture_output=True, text=True, timeout=30)
                
                # 查找生成的文件
                temp_files = list(Path(temp_dir).glob('*'))
                if temp_files:
                    temp_file = temp_files[0]
                    # 使用 sips 转换为 PNG
                    cmd_sips = [
                        'sips',
                        '-s', 'format', 'png',
                        '-s', 'formatOptions', 'compression',
                        str(temp_file),
                        '--out', output_path
                    ]
                    result_sips = subprocess.run(cmd_sips, capture_output=True, text=True, timeout=30)
                    
                    if result_sips.returncode == 0 and Path(output_path).exists():
                        logger.info(f"sips 转换成功：{Path(input_path).name}")
                        return True, output_path
            
            logger.warning(f"sips 转换失败，尝试其他方法")
            return False, "sips 转换失败"
            
        except Exception as e:
            logger.error(f"sips 转换异常：{e}")
            return False, f"sips 转换异常：{e}"
    
    def _convert_with_inkscape(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 Inkscape 转换"""
        try:
            # 先尝试与 API 路由一致的最小参数（稳定性最高）
            cmd_simple = [
                'inkscape',
                input_path,
                '--export-type=png',
                '--export-filename=' + output_path
            ]
            result_simple = subprocess.run(cmd_simple, capture_output=True, text=True, timeout=30)
            if result_simple.returncode == 0 and Path(output_path).exists():
                logger.info(f"Inkscape 转换成功 (simple): {Path(input_path).name}")
                return True, output_path

            # simple 失败后再尝试带 DPI 的参数
            cmd_dpi = [
                'inkscape',
                input_path,
                '--export-filename', output_path,
                '--export-type', 'png',
                '--export-dpi', '600'
            ]
            result_dpi = subprocess.run(cmd_dpi, capture_output=True, text=True, timeout=30)
            if result_dpi.returncode == 0 and Path(output_path).exists():
                logger.info(f"Inkscape 转换成功 (600 DPI): {Path(input_path).name}")
                return True, output_path

            # 最后兼容旧版 inkscape 参数
            cmd_old = [
                'inkscape',
                input_path,
                '--export-png', output_path,
                '--export-dpi', '600'
            ]
            result_old = subprocess.run(cmd_old, capture_output=True, text=True, timeout=30)
            if result_old.returncode == 0 and Path(output_path).exists():
                logger.info(f"Inkscape 转换成功 (legacy): {Path(input_path).name}")
                return True, output_path

            error = (result_simple.stderr or result_simple.stdout or '') + ' | ' + \
                    (result_dpi.stderr or result_dpi.stdout or '') + ' | ' + \
                    (result_old.stderr or result_old.stdout or '')
            logger.error(f"Inkscape 转换失败：{error}")
            return False, f"Inkscape 转换失败：{error}"
                
        except subprocess.TimeoutExpired:
            return False, "Inkscape 转换超时"
        except FileNotFoundError:
            return False, "Inkscape 未找到"
        except Exception as e:
            logger.error(f"Inkscape 转换异常：{e}")
            return False, f"Inkscape 转换异常：{e}"
    
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
                logger.info(f"LibreOffice 转换成功：{Path(input_path).name}")
                return True, output_path
            else:
                error = result.stderr or result.stdout
                logger.error(f"LibreOffice 转换失败：{error}")
                return False, f"LibreOffice 转换失败：{error}"
                
        except subprocess.TimeoutExpired:
            return False, "LibreOffice 转换超时"
        except Exception as e:
            logger.error(f"LibreOffice 转换异常：{e}")
            return False, f"LibreOffice 转换异常：{e}"
    
    def _convert_with_imagemagick(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """使用 ImageMagick 转换"""
        try:
            # 先尝试获取原始尺寸
            try:
                from PIL import Image
                with Image.open(input_path) as img:
                    width, height = img.size
                # 计算合适的密度
                target_width = 1000
                calculated_density = int((target_width / width) * 72)
                calculated_density = max(150, min(600, calculated_density))
            except:
                calculated_density = 300
            
            cmd_name = 'magick' if self.converter == 'magick' else 'convert'
            cmd = [
                cmd_name,
                '-density', str(calculated_density),
                input_path,
                '-quality', '95',
                '-background', 'white',
                '-alpha', 'remove',
                output_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and Path(output_path).exists():
                logger.info(f"ImageMagick 转换成功 ({calculated_density} DPI): {Path(input_path).name}")
                return True, output_path
            else:
                error = result.stderr or result.stdout
                if 'libreoffice' in error.lower() or 'delegate' in error.lower():
                    return False, "ImageMagick 需要 LibreOffice 支持 WMF 格式"
                logger.error(f"ImageMagick 转换失败：{error}")
                return False, f"ImageMagick 转换失败：{error}"
                
        except subprocess.TimeoutExpired:
            return False, "ImageMagick 转换超时"
        except Exception as e:
            logger.error(f"ImageMagick 转换异常：{e}")
            return False, f"ImageMagick 转换异常：{e}"


def convert_wmf_to_png(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    便捷函数：转换 WMF 文件为 PNG
    
    Args:
        input_path: 输入 WMF 文件路径
        output_path: 输出 PNG 文件路径
        
    Returns:
        (成功标志，输出文件路径或错误信息)
    """
    converter = WMFConverter()
    return converter.convert(input_path, output_path)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    
    converter = WMFConverter()
    print(f"检测到的转换工具：{converter.converter}")
