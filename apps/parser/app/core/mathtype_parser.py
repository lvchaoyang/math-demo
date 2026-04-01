"""
MathType OLE 对象解析器
用于从 Word 文档中提取 MathType 公式并转换为高质量图片
"""
import olefile
import struct
import subprocess
import tempfile
import shutil
import re
import hashlib
from pathlib import Path
from typing import Tuple, Optional, Dict, Any
import logging

from .mathtype_latex_engine import (
    is_latex_extraction_disabled,
    should_try_external_first,
    should_fallback_heuristic,
    try_external_latex,
    get_latex_mode,
)

logger = logging.getLogger(__name__)


class MathTypeParser:
    """MathType 公式解析器"""
    
    MTEF_SIGNATURE = b'DSMT'
    WMF_SIGNATURE = b'\xd7\xcd\xc6\x9a'
    EMF_SIGNATURE = b'\x20\x45\x4d\x46'
    
    def __init__(self):
        self.temp_dir = None
        # key: sha1(ole 文件字节) -> (latex 或 None, 状态码)
        self._latex_cache: Dict[str, Tuple[Optional[str], str]] = {}
    
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
            latex, _reason = self.extract_latex_with_detail(ole_path)
            return latex
        except Exception as e:
            logger.error(f"提取 LaTeX 失败: {e}")
            return None

    def extract_latex_with_detail(self, ole_path: str) -> Tuple[Optional[str], str]:
        """
        从 MathType 对象中提取 LaTeX，并返回失败分类码。
        """
        if is_latex_extraction_disabled():
            return None, "engine_disabled"

        try:
            file_key = self._build_file_cache_key(ole_path)
            if file_key and file_key in self._latex_cache:
                cached_latex, cached_status = self._latex_cache[file_key]
                if cached_latex:
                    return cached_latex, "cached_hit"
                return None, cached_status or "cached_empty"

            # 1) 外部引擎（auto / external）
            if should_try_external_first():
                ext_latex, ext_status = try_external_latex(ole_path)
                if ext_latex:
                    normalized = self._normalize_latex(ext_latex)
                    if normalized:
                        if file_key:
                            self._latex_cache[file_key] = (normalized, ext_status)
                        return normalized, ext_status
                    ext_status = "external_normalize_empty"
                if get_latex_mode() == "external":
                    if file_key:
                        self._latex_cache[file_key] = (None, ext_status)
                    return None, ext_status

            # 2) 启发式（auto / heuristic）
            if not should_fallback_heuristic():
                if file_key:
                    self._latex_cache[file_key] = (None, "heuristic_disabled")
                return None, "heuristic_disabled"

            ole = olefile.OleFileIO(ole_path)

            candidate_texts = []

            # 先扫描全部流，兼容不同版本 MathType OLE 布局。
            for stream in ole.listdir():
                try:
                    stream_data = ole.openstream(stream).read()
                except Exception:
                    continue
                candidate_texts.extend(self._extract_text_candidates(stream_data))

            # 兜底：至少读取 Equation Native，避免遗漏。
            if ole.exists('Equation Native'):
                native_data = ole.openstream('Equation Native').read()
                candidate_texts.extend(self._extract_text_candidates(native_data))

            ole.close()

            latex = self._select_best_latex(candidate_texts)
            if not latex:
                if file_key:
                    self._latex_cache[file_key] = (None, "no_latex_candidate")
                return None, "no_latex_candidate"
            latex = self._normalize_latex(latex)
            if not latex:
                if file_key:
                    self._latex_cache[file_key] = (None, "latex_normalize_empty")
                return None, "latex_normalize_empty"
            if file_key:
                self._latex_cache[file_key] = (latex, "ok_heuristic")
            return latex, "ok_heuristic"
            
        except Exception as e:
            logger.error(f"提取 LaTeX 失败: {e}")
            return None, "extract_exception"

    def _build_file_cache_key(self, ole_path: str) -> Optional[str]:
        try:
            with open(ole_path, "rb") as f:
                return hashlib.sha1(f.read()).hexdigest()
        except OSError:
            return None

    def _extract_text_candidates(self, data: bytes) -> list[str]:
        """从二进制流中提取可能的 LaTeX 文本片段。"""
        results: list[str] = []
        if not data:
            return results

        for encoding in ("utf-8", "utf-16le", "latin-1"):
            try:
                text = data.decode(encoding, errors="ignore")
            except Exception:
                continue
            if not text:
                continue
            results.extend(self._find_latex_like_segments(text))
        return results

    def _find_latex_like_segments(self, text: str) -> list[str]:
        """通过规则提取 latex-like 片段。"""
        candidates: list[str] = []
        if not text:
            return candidates

        # 1) 常见包裹：$$...$$ / $...$ / \[...\]
        patterns = [
            r"\$\$(.{1,2000}?)\$\$",
            r"(?<!\$)\$(.{1,1000}?)(?<!\$)\$",
            r"\\\[(.{1,2000}?)\\\]",
        ]
        for p in patterns:
            for m in re.finditer(p, text, re.DOTALL):
                seg = (m.group(1) or "").strip()
                if seg:
                    candidates.append(seg)

        # 2) 未包裹但含明显 LaTeX 命令
        for m in re.finditer(r"([^\r\n]{3,2000})", text):
            seg = (m.group(1) or "").strip()
            if self._looks_like_latex(seg):
                candidates.append(seg)

        return candidates

    def _looks_like_latex(self, text: str) -> bool:
        if not text or len(text) < 2:
            return False
        score = 0
        if "\\" in text:
            score += 1
        if any(tok in text for tok in ("\\frac", "\\sqrt", "\\sum", "\\int", "\\begin", "\\left", "\\right")):
            score += 2
        if "{" in text and "}" in text:
            score += 1
        if any(tok in text for tok in ("^", "_")):
            score += 1
        return score >= 2

    def _select_best_latex(self, candidates: list[str]) -> Optional[str]:
        """从候选中选择最可信的一条。"""
        if not candidates:
            return None

        normalized = []
        seen = set()
        for c in candidates:
            s = self._normalize_latex(c)
            if not s:
                continue
            if s in seen:
                continue
            seen.add(s)
            normalized.append(s)

        if not normalized:
            return None

        normalized.sort(key=lambda s: (self._latex_quality_score(s), len(s)), reverse=True)
        for best in normalized:
            if not self._looks_like_latex(best):
                continue
            if not self._is_plausible_math_latex(best):
                continue
            return best
        return None

    def _latex_quality_score(self, text: str) -> int:
        score = 0
        for tok in ("\\frac", "\\sqrt", "\\sum", "\\int", "\\alpha", "\\beta", "\\begin", "\\end"):
            if tok in text:
                score += 2
        score += text.count("\\")
        score += text.count("{") + text.count("}")
        return score

    def _normalize_latex(self, text: str) -> str:
        s = (text or "").strip()
        if not s:
            return ""
        s = s.replace("\x00", "")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.strip("$")
        s = s.replace("\\\\", "\\")
        # 去掉常见外围包装，统一由上层决定 inline/block
        if s.startswith("\\(") and s.endswith("\\)"):
            s = s[2:-2].strip()
        if s.startswith("\\[") and s.endswith("\\]"):
            s = s[2:-2].strip()
        # 清理明显非数学噪音
        s = re.sub(r"[^\S\r\n]+", " ", s).strip()
        if len(s) > 2000:
            s = s[:2000].rstrip()
        return s

    def _is_plausible_math_latex(self, text: str) -> bool:
        """
        过滤 OLE 噪声字符串，避免把二进制碎片误渲染成“伪 LaTeX”。
        """
        if not text:
            return False

        lower = text.lower()
        noisy_markers = (
            "design science",
            "teX input language".lower(),
            "winallbasiccodepages",
            "winallcodepages",
            "times new roman",
            "courier new",
            "mt extra",
            "dsmt",
        )
        if any(m in lower for m in noisy_markers):
            return False

        # 过多异常字符通常表示二进制解码噪声
        bad_chars = 0
        for ch in text:
            o = ord(ch)
            if ch in "\t\r\n":
                continue
            if o < 32:
                bad_chars += 1
                continue
            if 0xFFFD == o:
                bad_chars += 1
        if bad_chars > 0:
            return False

        # 数学 LaTeX 通常不会含大量自然语言片段
        alpha_words = re.findall(r"[A-Za-z]{4,}", text)
        if len(alpha_words) >= 8 and text.count("\\") < 4:
            return False

        return True
