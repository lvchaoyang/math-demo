"""
Enhanced Pandoc Converter v3.0 (Optimized Version)
==================================================
新增优化：
1. ✅ 智能缓存机制 - 基于公式内容哈希，避免重复编译
2. ✅ 编译监控统计 - 记录成功率、失败原因、性能指标
3. ✅ 前端渲染指导 - 提供详细的状态字段和渲染建议

功能清单：
1. 双重策略解析：LaTeX 源码提取 + WebTeX 图片兜底 (解决 MathType 丢失)。
2. 智能媒体关联：自动匹配公式引用的图片路径。
3. 【核心修复】LaTeX 编译引擎：自动注入导言区，调用本地 latex 引擎将公式编译为 PNG/PDF，
   彻底解决前端渲染乱码、公式挤压、字体缺失问题。
4. 详细日志与错误处理。

依赖要求:
    - Pandoc >= 2.7
    - LaTeX Engine (TeX Live / MacTeX / MiKTeX) -> 用于生成高质量图片
    - poppler-utils (pdftoppm) 或 ImageMagick -> 用于 PDF 转 PNG
"""

import subprocess
import tempfile
import shutil
import json
import re
import os
import logging
import hashlib
import time
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class CompilationStats:
    """LaTeX 编译统计信息"""
    total_attempts: int = 0
    successful: int = 0
    failed: int = 0
    cache_hits: int = 0
    cache_misses: int = 0
    total_compile_time: float = 0.0
    errors: List[Dict[str, Any]] = field(default_factory=list)
    
    @property
    def success_rate(self) -> float:
        if self.total_attempts == 0:
            return 0.0
        return (self.successful / self.total_attempts) * 100
    
    @property
    def cache_hit_rate(self) -> float:
        total = self.cache_hits + self.cache_misses
        if total == 0:
            return 0.0
        return (self.cache_hits / total) * 100
    
    @property
    def avg_compile_time(self) -> float:
        if self.successful == 0:
            return 0.0
        return self.total_compile_time / self.successful
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'total_attempts': self.total_attempts,
            'successful': self.successful,
            'failed': self.failed,
            'cache_hits': self.cache_hits,
            'cache_misses': self.cache_misses,
            'success_rate': f"{self.success_rate:.2f}%",
            'cache_hit_rate': f"{self.cache_hit_rate:.2f}%",
            'total_compile_time': f"{self.total_compile_time:.2f}s",
            'avg_compile_time': f"{self.avg_compile_time:.2f}s",
            'recent_errors': self.errors[-5:] if self.errors else []
        }


@dataclass
class RenderingGuidance:
    """前端渲染指导信息"""
    strategy: str  # 'image_first', 'latex_first', 'hybrid'
    priority: List[str]  # 渲染优先级列表
    fallback_chain: List[str]  # 降级链
    notes: List[str]  # 注意事项
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'strategy': self.strategy,
            'priority': self.priority,
            'fallback_chain': self.fallback_chain,
            'notes': self.notes
        }


class EnhancedPandocConverter:
    """基于 Pandoc 的增强型 DOCX 转换器 v3.0 - 带缓存和监控"""
    
    def __init__(self, cache_dir: Optional[str] = None):
        self.pandoc_path = self._detect_pandoc()
        self.latex_engine = self._detect_latex_engine()
        
        self.strategies = {
            'latex_raw': ['-t', 'latex', '--wrap=none'],
            'html_mathjax': ['-t', 'html', '--mathjax', '--wrap=none'],
            'html_webtex': ['-t', 'html', '--webtex', '--wrap=none'],
            'markdown': ['-t', 'markdown', '--wrap=none']
        }
        
        self.cache_dir = Path(cache_dir) if cache_dir else Path.cwd() / '.formula_cache'
        
        try:
            self.cache_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            fallback_dir = Path(tempfile.gettempdir()) / 'formula_cache'
            self.cache_dir = fallback_dir
            self.cache_dir.mkdir(parents=True, exist_ok=True)
            logger.warning(f"无法在当前目录创建缓存，使用临时目录: {self.cache_dir}")
        
        self.cache_index_file = self.cache_dir / 'cache_index.json'
        self.cache_index = self._load_cache_index()
        
        self.stats = CompilationStats()
        
        logger.info(f"缓存目录: {self.cache_dir}")
        logger.info(f"缓存索引: {len(self.cache_index)} 条记录")
    
    def _load_cache_index(self) -> Dict[str, Dict[str, Any]]:
        """加载缓存索引"""
        if self.cache_index_file.exists():
            try:
                with open(self.cache_index_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.warning(f"加载缓存索引失败: {e}")
        return {}
    
    def _save_cache_index(self):
        """保存缓存索引"""
        try:
            with open(self.cache_index_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache_index, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"保存缓存索引失败: {e}")
    
    def _get_formula_hash(self, latex_code: str) -> str:
        """计算公式的哈希值"""
        normalized = re.sub(r'\s+', ' ', latex_code.strip())
        return hashlib.sha256(normalized.encode('utf-8')).hexdigest()[:16]
    
    def _check_cache(self, formula_hash: str) -> Optional[Dict[str, Any]]:
        """检查缓存是否存在"""
        if formula_hash in self.cache_index:
            cache_info = self.cache_index[formula_hash]
            cached_file = Path(cache_info['path'])
            
            if cached_file.exists():
                self.stats.cache_hits += 1
                logger.debug(f"缓存命中: {formula_hash}")
                return {
                    'cached': True,
                    'path': str(cached_file),
                    'created_at': cache_info.get('created_at'),
                    'size': cached_file.stat().st_size
                }
        
        self.stats.cache_misses += 1
        return None
    
    def _update_cache(self, formula_hash: str, image_path: str, latex_code: str):
        """更新缓存"""
        self.cache_index[formula_hash] = {
            'path': str(image_path),
            'latex_preview': latex_code[:50],
            'created_at': datetime.now().isoformat(),
            'size': Path(image_path).stat().st_size
        }
        self._save_cache_index()
        logger.debug(f"缓存已更新: {formula_hash}")
    
    def _detect_pandoc(self) -> Optional[str]:
        """检测 Pandoc 是否可用"""
        pandoc = shutil.which('pandoc')
        if not pandoc:
            logger.error("未找到 Pandoc 可执行文件。请安装 Pandoc。")
            return None
        
        try:
            result = subprocess.run([pandoc, '--version'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version_info = result.stdout.split('\n')[0]
                logger.info(f"Pandoc 已就绪: {version_info}")
                
                version_match = re.search(r'(\d+)\.(\d+)', version_info)
                if version_match:
                    major, minor = int(version_match.group(1)), int(version_match.group(2))
                    if major < 2 or (major == 2 and minor < 7):
                        logger.warning(f"Pandoc 版本过低 ({major}.{minor})，建议升级到 2.7+。")
                
                return pandoc
        except Exception as e:
            logger.error(f"检测 Pandoc 版本失败: {e}")
        
        return None

    def _detect_latex_engine(self) -> Optional[str]:
        """检测系统中可用的 LaTeX 引擎"""
        engines = ['xelatex', 'pdflatex', 'lualatex']
        for engine in engines:
            path = shutil.which(engine)
            if path:
                logger.info(f"LaTeX 引擎就绪: {engine} (用于公式图片渲染)")
                return engine
        logger.warning("未找到 LaTeX 引擎。将无法生成高质量公式图片，只能返回 LaTeX 源码。")
        logger.warning("提示：安装 TeX Live (Linux/Win) 或 MacTeX (Mac) 以修复公式显示乱码问题。")
        return None

    def is_available(self) -> bool:
        return self.pandoc_path is not None
    
    def get_stats(self) -> Dict[str, Any]:
        """获取编译统计信息"""
        return self.stats.to_dict()
    
    def clear_cache(self):
        """清空缓存"""
        try:
            for cache_file in self.cache_dir.glob('*'):
                if cache_file.is_file():
                    cache_file.unlink()
            self.cache_index.clear()
            self._save_cache_index()
            logger.info("缓存已清空")
        except Exception as e:
            logger.error(f"清空缓存失败: {e}")

    def convert_with_strategy(
        self, 
        docx_path: str, 
        strategy: str = 'latex_raw',
        extract_media: bool = True,
        media_dir: Optional[str] = None,
        standalone: bool = False
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """使用指定策略转换文档"""
        if not self.is_available():
            return False, "Pandoc 不可用", {}

        docx_path = Path(docx_path).resolve()
        if not docx_path.exists():
            return False, f"文件不存在: {docx_path}", {}

        if strategy not in self.strategies:
            return False, f"不支持的策略: {strategy}", {}

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                actual_media_dir = ""
                real_media_path = None
                
                if extract_media:
                    if media_dir:
                        actual_media_dir = media_dir
                        Path(actual_media_dir).mkdir(parents=True, exist_ok=True)
                    else:
                        actual_media_dir = str(temp_path / 'media')
                        Path(actual_media_dir).mkdir(parents=True, exist_ok=True)
                    
                    real_media_path = Path(actual_media_dir) / 'media'

                cmd = [
                    self.pandoc_path,
                    str(docx_path),
                    *self.strategies[strategy],
                ]

                if extract_media and actual_media_dir:
                    cmd.extend(['--extract-media', actual_media_dir])

                logger.debug(f"执行 Pandoc 命令: {' '.join(cmd)}")

                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=120
                )

                if result.returncode != 0:
                    error_detail = result.stderr.strip() or result.stdout.strip()
                    logger.error(f"Pandoc 转换失败: {error_detail}")
                    return False, f"转换错误: {error_detail}", {}

                content = result.stdout
                
                media_files = []
                if extract_media and real_media_path and real_media_path.exists():
                    for f in real_media_path.rglob('*'):
                        if f.is_file():
                            rel_path = f.relative_to(real_media_path)
                            file_type = 'image' if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.svg', '.gif'] else 'other'
                            media_files.append({
                                'filename': str(rel_path),
                                'absolute_path': str(f),
                                'type': file_type,
                                'size': f.stat().st_size
                            })
                
                metadata = {
                    'strategy': strategy,
                    'media_count': len(media_files),
                    'media_files': media_files,
                    'media_dir': str(real_media_path) if (extract_media and real_media_path) else None
                }

                logger.info(f"转换成功 (策略: {strategy}), 提取媒体文件: {len(media_files)}")
                return True, content, metadata

        except subprocess.TimeoutExpired:
            logger.error("转换超时")
            return False, "转换超时", {}
        except Exception as e:
            logger.exception(f"转换过程发生异常: {e}")
            return False, f"系统异常: {str(e)}", {}

    def parse_formulas_advanced(self, docx_path: str, output_dir: str) -> Dict[str, Any]:
        """
        【核心功能】高级公式解析 + 渲染修复 + 缓存 + 监控
        
        返回值包含：
        - formulas: 公式列表（带渲染指导）
        - stats: 编译统计
        - rendering_guidance: 前端渲染建议
        """
        if not self.is_available():
            return {'success': False, 'error': 'Pandoc unavailable'}

        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        formulas = []
        all_media_files = []
        raw_latex_content = ""

        logger.info("正在执行策略 1/2: 提取 LaTeX 源码并尝试编译...")
        success_latex, latex_content, meta_latex = self.convert_with_strategy(
            docx_path, 
            strategy='latex_raw', 
            extract_media=True, 
            media_dir=output_dir
        )
        
        if success_latex:
            raw_latex_content = latex_content
            all_media_files = meta_latex.get('media_files', [])
            
            raw_formulas = self._extract_and_clean_latex(latex_content)
            
            for i, formula in enumerate(raw_formulas):
                if formula['type'] == 'image_ref':
                    img_name = formula.get('image_ref')
                    matched_media = None
                    if img_name:
                        for media in all_media_files:
                            if media['filename'] == img_name or img_name in media['filename']:
                                matched_media = media
                                break
                    
                    if matched_media:
                        formula['image_path'] = matched_media['absolute_path']
                        formula['has_image_fallback'] = True
                        formula['image_filename'] = matched_media['filename']

                if formula['latex']:
                    formula_hash = self._get_formula_hash(formula['latex'])
                    formula['hash'] = formula_hash
                    
                    cache_result = self._check_cache(formula_hash)
                    if cache_result:
                        formula['rendered_image'] = cache_result['path']
                        formula['status'] = 'cached'
                        formula['cached'] = True
                        formula['rendering_guidance'] = self._get_rendering_guidance(formula).to_dict()
                    else:
                        compiled_img_path = self._compile_formula_to_image(
                            formula['latex'], 
                            output_dir, 
                            filename=f"formula_{formula_hash}"
                        )
                        if compiled_img_path:
                            formula['rendered_image'] = compiled_img_path
                            formula['status'] = 'compiled'
                            formula['cached'] = False
                            self._update_cache(formula_hash, compiled_img_path, formula['latex'])
                        else:
                            formula['status'] = 'source_only'
                            formula['cached'] = False
                        
                        formula['rendering_guidance'] = self._get_rendering_guidance(formula).to_dict()
                
                formulas.append(formula)
        else:
            logger.warning(f"LaTeX 提取失败: {latex_content}")

        logger.info("正在执行策略 2/2: 提取公式图片 (WebTeX 兜底)...")
        success_html, html_content, meta_html = self.convert_with_strategy(
            docx_path,
            strategy='html_webtex',
            extract_media=True,
            media_dir=output_dir
        )
        
        if success_html:
            html_media = meta_html.get('media_files', [])
            webtex_images = [m for m in html_media if m['type'] == 'image']
            logger.info(f"通过 WebTeX 策略发现 {len(webtex_images)} 个公式图片")
            
            existing_img_refs = {f.get('image_filename') for f in formulas if f.get('image_filename')}
            
            for img in webtex_images:
                if img['filename'] not in existing_img_refs:
                    formula = {
                        'type': 'image_only',
                        'latex': '',
                        'image_path': img['absolute_path'],
                        'image_filename': img['filename'],
                        'rendered_image': img['absolute_path'],
                        'status': 'fallback_image',
                        'note': 'MathType/OLE 对象 (仅图片)',
                        'cached': False
                    }
                    formula['rendering_guidance'] = self._get_rendering_guidance(formula).to_dict()
                    formulas.append(formula)
        else:
            logger.warning(f"WebTeX 提取失败: {html_content}")

        compiled_count = sum(1 for f in formulas if f.get('status') in ['compiled', 'cached'])
        cached_count = sum(1 for f in formulas if f.get('cached', False))
        
        return {
            'success': True,
            'formula_count': len(formulas),
            'formulas': formulas,
            'raw_latex': raw_latex_content if success_latex else None,
            'media_files': all_media_files,
            'stats': {
                **self.stats.to_dict(),
                'latex_success': success_latex,
                'webtex_success': success_html,
                'has_latex_engine': self.latex_engine is not None,
                'compiled_count': compiled_count,
                'cached_count': cached_count
            },
            'rendering_guidance': {
                'global_strategy': self._get_global_rendering_strategy(formulas),
                'recommendations': self._get_rendering_recommendations(formulas)
            }
        }

    def _get_rendering_guidance(self, formula: Dict[str, Any]) -> RenderingGuidance:
        """为单个公式生成渲染指导"""
        status = formula.get('status', 'unknown')
        
        if status in ['compiled', 'cached']:
            return RenderingGuidance(
                strategy='image_first',
                priority=['rendered_image', 'latex'],
                fallback_chain=['rendered_image -> latex (if image fails)'],
                notes=[
                    '优先显示预编译的高质量图片',
                    '如图片加载失败，可使用 MathJax/KaTeX 渲染 LaTeX 源码',
                    '图片已优化字体和布局，无需额外处理'
                ]
            )
        elif status == 'fallback_image':
            return RenderingGuidance(
                strategy='image_only',
                priority=['rendered_image'],
                fallback_chain=['rendered_image (no fallback)'],
                notes=[
                    '仅图片可用（MathType/OLE 对象）',
                    '无法转换为 LaTeX 源码',
                    '确保图片路径正确'
                ]
            )
        elif status == 'source_only':
            return RenderingGuidance(
                strategy='latex_first',
                priority=['latex'],
                fallback_chain=['latex -> placeholder (if rendering fails)'],
                notes=[
                    '仅有 LaTeX 源码，需要前端渲染',
                    '推荐使用 MathJax 或 KaTeX',
                    '注意处理特殊符号和字体'
                ]
            )
        else:
            return RenderingGuidance(
                strategy='unknown',
                priority=[],
                fallback_chain=[],
                notes=['状态未知，请检查数据']
            )

    def _get_global_rendering_strategy(self, formulas: List[Dict[str, Any]]) -> Dict[str, Any]:
        """获取全局渲染策略"""
        total = len(formulas)
        if total == 0:
            return {
                'strategy': 'none',
                'reason': '无公式',
                'performance': 'none',
                'recommendation': '文档中没有公式'
            }
        
        compiled = sum(1 for f in formulas if f.get('status') in ['compiled', 'cached'])
        fallback = sum(1 for f in formulas if f.get('status') == 'fallback_image')
        source_only = sum(1 for f in formulas if f.get('status') == 'source_only')
        
        if compiled / total >= 0.8:
            return {
                'strategy': 'image_preferred',
                'reason': f'{compiled}/{total} 公式已预编译为图片',
                'performance': 'excellent',
                'recommendation': '优先加载图片，降级到 LaTeX 渲染'
            }
        elif source_only / total >= 0.5:
            return {
                'strategy': 'latex_required',
                'reason': f'{source_only}/{total} 公式需要前端渲染',
                'performance': 'moderate',
                'recommendation': '必须集成 MathJax/KaTeX，考虑服务端编译优化'
            }
        else:
            return {
                'strategy': 'hybrid',
                'reason': '混合模式：部分图片，部分 LaTeX',
                'performance': 'good',
                'recommendation': '根据每个公式的 guidance 字段选择渲染方式'
            }

    def _get_rendering_recommendations(self, formulas: List[Dict[str, Any]]) -> List[str]:
        """获取渲染建议列表"""
        recommendations = []
        
        compiled_count = sum(1 for f in formulas if f.get('status') in ['compiled', 'cached'])
        if compiled_count > 0:
            recommendations.append(f"✅ {compiled_count} 个公式已预编译，直接显示图片即可")
        
        cached_count = sum(1 for f in formulas if f.get('cached', False))
        if cached_count > 0:
            recommendations.append(f"⚡ {cached_count} 个公式来自缓存，加载速度更快")
        
        source_only_count = sum(1 for f in formulas if f.get('status') == 'source_only')
        if source_only_count > 0:
            recommendations.append(f"⚠️ {source_only_count} 个公式需要前端 LaTeX 渲染")
            recommendations.append("   建议：集成 MathJax 3.x 或 KaTeX")
        
        fallback_count = sum(1 for f in formulas if f.get('status') == 'fallback_image')
        if fallback_count > 0:
            recommendations.append(f"🖼️ {fallback_count} 个公式仅图片可用（MathType）")
        
        if self.stats.success_rate < 80 and self.stats.total_attempts > 0:
            recommendations.append(f"⚠️ 编译成功率较低 ({self.stats.success_rate:.1f}%)，建议检查 LaTeX 环境")
        
        return recommendations

    def _compile_formula_to_image(self, latex_code: str, output_dir: str, filename: str) -> Optional[str]:
        """将 LaTeX 代码编译为 PNG 图片（带监控）"""
        if not self.latex_engine:
            logger.debug(f"无可用 LaTeX 引擎，跳过编译：{latex_code[:30]}...")
            return None

        self.stats.total_attempts += 1
        start_time = time.time()

        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                tex_file = Path(tmp_dir) / f"{filename}.tex"
                pdf_file = Path(tmp_dir) / f"{filename}.pdf"
                png_file = Path(output_dir) / f"{filename}.png"

                full_doc = r"""
\documentclass[border=2pt]{standalone}
\usepackage{amsmath, amssymb, amsfonts, amsthm}
\usepackage{mathtools}
\usepackage{geometry}
\usepackage{unicode-math}
\setmathfont{Latin Modern Math} 
\pagestyle{empty}
\begin{document}
%s
\end{document}
""" % latex_code

                with open(tex_file, 'w', encoding='utf-8') as f:
                    f.write(full_doc)

                cmd_pdf = [
                    self.latex_engine,
                    '-interaction=nonstopmode',
                    '-output-directory', tmp_dir,
                    str(tex_file)
                ]
                
                logger.debug(f"编译公式：{latex_code[:50]}...")
                res = subprocess.run(cmd_pdf, capture_output=True, timeout=15)
                
                if res.returncode != 0:
                    error_msg = res.stderr.decode()[:200] if res.stderr else 'Unknown error'
                    self.stats.failed += 1
                    self.stats.errors.append({
                        'timestamp': datetime.now().isoformat(),
                        'latex_preview': latex_code[:50],
                        'error': error_msg,
                        'type': 'compilation_failed'
                    })
                    logger.debug(f"LaTeX 编译失败：{error_msg}")
                    return None

                if not pdf_file.exists():
                    self.stats.failed += 1
                    logger.debug(f"PDF 文件未生成：{filename}")
                    return None

                try:
                    if shutil.which('pdftoppm'):
                        cmd_img = ['pdftoppm', '-png', '-r', '300', str(pdf_file), str(Path(tmp_dir)/filename)]
                        subprocess.run(cmd_img, check=True, capture_output=True, timeout=10)
                        generated_png = Path(tmp_dir) / f"{filename}-1.png"
                        if generated_png.exists():
                            shutil.move(str(generated_png), str(png_file))
                            self.stats.successful += 1
                            self.stats.total_compile_time += time.time() - start_time
                            logger.debug(f"公式编译成功 (pdftoppm): {filename}")
                            return str(png_file)
                    
                    elif shutil.which('convert'):
                        cmd_img = ['convert', '-density', '300', str(pdf_file), str(png_file)]
                        subprocess.run(cmd_img, check=True, capture_output=True, timeout=10)
                        if png_file.exists():
                            self.stats.successful += 1
                            self.stats.total_compile_time += time.time() - start_time
                            logger.debug(f"公式编译成功 (imagemagick): {filename}")
                            return str(png_file)
                            
                except Exception as img_err:
                    logger.warning(f"公式转图片失败：{img_err}，保存 PDF 作为兜底")
                    final_pdf = Path(output_dir) / f"{filename}.pdf"
                    shutil.move(str(pdf_file), str(final_pdf))
                    self.stats.successful += 1
                    self.stats.total_compile_time += time.time() - start_time
                    return str(final_pdf)

        except subprocess.TimeoutExpired:
            self.stats.failed += 1
            self.stats.errors.append({
                'timestamp': datetime.now().isoformat(),
                'latex_preview': latex_code[:50],
                'error': 'Compilation timeout',
                'type': 'timeout'
            })
            logger.debug("LaTeX 编译超时")
        except Exception as e:
            self.stats.failed += 1
            self.stats.errors.append({
                'timestamp': datetime.now().isoformat(),
                'latex_preview': latex_code[:50],
                'error': str(e),
                'type': 'exception'
            })
            logger.debug(f"编译过程异常：{e}")
        
        return None
    
    def _extract_and_clean_latex(self, latex_content: str) -> list:
        """从 LaTeX 内容中提取公式"""
        formulas = []
        
        # 提取行内公式 $...$
        inline_pattern = r'(?<!\$)\$((?:[^$]|\\\$)+?)\$(?!\$)'
        for match in re.finditer(inline_pattern, latex_content):
            raw = match.group(1).strip()
            if raw: 
                formulas.append({
                    'type': 'inline',
                    'latex': self._clean_latex(raw),
                    'context': match.group(0)[:60]
                })

        # 提取块级公式 $$...$$
        block_pattern = r'\$\$(.*?)\$\$'
        for match in re.finditer(block_pattern, latex_content, re.DOTALL):
            raw = match.group(1).strip()
            if raw:
                formulas.append({
                    'type': 'block',
                    'latex': self._clean_latex(raw),
                    'context': match.group(0)[:60]
                })

        # 提取环境公式 \begin{equation}...\end{equation}
        env_types = ['equation', 'align', 'gather', 'multline', 'eqnarray', 'split', 'alignat']
        for env in env_types:
            pattern = rf'\\begin\{{{env}\}}(.*?)\\end\{{{env}\}}'
            for match in re.finditer(pattern, latex_content, re.DOTALL):
                raw = match.group(1).strip()
                formulas.append({
                    'type': 'environment',
                    'env_name': env,
                    'latex': self._clean_latex(raw),
                    'context': match.group(0)[:60]
                })
                
        # 提取图片引用 \includegraphics{...}
        img_pattern = r'\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}'
        for match in re.finditer(img_pattern, latex_content):
             formulas.append({
                'type': 'image_ref',
                'image_ref': match.group(1),
                'latex': '', 
                'context': match.group(0)
            })

        return formulas
        for env in env_types:
            pattern = rf'\\begin\{{{env}\}}(.*?)\\end\{{{env}\}}'
            for match in re.finditer(pattern, latex_content, re.DOTALL):
                raw = match.group(1).strip()
                formulas.append({
                    'type': 'environment',
                    'env_name': env,
                    'latex': self._clean_latex(raw),
                    'context': match.group(0)[:60]
                })
                
        img_pattern = r'\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}'
        for match in re.finditer(img_pattern, latex_content):
             formulas.append({
                'type': 'image_ref',
                'image_ref': match.group(1),
                'latex': '', 
                'context': match.group(0)
            })

        return formulas

    def _clean_latex(self, latex: str) -> str:
        """清洗 LaTeX 字符串"""
        if not latex:
            return ""
        latex = re.sub(r'\n\s*\n', '\n', latex)
        lines = [line.strip() for line in latex.split('\n')]
        latex = '\n'.join(line for line in lines if line)
        return latex


def convert_docx_enhanced(docx_path: str, output_dir: str, cache_dir: Optional[str] = None) -> Dict[str, Any]:
    """便捷入口函数"""
    converter = EnhancedPandocConverter(cache_dir=cache_dir)
    if not converter.is_available():
        return {
            'success': False, 
            'error': 'Pandoc not installed.',
            'formula_count': 0,
            'formulas': [],
            'stats': converter.get_stats()
        }
    
    return converter.parse_formulas_advanced(docx_path, output_dir)


PandocConverter = EnhancedPandocConverter


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else "./extracted_formulas"
        
        print(f"🚀 开始解析: {input_file}")
        result = convert_docx_enhanced(input_file, out_dir)
        
        if result['success']:
            print(f"\n✅ 解析成功! 共提取: {result['formula_count']} 个公式")
            print(f"\n📊 编译统计:")
            stats = result['stats']
            print(f"   - 成功率: {stats['success_rate']}")
            print(f"   - 缓存命中率: {stats['cache_hit_rate']}")
            print(f"   - 平均编译时间: {stats['avg_compile_time']}")
            
            print(f"\n🎨 渲染策略:")
            guidance = result['rendering_guidance']
            print(f"   - 全局策略: {guidance['global_strategy']['strategy']}")
            print(f"   - 原因: {guidance['global_strategy']['reason']}")
            
            print(f"\n💡 渲染建议:")
            for rec in guidance['recommendations']:
                print(f"   {rec}")
            
            print("\n--- 前 5 个公式预览 ---")
            for i, f in enumerate(result['formulas'][:5]):
                status = f.get('status', 'unknown')
                cached = '⚡缓存' if f.get('cached') else ''
                if f.get('rendered_image'):
                    print(f"[{i+1}] ✅ 已渲染: {os.path.basename(f['rendered_image'])} ({status}) {cached}")
                elif f['latex']:
                    print(f"[{i+1}] 📝 LaTeX: {f['latex'][:40]}... ({status})")
                else:
                    print(f"[{i+1}] 🖼️ 仅图片: {f.get('image_filename', 'N/A')}")
        else:
            print(f"\n❌ 失败: {result.get('error')}")
    else:
        print("用法: python pandoc_converter.py <input.docx> [output_dir]")
