"""
统一公式渲染引擎 (P2/P3)

阶段目标：
1) 提供统一入口和缓存键规则
2) 生成可审计的渲染计划（render_plan）
3) P3 首版：尝试将 OMML/LaTeX 公式预渲染为 PNG（环境可用时）
"""

from pathlib import Path
from typing import Dict, Any, List, Optional
import hashlib
import shutil
import subprocess
import tempfile
from .wmf_converter import WMFConverter


class FormulaRenderer:
    """统一公式渲染器（骨架实现）"""

    def __init__(self, cache_root: Optional[str] = None):
        default_cache = Path.cwd() / ".formula_asset_cache"
        self.cache_root = Path(cache_root) if cache_root else default_cache
        self.cache_root.mkdir(parents=True, exist_ok=True)
        self.latex_engine = self._detect_latex_engine()
        self.pdf_to_image = self._detect_pdf_to_image_tool()
        self.wmf_converter = WMFConverter()
        self.magick_cmd = shutil.which("magick")

    def build_cache_key(self, asset: Dict[str, Any]) -> str:
        source_type = str(asset.get("source_type", "unknown"))
        display_type = str(asset.get("display_type", "inline"))
        latex = str(asset.get("latex", ""))
        image_filename = str(asset.get("image_filename", ""))
        raw = f"{source_type}|{display_type}|{latex}|{image_filename}"
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]

    def build_render_plan(
        self,
        formula_assets: List[Dict[str, Any]],
        file_id: Optional[str] = None,
        output_dir: Optional[str] = None,
        render_omml: bool = True,
        source_image_dir: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        生成渲染计划。
        P3 首版：对 OMML 资产尝试预渲染 PNG；其他类型先保留规划状态。
        """
        plan: List[Dict[str, Any]] = []
        file_prefix = file_id or "unknown"
        target_root = Path(output_dir) if output_dir else self.cache_root
        target_root.mkdir(parents=True, exist_ok=True)

        for asset in formula_assets:
            cache_key = self.build_cache_key(asset)
            target_name = f"{file_prefix}_{cache_key}.png"
            target_path = str(target_root / target_name)

            source_type = asset.get("source_type")
            if source_type == "omml":
                action = "render_from_latex"
            elif source_type == "mathtype_latex":
                action = "render_from_latex"
            elif source_type == "mathtype_ole":
                action = "use_mathtype_png"
            elif source_type == "wmf_preview":
                action = "render_from_wmf_preview"
            else:
                action = "skip"

            status = "planned"
            note = ""
            rendered_image = None
            source_filename = None
            source_path = None

            if source_type in ("omml", "mathtype_latex") and render_omml:
                latex = (asset.get("latex") or "").strip()
                if not latex:
                    status = "source_only"
                    note = "empty_latex"
                else:
                    ok, msg = self._render_latex_to_png(latex, Path(target_path))
                    if ok:
                        status = "rendered"
                        rendered_image = target_path
                        asset["rendered_image"] = target_path
                    else:
                        status = "source_only"
                        note = msg
            elif source_type == "wmf_preview":
                source_filename = str(asset.get("image_filename", ""))
                if not source_filename or not source_image_dir:
                    status = "source_only"
                    note = "wmf_source_missing"
                else:
                    source_path_obj = Path(source_image_dir) / source_filename
                    source_path = str(source_path_obj)
                    ok, msg = self._render_wmf_preview_to_png(source_path_obj, Path(target_path))
                    if ok:
                        status = "rendered"
                        rendered_image = target_path
                        asset["rendered_image"] = target_path
                    else:
                        status = "source_only"
                        note = msg
            elif source_type == "mathtype_ole":
                source_filename = str(asset.get("image_filename", ""))
                if not source_filename or not source_image_dir:
                    status = "source_only"
                    note = "mathtype_source_missing"
                else:
                    source_path = str(Path(source_image_dir) / source_filename)
                    if Path(source_path).exists():
                        status = "rendered"
                        rendered_image = str(source_path)
                        asset["rendered_image"] = str(source_path)
                        note = "mathtype_png_reused"
                    else:
                        status = "source_only"
                        note = "mathtype_png_not_found"

            plan.append(
                {
                    "asset_id": asset.get("id"),
                    "source_type": source_type,
                    "action": action,
                    "cache_key": cache_key,
                    "source_filename": source_filename,
                    "source_path": source_path,
                    "target_path": target_path,
                    "status": status,
                    "note": note,
                    "rendered_image": rendered_image,
                    "meta": asset.get("meta") or {},
                }
            )

        return plan

    def _detect_latex_engine(self) -> Optional[str]:
        for name in ("xelatex", "pdflatex", "lualatex"):
            if shutil.which(name):
                return name
        return None

    def _detect_pdf_to_image_tool(self) -> Optional[str]:
        if shutil.which("pdftoppm"):
            return "pdftoppm"
        if shutil.which("magick"):
            return "magick"
        return None

    def _render_latex_to_png(self, latex: str, output_path: Path) -> (bool, str):
        if not self.latex_engine:
            return False, "latex_engine_unavailable"
        if not self.pdf_to_image:
            return False, "pdf_to_image_tool_unavailable"

        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with tempfile.TemporaryDirectory() as td:
                td_path = Path(td)
                tex_path = td_path / "formula.tex"
                pdf_path = td_path / "formula.pdf"
                png_prefix = td_path / "formula"

                tex_doc = r"""
\documentclass[border=2pt]{standalone}
\usepackage{amsmath,amssymb,mathtools}
\pagestyle{empty}
\begin{document}
%s
\end{document}
""" % latex
                tex_path.write_text(tex_doc, encoding="utf-8")

                latex_cmd = [
                    self.latex_engine,
                    "-interaction=nonstopmode",
                    "-output-directory",
                    str(td_path),
                    str(tex_path),
                ]
                latex_run = subprocess.run(latex_cmd, capture_output=True, text=True, timeout=20)
                if latex_run.returncode != 0 or not pdf_path.exists():
                    return False, "latex_compile_failed"

                if self.pdf_to_image == "pdftoppm":
                    img_cmd = ["pdftoppm", "-png", "-r", "300", str(pdf_path), str(png_prefix)]
                    img_run = subprocess.run(img_cmd, capture_output=True, text=True, timeout=15)
                    out_png = td_path / "formula-1.png"
                    if img_run.returncode == 0 and out_png.exists():
                        shutil.move(str(out_png), str(output_path))
                        return True, "ok"
                    return False, "pdftoppm_failed"

                img_cmd = ["magick", "-density", "300", str(pdf_path), str(output_path)]
                img_run = subprocess.run(img_cmd, capture_output=True, text=True, timeout=15)
                if img_run.returncode == 0 and output_path.exists():
                    return True, "ok"
                return False, "magick_failed"
        except subprocess.TimeoutExpired:
            return False, "render_timeout"
        except Exception:
            return False, "render_exception"

    def _render_wmf_preview_to_png(self, source_path: Path, output_path: Path) -> (bool, str):
        if not source_path.exists():
            return False, "wmf_source_not_found"
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            success, result = self.wmf_converter.convert(str(source_path), str(output_path))
            if success and output_path.exists():
                # 二次归一化：仅做等比 contain，避免 ! 拉伸导致字符挤压/重叠
                self._normalize_png_contain_safe(output_path)
                return True, "ok"
            return False, f"wmf_convert_failed:{result}"
        except Exception:
            return False, "wmf_render_exception"

    def _normalize_png_contain_safe(self, image_path: Path) -> None:
        """仅做安全的等比归一化，不执行强制拉伸。"""
        if not self.magick_cmd or not image_path.exists():
            return

        try:
            # 读取原始尺寸
            identify = subprocess.run(
                [self.magick_cmd, "identify", "-format", "%w %h", str(image_path)],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if identify.returncode != 0:
                return
            parts = identify.stdout.strip().split()
            if len(parts) != 2:
                return
            width = int(parts[0])
            height = int(parts[1])
            if width <= 0 or height <= 0:
                return

            # 估算当前内容占比，已足够则跳过
            trim_info = subprocess.run(
                [self.magick_cmd, "convert", str(image_path), "-trim", "+repage", "-format", "%w %h", "info:"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if trim_info.returncode != 0:
                return
            trim_parts = trim_info.stdout.strip().split()
            if len(trim_parts) != 2:
                return
            trim_w = int(trim_parts[0])
            trim_h = int(trim_parts[1])
            if trim_w <= 0 or trim_h <= 0:
                return
            fill_ratio = trim_w / width
            if fill_ratio >= 0.92:
                return

            temp_contain = image_path.with_suffix(".contain.tmp.png")

            # 仅 contain：保持长宽比，最多居中留白，不改变字形几何
            contain_cmd = [
                self.magick_cmd,
                "convert",
                str(image_path),
                "-trim",
                "+repage",
                "-filter",
                "Lanczos",
                "-resize",
                f"{width}x{height}",
                "-gravity",
                "center",
                "-extent",
                f"{width}x{height}",
                str(temp_contain),
            ]
            contain_run = subprocess.run(contain_cmd, capture_output=True, text=True, timeout=20)
            if contain_run.returncode == 0 and temp_contain.exists():
                shutil.move(str(temp_contain), str(image_path))

            if temp_contain.exists():
                temp_contain.unlink()
        except Exception:
            return
