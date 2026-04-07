"""
Word 导出器
将选中的题目导出为 Word 文档，支持水印
"""

from typing import List, Dict, Any, Optional
from pathlib import Path
from datetime import datetime
import io
import os
import re
import tempfile
import logging
import html as html_module
from urllib.parse import unquote

from PIL import Image
from bs4 import BeautifulSoup, NavigableString, Tag

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement, parse_xml

from .splitter import Question
from .wmf_converter import WMFConverter
from .latex_omml_export import latex_to_omml_element, should_try_omml, strip_latex_delimiters
from .convert_equations_cli import (
    get_convert_equations_exe,
    run_latex_to_mathtype_payload,
    decode_math_type_model,
)
from .mathtype_ole_embed import embed_mathtype_ole_in_paragraph, minimal_preview_png_bytes

logger = logging.getLogger(__name__)

# PIL 不可用或读图失败时的回退宽度（英寸）
_FALLBACK_BLOCK_IN = 1.65
_FALLBACK_INLINE_IN = 0.64
_FALLBACK_OPTION_IN = 0.34
# 首选「宽×高」盒（英寸）；行内 max_h 不宜过大，否则瘦高图（△、竖分式）会压过汉字
_BOX_BLOCK = (1.75, 0.90)
_BOX_INLINE = (1.52, 0.32)
_BOX_OPTION = (0.40, 0.205)
_BOX_DEFAULT = (1.36, 0.32)
# 为满足最小高度所需的宽度可超过盒宽，但不超过下列硬上限（避免撑破版心）
_HARD_MAX_W_BLOCK = 2.12
_HARD_MAX_W_INLINE = 2.72
_HARD_MAX_W_DEFAULT = 2.15
_HARD_MAX_W_OPTION = 0.43
# 各类别最小显示高度（英寸），与 12pt 汉字行高约对齐，避免公式像脚注
_MIN_DISPLAY_H_BLOCK = 0.138
_MIN_DISPLAY_H_INLINE = 0.124
_MIN_DISPLAY_H_OPTION = 0.064
_MIN_DISPLAY_H_DEFAULT = 0.115
# 最小显示宽度（英寸）
_MIN_W_BLOCK = 0.152
_MIN_W_INLINE = 0.142
_MIN_W_OPTION = 0.108
_MIN_W_DEFAULT = 0.132
# 行内图：像素「宽/高」较小时（瘦高符号），再压低显示高度上限，避免单符号过大
_INLINE_TALL_AR_CAP = 0.62
_INLINE_TALL_MAX_H_IN = 0.28

# 导出 docx 元数据（避免默认「python-docx」「Macintosh Word」等与资源管理器/属性页观感异常）
_DOC_META_AUTHOR = "题目导出"
_DOC_META_LAST_MODIFIED_BY = "题目导出"
_APP_XML_MAC_WORD = b"Microsoft Macintosh Word"
_APP_XML_WIN_WORD = b"Microsoft Office Word"


class WatermarkType:
    """水印类型"""
    TEXT = "text"
    IMAGE = "image"


class WordExporter:
    """Word 文档导出器"""
    
    def __init__(self):
        self.document = None
        self.image_dir: Optional[Path] = None
        # data/images，其下每个子目录为一个上传 file_id（与 HTML 里 /api/v1/images/{file_id}/ 一致）
        self.images_library_root: Optional[Path] = None
        self._embedded_basenames: set = set()
        # 单题内同一磁盘文件只嵌一次（_fill_inline / 列表补图 / 历史逻辑重叠时防重复）
        self._embedded_resolved_paths: set = set()
        # 已成功处理过的 <img src>（规范化路径键），避免 DOM 嵌完后正则再往段尾插一遍
        self._embedded_img_src_keys: set = set()
        self._omml_fail_latex: set = set()
        self._convert_equations_payload_cache: Dict[str, Optional[Dict[str, Any]]] = {}

    def create_document(self, title: str = "", watermark_text: str = None, image_dir: Optional[str] = None) -> Document:
        """
        创建新文档
        
        Args:
            title: 文档标题
            watermark_text: 水印文字
            
        Returns:
            Document 对象
        """
        self.document = Document()
        self.image_dir = Path(image_dir) if image_dir else None
        self.images_library_root = self.image_dir.parent if self.image_dir else None
        self._embedded_basenames = set()
        self._embedded_resolved_paths = set()
        self._embedded_img_src_keys = set()
        self._omml_fail_latex = set()
        self._convert_equations_payload_cache = {}

        # 设置默认字体
        self._set_default_font()
        
        # 添加标题
        if title:
            self._add_title(title)
            
        # 添加水印
        if watermark_text:
            self._add_watermark(watermark_text)

        self._apply_document_metadata(title=title)
        return self.document

    def _apply_document_metadata(self, title: str = "") -> None:
        """覆盖 python-docx 默认属性，使 Windows 下作者/备注/程序名更接近 Word 本机保存。"""
        if not self.document:
            return
        cp = self.document.core_properties
        cp.author = _DOC_META_AUTHOR
        cp.last_modified_by = _DOC_META_LAST_MODIFIED_BY
        cp.title = (title or "").strip() or "导出的题目"
        cp.subject = ""
        cp.keywords = ""
        cp.category = ""
        cp.comments = ""
        now = datetime.now()
        cp.created = now
        cp.modified = now
        try:
            cp.revision = 1
        except Exception:
            pass
        self._patch_app_xml_application()

    def _refresh_document_modified_time(self) -> None:
        """保存前刷新修改时间。"""
        if self.document:
            self.document.core_properties.modified = datetime.now()

    def _patch_app_xml_application(self) -> None:
        """将 docProps/app.xml 中默认的 Macintosh Word 改为 Office Word（仅替换字节串）。"""
        if not self.document:
            return
        for part in self.document.part.package.iter_parts():
            if str(part.partname) != "/docProps/app.xml":
                continue
            blob = part.blob
            if not blob:
                break
            if _APP_XML_MAC_WORD in blob:
                part._blob = blob.replace(_APP_XML_MAC_WORD, _APP_XML_WIN_WORD)
            break

    def _set_default_font(self):
        """设置默认字体"""
        # 设置正文样式
        style = self.document.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        
        # 设置中文字体
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        
    def _add_title(self, title: str):
        """添加文档标题"""
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(title)
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        
        # 添加空行
        self.document.add_paragraph()
        
    def _add_watermark(self, text: str):
        """
        添加文字水印
        
        通过修改文档的 header 来实现水印效果
        """
        # 获取或创建 header
        section = self.document.sections[0]
        header = section.header
        
        # 创建水印的 XML
        watermark_xml = f'''
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:w10="urn:schemas-microsoft-com:office:word">
            <w:pict>
                <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" 
                    path="m@7,l@8,m@5,21600l@6,21600e">
                    <v:formulas>
                        <v:f eqn="sum #0 0 10800"/>
                        <v:f eqn="prod #0 2 1"/>
                        <v:f eqn="sum 21600 0 @1"/>
                        <v:f eqn="sum 0 0 @2"/>
                        <v:f eqn="sum 21600 0 @3"/>
                        <v:f eqn="if @0 @3 0"/>
                        <v:f eqn="if @0 21600 @1"/>
                        <v:f eqn="if @0 0 @2"/>
                        <v:f eqn="if @0 @4 21600"/>
                        <v:f eqn="mid @5 @6"/>
                        <v:f eqn="mid @8 @5"/>
                        <v:f eqn="mid @7 @8"/>
                        <v:f eqn="mid @6 @7"/>
                        <v:f eqn="sum @6 0 @5"/>
                    </v:formulas>
                    <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
                    <v:textpath on="t" fitshape="t"/>
                    <v:handles>
                        <v:h position="#0,bottomRight" xrange="6629,14971"/>
                    </v:handles>
                    <o:lock v:ext="edit" text="t" shapetype="t"/>
                </v:shapetype>
                <v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" 
                    style="position:absolute;margin-left:0;margin-top:0;width:500pt;height:100pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                    o:allowincell="f" fillcolor="silver" stroked="f">
                    <v:fill opacity=".5"/>
                    <v:textpath style="font-family:&quot;宋体&quot;;font-size:60pt" string="{text}"/>
                    <w10:wrap anchorx="margin" anchory="margin"/>
                </v:shape>
            </w:pict>
        </w:r>
        '''
        
        # 添加水印到 header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        run = paragraph.add_run()
        
        # 解析并插入水印 XML
        # python-docx 没有 OxmlElement.fromstring，这里使用 parse_xml。
        try:
            paragraph._p.append(parse_xml(watermark_xml))
        except Exception:
            # 水印失败不应阻断导出主流程
            pass
        
    def add_question(self, question: Question, include_answer: bool = False, 
                     include_analysis: bool = False):
        """
        添加题目到文档
        
        Args:
            question: 题目对象
            include_answer: 是否包含答案
            include_analysis: 是否包含解析
        """
        # 每题重置「本题内已嵌图片」集合，避免跨题误伤；题干与选项共用本题集合
        self._embedded_resolved_paths.clear()
        self._embedded_img_src_keys.clear()
        # 题号
        para = self.document.add_paragraph()
        run = para.add_run(f"{question.number}. ")
        run.font.bold = True
        run.font.size = Pt(12)
        
        # 题目内容：优先用 content_html（含公式图 <img>），否则纯文本 + 尾部图片列表
        stem_html = (question.content_html or '').strip()
        if stem_html:
            stem_html = html_module.unescape(stem_html)
            stem_html = stem_html.replace('\\"', '"').replace("\\'", "'").replace("\\/", "/")
            # 用 div 根节点，避免块级 <div> 包在 <span> 内被浏览器式解析拆乱
            root = BeautifulSoup(f'<div class="stem-root">{stem_html}</div>', 'html.parser').select_one(
                '.stem-root'
            )
            if root:
                self._strip_leading_question_number(root, question.number)
                self._fill_stem_root(para, root)
            # HTML 含 <img> 时顺序以 DOM 为准；勿再向题号段尾正则/列表补图，否则会与 content_html 顺序不一致
            stem_has_img = bool(re.search(r"<img\b", stem_html, re.I))
            if not stem_has_img:
                self._add_missing_images_inline(para, question.images)
        else:
            self._add_formatted_text(para, question.content)
            self._add_question_images(question.images)
        
        # 选项（选择题）
        if question.options:
            self._add_options(question.options)
            
        # 答案（与题干一致：$...$ 走 MathType OLE → OMML → 斜体回退）
        if include_answer and question.answer:
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.3)
            run = para.add_run("【答案】")
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 128, 0)
            self._add_formatted_text(para, question.answer, font_size_pt=10.5)
            
        # 解析
        if include_analysis and question.analysis:
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.3)
            run = para.add_run("【解析】")
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 0, 255)
            self._add_formatted_text(para, question.analysis, font_size_pt=10.5)
            
        # 添加空行
        self.document.add_paragraph()

    def _resolve_image_path(self, filename: str) -> Optional[Path]:
        """解析图片在本地磁盘中的真实路径"""
        if not self.image_dir:
            return None

        filename = (filename or "").strip().replace("\\", "/")
        if not filename:
            return None

        # 支持 filename 中可能带有 media/ 前缀
        candidates = [
            self.image_dir / filename,
            self.image_dir / "media" / filename,
            self.image_dir / Path(filename).name,
        ]
        if filename.startswith("media/"):
            candidates.append(self.image_dir / filename.replace("media/", "", 1))

        for p in candidates:
            if p.exists() and p.is_file():
                return p

        # URL 与磁盘子目录不一致时，按文件名递归查找（与 parser 图片路由行为一致）
        base = Path(filename).name
        if self.image_dir.is_dir() and base:
            try:
                for p in self.image_dir.rglob(base):
                    if p.is_file():
                        return p
            except OSError as e:
                logger.warning("rglob 图片失败: %s", e)

        # HTML 里 file_id 与当前导出会话不一致时（缓存/复制题目），在整个 data/images 下按文件名找
        if self.images_library_root and self.images_library_root.is_dir() and base:
            try:
                for p in self.images_library_root.rglob(base):
                    if p.is_file():
                        return p
            except OSError as e:
                logger.warning("库根 rglob 图片失败: %s", e)
        return None

    def _resolve_path_for_img_src(self, src: str) -> Optional[Path]:
        """按 /api/v1/images/{file_id}/filename 定位磁盘文件（与 Node 图片路由一致）"""
        src = self._normalize_img_src(src)
        m = re.search(r"(?:https?://[^/]+)?/api/v1/images/([^/]+)/([^?#]+)", src)
        if not m or not self.images_library_root:
            return None
        fid = m.group(1)
        fn = unquote(m.group(2).strip())
        sub = self.images_library_root / fid
        candidates = [
            sub / fn,
            sub / Path(fn).name,
            sub / "media" / Path(fn).name,
        ]
        for p in candidates:
            if p.is_file():
                return p
        name = Path(fn).name
        if sub.is_dir():
            try:
                for p in sub.rglob(name):
                    if p.is_file():
                        return p
            except OSError:
                pass
        return None

    def _add_question_images(self, images: List[str]):
        """将题目关联图片插入到文档"""
        if not images:
            return

        # 去重但保序，避免同图重复写入
        deduped = list(dict.fromkeys(images))
        for img in deduped:
            if Path(img).name in self._embedded_basenames:
                continue
            img_path = self._resolve_image_path(img)
            if not img_path:
                continue
            embed = self._embed_image_path_for_word(img_path)
            if not embed:
                continue
            try:
                p = self.document.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                w_in = self._picture_display_width_inches(embed, ['formula-image-block'])
                p.add_run().add_picture(str(embed), width=Inches(w_in))
                self._embedded_basenames.add(Path(img_path).name)
                try:
                    self._embedded_resolved_paths.add(str(img_path.resolve()))
                except OSError:
                    self._embedded_resolved_paths.add(str(img_path))
            except Exception:
                continue
            finally:
                self._cleanup_temp_embed(embed, img_path)
        
    def _img_src_dedupe_key(self, src: str) -> str:
        """同一图片 URL 不同写法（主机、转义）映射为同一键，供 src 级去重。"""
        s = self._normalize_img_src(src)
        if not s or s.startswith("data:"):
            return ""
        m = re.search(r"(?:https?://[^/]+)?(/api/v1/images/[^?#]+)", s)
        if m:
            try:
                return unquote(m.group(1).rstrip("/"))
            except Exception:
                return m.group(1).rstrip("/")
        return s.split("?")[0].strip()

    def _normalize_img_src(self, src: str) -> str:
        """消除 JSON/HTML 序列化产生的转义，便于解析真实 URL"""
        if not src:
            return ""
        s = src.strip()
        s = s.replace("\\/", "/")
        s = s.replace('\\"', '"').replace("\\'", "'")
        s = html_module.unescape(s)
        return s.strip().strip('"').strip("'")

    def _src_to_filename(self, src: str) -> Optional[str]:
        """从 /api/v1/images/{file_id}/xxx.png 等 URL 取出路径最后一段文件名"""
        src = self._normalize_img_src(src)
        if not src or src.startswith("data:"):
            return None
        # 支持绝对 URL: http://host/api/v1/images/...
        m = re.search(r"(?:https?://[^/]+)?/api/v1/images/[^/]+/([^?#]+)", src)
        if m:
            return unquote(m.group(1))
        try:
            return unquote(Path(src).name)
        except Exception:
            return None

    def _embed_image_path_for_word(self, path: Path) -> Optional[Path]:
        """python-docx 对 WMF 支持差，先转为 PNG 再嵌入"""
        if not path.exists():
            return None
        suf = path.suffix.lower()
        if suf not in ('.wmf', '.emf'):
            return path
        conv = WMFConverter()
        fd, tmp_png = tempfile.mkstemp(suffix='.png')
        os.close(fd)
        ok, _ = conv.convert(str(path), tmp_png)
        if ok and Path(tmp_png).exists():
            return Path(tmp_png)
        try:
            Path(tmp_png).unlink(missing_ok=True)
        except OSError:
            pass
        return None

    @staticmethod
    def _cleanup_temp_embed(embed: Path, original: Path) -> None:
        if embed != original and embed.exists():
            try:
                embed.unlink()
            except OSError:
                pass

    def _fallback_width_for_classes(self, classes: List[str]) -> float:
        cls = set(classes or [])
        if 'formula-image-block' in cls or 'math-block' in cls:
            return _FALLBACK_BLOCK_IN
        if 'option-image' in cls:
            return _FALLBACK_OPTION_IN
        if 'formula-image' in cls:
            return _FALLBACK_INLINE_IN
        return _FALLBACK_INLINE_IN

    @staticmethod
    def _box_floors_and_hard_cap(classes: List[str]):
        cls = set(classes or [])
        if 'formula-image-block' in cls or 'math-block' in cls:
            return _BOX_BLOCK, _MIN_DISPLAY_H_BLOCK, _MIN_W_BLOCK, _HARD_MAX_W_BLOCK
        if 'option-image' in cls:
            return _BOX_OPTION, _MIN_DISPLAY_H_OPTION, _MIN_W_OPTION, _HARD_MAX_W_OPTION
        if 'formula-image' in cls:
            return _BOX_INLINE, _MIN_DISPLAY_H_INLINE, _MIN_W_INLINE, _HARD_MAX_W_INLINE
        return _BOX_DEFAULT, _MIN_DISPLAY_H_DEFAULT, _MIN_W_DEFAULT, _HARD_MAX_W_DEFAULT

    def _picture_display_width_inches(self, embed_path: Path, classes: List[str]) -> float:
        """
        在盒内 contain；若高度低于 min_h，可放宽到 hard_max_w（解决长横式在盒宽下过扁）。
        瘦高位图（ar 小）额外压低允许高度，减轻 △、分式等单符号撑满一行的问题。
        """
        (max_w, max_h), min_h_floor, min_w_floor, hard_max_w = self._box_floors_and_hard_cap(
            classes
        )
        cls_set = set(classes or [])
        try:
            with Image.open(embed_path) as im:
                w_px, h_px = im.size
            if w_px < 1 or h_px < 1:
                return self._fallback_width_for_classes(classes)
            ar = w_px / float(h_px)
            # 块级整段公式不压；选项单独一套盒；行内/默认对瘦高图收紧高度上限
            mh = max_h
            if (
                'formula-image-block' not in cls_set
                and 'math-block' not in cls_set
                and 'option-image' not in cls_set
                and ar < _INLINE_TALL_AR_CAP
            ):
                mh = min(mh, _INLINE_TALL_MAX_H_IN)
            w_in = min(max_w, mh * ar)
            h_in = w_in / ar
            if h_in < min_h_floor:
                # 需要更宽才能达到 min_h；允许超过盒宽，直至 hard_max_w
                w_in = min(hard_max_w, min_h_floor * ar)
                h_in = w_in / ar
            if h_in > mh:
                w_in = mh * ar
                h_in = mh
            w_in = min(w_in, hard_max_w)
            if w_in < min_w_floor and (min_w_floor / ar) <= mh + 1e-9:
                w_in = min_w_floor
            return float(w_in)
        except Exception:
            return self._fallback_width_for_classes(classes)

    @staticmethod
    def _strip_leading_question_number(root: Tag, number: Any) -> None:
        """题干 HTML 若已含「1.」等与题号重复的前缀，去掉首处文本，避免出现「1. 1.」。"""
        if number is None or root is None:
            return
        try:
            n = int(number)
        except (TypeError, ValueError):
            return
        pat = re.compile(rf"^\s*{n}[\.\、．]\s*")
        for node in root.descendants:
            if not isinstance(node, NavigableString):
                continue
            s = str(node)
            if not s.strip():
                continue
            if pat.match(s):
                node.replace_with(NavigableString(pat.sub("", s, count=1)))
            break

    def _add_picture_to_run(self, run, path: Path, classes: List[str]) -> None:
        try:
            key = str(path.resolve())
        except OSError:
            key = str(path)
        if key in self._embedded_resolved_paths:
            return
        embed = self._embed_image_path_for_word(path)
        if not embed:
            return
        try:
            w = self._picture_display_width_inches(embed, classes)
            run.add_picture(str(embed), width=Inches(w))
            self._embedded_basenames.add(Path(path).name)
            self._embedded_resolved_paths.add(key)
        except Exception as e:
            logger.warning('嵌入图片失败: %s', e)
        finally:
            self._cleanup_temp_embed(embed, path)

    def _add_missing_images_inline(self, paragraph, images: List[str]) -> None:
        """按 question.images 补嵌尚未成功的文件（与 HTML/正则路径去重靠 _embedded_resolved_paths）。"""
        if not images or not (self.images_library_root or self.image_dir):
            return
        for img in dict.fromkeys(images):
            base = Path(img).name
            pth = self._resolve_image_path(img) or self._resolve_image_path(base)
            if not pth:
                continue
            try:
                k = str(pth.resolve())
            except OSError:
                k = str(pth)
            if k in self._embedded_resolved_paths:
                continue
            run = paragraph.add_run()
            self._add_picture_to_run(run, pth, ['formula-image'])

    def _extract_latex_from_math_element(self, element: Tag) -> str:
        raw = (element.get("data-latex") or "").strip()
        if raw:
            return strip_latex_delimiters(html_module.unescape(raw))
        inner = (element.get_text() or "").strip()
        return strip_latex_delimiters(inner)

    def _clear_paragraph_text_runs(self, paragraph) -> None:
        p_el = paragraph._p
        for child in list(p_el):
            if child.tag == qn("w:r"):
                p_el.remove(child)

    def _mark_image_path_as_embedded(self, filename_or_path: str) -> None:
        raw = (filename_or_path or "").strip()
        if not raw:
            return
        raw = html_module.unescape(raw.replace("\\/", "/").strip())
        pth: Optional[Path] = None
        if "/api/v1/images/" in raw or raw.startswith("http"):
            pth = self._resolve_path_for_img_src(raw)
        if not pth:
            pth = self._resolve_image_path(raw) or self._resolve_image_path(Path(raw).name)
        if not pth:
            return
        try:
            self._embedded_resolved_paths.add(str(pth.resolve()))
        except OSError:
            self._embedded_resolved_paths.add(str(pth))
        self._embedded_basenames.add(Path(pth).name)

    def _try_insert_math_omml(self, paragraph, element: Tag) -> bool:
        classes = element.get("class") or []
        if isinstance(classes, str):
            classes = [classes]
        is_block = any(c in ("math-block", "math-display") for c in classes)
        inline = not is_block
        latex = self._extract_latex_from_math_element(element)
        if not latex or not should_try_omml(latex):
            return False
        cache_key = f"{'B' if is_block else 'I'}|{latex}"
        if cache_key in self._omml_fail_latex:
            return False
        el = latex_to_omml_element(latex, inline=inline)
        if el is None:
            self._omml_fail_latex.add(cache_key)
            return False
        try:
            if inline:
                run = paragraph.add_run()
                run._element.append(el)
            else:
                self._clear_paragraph_text_runs(paragraph)
                paragraph._p.append(el)
            data_img = (element.get("data-image") or "").strip()
            if data_img:
                self._mark_image_path_as_embedded(data_img)
            return True
        except Exception as e:
            logger.debug("写入 OMML 到段落失败: %s", e)
            self._omml_fail_latex.add(cache_key)
            return False

    def _preview_png_for_mathtype_ole(self, wmf_b: Optional[bytes]) -> bytes:
        if not wmf_b or len(wmf_b) <= 8:
            return minimal_preview_png_bytes()
        tmp_wmf = None
        tmp_png = None
        try:
            conv = WMFConverter()
            tw = tempfile.NamedTemporaryFile(suffix=".wmf", delete=False)
            tw.write(wmf_b)
            tw.flush()
            tw.close()
            tmp_wmf = tw.name
            tp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            tp.close()
            tmp_png = tp.name
            ok, _ = conv.convert(tmp_wmf, tmp_png)
            if ok and Path(tmp_png).is_file():
                png = Path(tmp_png).read_bytes()
                if png:
                    return png
        except Exception as e:
            logger.debug("WMF→PNG 预览失败，使用占位 PNG: %s", e)
        finally:
            for p in (tmp_wmf, tmp_png):
                if p:
                    try:
                        Path(p).unlink(missing_ok=True)
                    except OSError:
                        pass
        return minimal_preview_png_bytes()

    def _try_insert_mathtype_via_convert_equations(self, paragraph, element: Tag) -> bool:
        if get_convert_equations_exe() is None:
            return False
        latex = self._extract_latex_from_math_element(element)
        if not latex:
            return False
        classes = element.get("class") or []
        if isinstance(classes, str):
            classes = [classes]
        is_block = any(c in ("math-block", "math-display") for c in classes)
        if latex not in self._convert_equations_payload_cache:
            data, err = run_latex_to_mathtype_payload(latex)
            self._convert_equations_payload_cache[latex] = data
            if not data:
                logger.debug("ConvertEquations 未返回数据: %s", err)
        else:
            data = self._convert_equations_payload_cache[latex]
        if not data:
            return False
        ole_b, wmf_b = decode_math_type_model(data)
        if not ole_b:
            if wmf_b:
                logger.debug("ConvertEquations 仅返回 WMF（%d 字节），无 OLE，走 OMML/图", len(wmf_b))
            return False
        prog_id = (os.environ.get("EXPORT_MATHTYPE_OLE_PROGID") or "").strip() or None
        png_b = self._preview_png_for_mathtype_ole(wmf_b)
        ok = embed_mathtype_ole_in_paragraph(
            paragraph,
            ole_b,
            png_b,
            prog_id=prog_id,
            is_block=is_block,
        )
        if ok:
            data_img = (element.get("data-image") or "").strip()
            if data_img:
                self._mark_image_path_as_embedded(data_img)
            logger.debug("已内嵌 MathType OLE（%d 字节）", len(ole_b))
            return True
        return False

    def _fill_stem_root(self, para, root: Tag) -> None:
        """题干根节点：顶层多个 <p> / 块级 <div> 分段落，题号与首段正文同一段。"""
        first_top = True
        for child in root.children:
            if isinstance(child, NavigableString):
                s = str(child)
                if s.strip():
                    self._add_formatted_text(para, s)
                    first_top = False
                continue
            if not isinstance(child, Tag):
                continue
            if child.name == 'p':
                target = para if first_top else self.document.add_paragraph()
                self._fill_inline(target, child)
                first_top = False
            elif child.name == 'div':
                cl = child.get('class') or []
                if isinstance(cl, str):
                    cl = [cl]
                is_center = any(
                    x in ('math-block', 'math-display', 'formula-image-block') for x in cl
                )
                # 块级居中公式单独成段，避免把题号所在段整体居中
                if is_center:
                    target = self.document.add_paragraph()
                    target.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif first_top:
                    target = para
                else:
                    target = self.document.add_paragraph()
                self._fill_inline(target, child)
                first_top = False
            else:
                self._fill_inline(para, child)
                first_top = False
        try:
            para.paragraph_format.space_after = Pt(4)
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        except Exception:
            pass

    def _emit_img_tag(self, paragraph, element: Tag) -> None:
        """嵌入单个 <img>（含作为遍历根节点的 img：无子节点，必须在 _fill_inline 入口单独处理）。"""
        src = element.get('src') or ''
        cls = element.get('class') or []
        if isinstance(cls, str):
            cls = [cls]
        fn = self._src_to_filename(src)
        if fn or src:
            pth = self._resolve_path_for_img_src(src) or (
                self._resolve_image_path(fn) if fn else None
            )
            if pth:
                run = paragraph.add_run()
                self._add_picture_to_run(run, pth, cls)
                sk = self._img_src_dedupe_key(src)
                if sk:
                    self._embedded_img_src_keys.add(sk)
            else:
                logger.warning(
                    "导出找不到图片文件: src=%r -> fn=%r image_dir=%s lib=%s",
                    src[:200] if src else "",
                    fn,
                    self.image_dir,
                    self.images_library_root,
                )
        else:
            logger.warning("导出无法从 img 解析 src: %r", (element.get("src") or "")[:200])

    def _fill_inline(self, paragraph, element: Tag) -> None:
        """将 HTML 片段写入段落：文本 + 行内公式图（按子节点顺序，与 content_html 一致）。"""
        # <img> 无子节点；若仅作为递归根传入，上面 for 不会执行，导致图片丢失、顺序后移
        if element.name == "img":
            self._emit_img_tag(paragraph, element)
            return
        # math-inline/math-block：可选 ConvertEquations(OLE) → OMML → data-image → 斜体文本
        if element.name in ("span", "div"):
            classes = element.get("class") or []
            if isinstance(classes, str):
                classes = [classes]
            is_math_node = any(c in ("math-inline", "math-block", "math-display") for c in classes)
            if is_math_node:
                if self._try_insert_mathtype_via_convert_equations(paragraph, element):
                    return
                if self._try_insert_math_omml(paragraph, element):
                    return
                data_image = (element.get("data-image") or "").strip()
                if data_image:
                    pth = self._resolve_image_path(data_image)
                    if pth:
                        run = paragraph.add_run()
                        img_classes = ["formula-image-block"] if "math-block" in classes or "math-display" in classes else ["formula-image"]
                        self._add_picture_to_run(run, pth, img_classes)
                        return
        for child in element.children:
            if isinstance(child, NavigableString):
                s = str(child)
                if s:
                    self._add_formatted_text(paragraph, s)
                continue
            if not isinstance(child, Tag):
                continue
            if child.name == 'img':
                self._emit_img_tag(paragraph, child)
                continue
            if child.name == 'br':
                paragraph.add_run().add_break()
                continue
            if child.name == 'div':
                cl = child.get('class') or []
                if isinstance(cl, str):
                    cl = [cl]
                is_block = any(
                    x in ('math-block', 'math-display', 'formula-image-block') for x in cl
                )
                # 已在居中段内嵌套块级 div 时不再新开段落，减少同图分段、叠盖
                if is_block and paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    self._fill_inline(paragraph, child)
                else:
                    np = self.document.add_paragraph()
                    if is_block:
                        np.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self._fill_inline(np, child)
                continue
            if child.name in ('span', 'strong', 'b', 'em', 'i', 'u', 'sub', 'sup'):
                self._fill_inline(paragraph, child)
                continue
            # 其它标签（如 p、font）：展开子节点
            self._fill_inline(paragraph, child)

    @staticmethod
    def _escape_html_body(text: str) -> str:
        """与 splitter 题干 HTML 一致：正文内 & < >。"""
        t = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return t

    def _synthetic_math_inline_span(self, latex: str) -> Optional[Tag]:
        """为纯文本中的 $...$ 构造与 content_html 同结构的 math-inline，供 OLE/OMML 插入。"""
        raw = (latex or "").strip()
        if not raw:
            return None
        attr = html_module.escape(raw, quote=True)
        body = self._escape_html_body(raw)
        frag = f'<span class="math-inline" data-latex="{attr}">${body}$</span>'
        soup = BeautifulSoup(frag, "html.parser")
        return soup.select_one("span.math-inline")

    def _append_latex_inline_best_effort(self, paragraph, latex: str, font_size_pt: float) -> None:
        """MathType OLE → OMML → 斜体（与 _fill_inline 中行内公式策略一致，但不递归解析子节点）。"""
        span = self._synthetic_math_inline_span(latex)
        if span is None:
            return
        if self._try_insert_mathtype_via_convert_equations(paragraph, span):
            return
        if self._try_insert_math_omml(paragraph, span):
            return
        run = paragraph.add_run(f" ${latex.strip()} ")
        run.italic = True
        run.font.size = Pt(font_size_pt)

    def _add_formatted_text(self, paragraph, text: str, font_size_pt: float = 12):
        """普通文本 + `$...$` 行内公式；公式走 MathType/OMML，与题干同源。"""
        parts = (text or "").split("$")
        for i, part in enumerate(parts):
            if i % 2 == 0:
                if part:
                    run = paragraph.add_run(part)
                    run.font.size = Pt(font_size_pt)
            else:
                self._append_latex_inline_best_effort(paragraph, part, font_size_pt)
                
    def _add_options(self, options: List[Any]):
        """添加选项"""
        # 每行两个选项
        for i in range(0, len(options), 2):
            para = self.document.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.5)
            
            # 第一个选项
            opt1 = options[i]
            run = para.add_run(f"{opt1.label}. ")
            run.font.bold = True
            self._add_option_body(para, opt1)
            
            # 第二个选项（如果有）
            if i + 1 < len(options):
                opt2 = options[i + 1]
                # 添加制表符分隔
                para.add_run('\t\t')
                run = para.add_run(f"{opt2.label}. ")
                run.font.bold = True
                self._add_option_body(para, opt2)

    def _fill_option_root(self, paragraph, root: Tag) -> None:
        """选项根：顶层多段与题干一致分段。"""
        first_top = True
        for child in root.children:
            if isinstance(child, NavigableString):
                s = str(child)
                if s.strip():
                    self._add_formatted_text(paragraph, s)
                    first_top = False
                continue
            if not isinstance(child, Tag):
                continue
            if child.name == 'p':
                target = paragraph if first_top else self.document.add_paragraph()
                target.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
                self._fill_inline(target, child)
                first_top = False
            elif child.name == 'div':
                cl = child.get('class') or []
                if isinstance(cl, str):
                    cl = [cl]
                is_center = any(
                    x in ('math-block', 'math-display', 'formula-image-block') for x in cl
                )
                if is_center:
                    target = self.document.add_paragraph()
                    target.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
                    target.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif first_top:
                    target = paragraph
                else:
                    target = self.document.add_paragraph()
                    target.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
                self._fill_inline(target, child)
                first_top = False
            else:
                self._fill_inline(paragraph, child)
                first_top = False

    def _add_option_body(self, paragraph, opt: Any) -> None:
        oh = (getattr(opt, 'content_html', None) or '').strip()
        if oh:
            oh = html_module.unescape(oh)
            oh = oh.replace('\\"', '"').replace("\\'", "'").replace("\\/", "/")
            wrapped = BeautifulSoup(f'<div class="opt-root">{oh}</div>', 'html.parser')
            root = wrapped.select_one('.opt-root')
            if root:
                for span in root.select('span.option-label'):
                    span.decompose()
                self._fill_option_root(paragraph, root)
                if not re.search(r"<img\b", oh, re.I):
                    self._add_missing_images_inline(paragraph, getattr(opt, 'images', None) or [])
        else:
            self._add_formatted_text(paragraph, opt.content)
                
    def add_section_header(self, text: str):
        """添加章节标题（如"一、选择题"）"""
        para = self.document.add_paragraph()
        run = para.add_run(text)
        run.font.bold = True
        run.font.size = Pt(14)
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        para.paragraph_format.space_before = Pt(12)
        para.paragraph_format.space_after = Pt(6)
        
    def save(self, output_path: str):
        """保存文档"""
        self._refresh_document_modified_time()
        self.document.save(output_path)

    def to_bytes(self) -> bytes:
        """导出为字节流"""
        self._refresh_document_modified_time()
        buffer = io.BytesIO()
        self.document.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()


def export_questions(questions: List[Question], output_path: str, 
                     title: str = "", watermark: str = None,
                     image_dir: Optional[str] = None,
                     include_answer: bool = False, 
                     include_analysis: bool = False) -> str:
    """
    便捷函数：导出题目到 Word
    
    Args:
        questions: 题目列表
        output_path: 输出路径
        title: 文档标题
        watermark: 水印文字
        include_answer: 是否包含答案
        include_analysis: 是否包含解析
        
    Returns:
        输出文件路径
    """
    exporter = WordExporter()
    exporter.create_document(title=title, watermark_text=watermark, image_dir=image_dir)
    
    # 按题型分组
    current_type = None
    for question in questions:
        # 添加题型标题
        if question.type_name and question.type_name != current_type:
            exporter.add_section_header(f"{question.type_name}")
            current_type = question.type_name
            
        exporter.add_question(question, include_answer, include_analysis)
        
    exporter.save(output_path)
    return output_path
