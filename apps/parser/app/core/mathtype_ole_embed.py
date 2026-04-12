"""将 MathType OLE 字节嵌入 WordprocessingML（w:object + word/embeddings/*.bin）。"""

from __future__ import annotations

import hashlib
import io
import logging
import re
import uuid

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

_DEFAULT_MATHTYPE_PROGID = "Equation.DSMT4"


def minimal_preview_png_bytes() -> bytes:
    from PIL import Image as PILImage

    buf = io.BytesIO()
    PILImage.new("RGB", (8, 8), (255, 255, 255)).save(buf, format="PNG")
    return buf.getvalue()


def _png_size_pt(png_bytes: bytes) -> tuple[float, float]:
    from PIL import Image as PILImage

    with PILImage.open(io.BytesIO(png_bytes)) as im:
        w, h = im.size
    return (max(w * 72.0 / 96.0, 12.0), max(h * 72.0 / 96.0, 12.0))


def unique_preview_png_for_ole(
    preview_png_bytes: bytes, ole_bytes: bytes, *, nonce: bytes
) -> bytes:
    """
    python-docx 的 get_or_add_image_part 按 PNG 二进制去重。
    若多道 MathType OLE 共用同一份预览图（如均为占位 PNG、WMF→PNG 相同、或 LaTeX 不同但 OLE 字节相同），
    Word 里会多处显示同一张预览，表现为「顺序对但式子重复/错位」。
    nonce 须每次嵌入唯一（如 uuid），使每道 OLE 独占图片部件，避免跨题/同题重复式串图。
    """
    from PIL import Image as PILImage

    tag = hashlib.sha256(ole_bytes + nonce).digest()[:3]
    r, g, b = tag[0], tag[1], tag[2]
    try:
        im = PILImage.open(io.BytesIO(preview_png_bytes)).convert("RGBA")
        w, h = im.size
        if w < 1 or h < 1:
            raise ValueError("empty png")
        px = im.load()
        px[w - 1, h - 1] = (r, g, b, 255)
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        buf = io.BytesIO()
        PILImage.new("RGB", (16, 16), (r, g, b)).save(buf, format="PNG")
        return buf.getvalue()


def embed_mathtype_ole_in_paragraph(
    paragraph,
    ole_bytes: bytes,
    preview_png_bytes: bytes,
    *,
    prog_id: str | None = None,
    is_block: bool = False,
) -> bool:
    if not ole_bytes:
        return False
    png_b = preview_png_bytes or minimal_preview_png_bytes()
    png_b = unique_preview_png_for_ole(png_b, ole_bytes, nonce=uuid.uuid4().bytes)
    pid = (prog_id or "").strip() or _DEFAULT_MATHTYPE_PROGID

    doc_part = paragraph.part
    pkg = doc_part.package

    try:
        partname = pkg.next_partname("/word/embeddings/embedding%d.bin")
        ole_part = Part(partname, CT.OFC_OLE_OBJECT, ole_bytes, package=pkg)
        r_id_ole = doc_part.relate_to(ole_part, RT.OLE_OBJECT)

        png_stream = io.BytesIO(png_b)
        png_stream.seek(0)
        image_part = pkg.get_or_add_image_part(png_stream)
        r_id_img = doc_part.relate_to(image_part, RT.IMAGE)

        w_pt, h_pt = _png_size_pt(png_b)
        dxa = int(w_pt * 20)
        dya = int(h_pt * 20)
        shape_id = "_x0000_i" + uuid.uuid4().hex[:10]
        object_id = "_" + uuid.uuid4().hex[:16]
        safe_pid = re.sub(r'[<>&"]', "", pid)

        xml = f'''<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:object w:dxaOrig="{dxa}" w:dyaOrig="{dya}">
    <v:shape id="{shape_id}" o:ole="" type="#_x0000_t75" style="width:{w_pt:.2f}pt;height:{h_pt:.2f}pt">
      <v:imagedata r:id="{r_id_img}" o:title=""/>
    </v:shape>
    <o:OLEObject Type="Embed" ProgID="{safe_pid}" ShapeID="{shape_id}" DrawAspect="Content" ObjectID="{object_id}" r:id="{r_id_ole}"/>
  </w:object>
</w:r>'''
        r_el = parse_xml(xml)

        p_el = paragraph._p
        if is_block:
            for child in list(p_el):
                if child.tag == qn("w:r"):
                    p_el.remove(child)
        p_el.append(r_el)
        return True
    except Exception as e:
        logger.warning("嵌入 MathType OLE 失败: %s", e)
        return False
