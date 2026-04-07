"""LaTeX → MathML → OMML（Word 原生公式），用于导出器优先插入可编辑公式。"""

from __future__ import annotations

import logging

from docx.oxml import parse_xml

logger = logging.getLogger(__name__)


def strip_latex_delimiters(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return ""
    if s.startswith("$$") and s.endswith("$$") and len(s) >= 4:
        return s[2:-2].strip()
    if s.startswith("$") and s.endswith("$") and len(s) >= 2:
        return s[1:-1].strip()
    return s


def should_try_omml(latex: str) -> bool:
    """过简或过长的 LaTeX 不尝试 OMML，避免无谓失败。"""
    t = (latex or "").strip()
    if len(t) < 1 or len(t) > 8000:
        return False
    return True


def _ensure_omml_xmlns(omml: str) -> str:
    """mathml2omml 常输出带 m: 前缀但无 xmlns 声明的片段，lxml/parse_xml 会报未定义前缀。"""
    s = (omml or "").strip()
    if not s or "xmlns:m=" in s[:300]:
        return s
    end = s.find(">")
    if end <= 0 or not s.startswith("<m:"):
        return s
    ns = ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    return s[:end] + ns + s[end:]


def latex_to_omml_element(latex: str, inline: bool = True):
    """
    返回可挂到 run._element 或 paragraph._p 的 oMath/oMathPara 元素；失败返回 None。
    """
    try:
        from latex2mathml.converter import convert as latex_to_mathml
        from mathml2omml import convert as mathml_to_omml
    except ImportError as e:
        logger.debug("latex2mathml/mathml2omml 未安装: %s", e)
        return None

    latex = (latex or "").strip()
    if not latex:
        return None
    try:
        mathml = latex_to_mathml(latex)
        omml = mathml_to_omml(mathml)
        if not omml or not omml.strip():
            return None
        omml = _ensure_omml_xmlns(omml)
        el = parse_xml(omml)
        tag = el.tag
        if inline and tag.endswith("oMathPara"):
            for child in list(el):
                if child.tag.endswith("oMath"):
                    return child
        if not inline and tag.endswith("oMath"):
            wrap = (
                '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                "</m:oMathPara>"
            )
            para = parse_xml(wrap)
            para.append(el)
            return para
        return el
    except Exception as e:
        logger.debug("LaTeX→OMML 失败: %s", e)
        return None
