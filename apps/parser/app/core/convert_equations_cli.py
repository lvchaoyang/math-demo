"""
可选：调用 tools/latexToMathType 编译出的 ConvertEquations.exe（--latex），
从 LaTeX 经 Word/MathType 生成 OLE/WMF（JSON 输出）。

需本机安装 MathType + Word；InvalidLatex 等对 LaTeX 限制见 C# 端。
"""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import subprocess
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)

_ENV_EXE = "EXPORT_CONVERT_EQUATIONS_EXE"
_DEFAULT_TIMEOUT = 120


def get_convert_equations_exe() -> Optional[Path]:
    raw = (os.environ.get(_ENV_EXE) or "").strip().strip('"')
    if not raw:
        return None
    p = Path(raw)
    if p.is_file():
        return p
    logger.warning("%s 指向的文件不存在: %s", _ENV_EXE, raw)
    return None


def _parse_json_stdout(stdout: str) -> Optional[Dict[str, Any]]:
    s = (stdout or "").strip()
    if not s:
        return None
    line = s.splitlines()[0].strip()
    try:
        return json.loads(line)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", s)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    return None


def run_latex_to_mathtype_payload(
    latex: str, timeout: Optional[int] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """执行 ConvertEquations --latex <latex>，解析 stdout 为 JSON。返回 (payload, error_message)。"""
    exe = get_convert_equations_exe()
    if not exe:
        return None, "exe_not_configured"
    latex = (latex or "").strip()
    if not latex:
        return None, "empty_latex"
    t = timeout if timeout is not None else int(
        os.environ.get("EXPORT_CONVERT_EQUATIONS_TIMEOUT", str(_DEFAULT_TIMEOUT))
    )
    t = max(5, min(300, t))
    try:
        run_kw: Dict[str, Any] = {
            "capture_output": True,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "timeout": t,
        }
        if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            run_kw["creationflags"] = subprocess.CREATE_NO_WINDOW
        proc = subprocess.run([str(exe), "--latex", latex], **run_kw)
    except subprocess.TimeoutExpired:
        return None, "timeout"
    except OSError as e:
        logger.warning("ConvertEquations 启动失败: %s", e)
        return None, str(e)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()[:500]
        logger.debug("ConvertEquations 退出码 %s: %s", proc.returncode, err)
        return None, f"exit_{proc.returncode}"
    data = _parse_json_stdout(proc.stdout or "")
    if not data:
        logger.debug("ConvertEquations 无有效 JSON 输出: %s", (proc.stdout or "")[:300])
        return None, "no_json"
    return data, None


def decode_math_type_model(data: Dict[str, Any]) -> Tuple[Optional[bytes], Optional[bytes]]:
    """从 MathTypeModel JSON 解码 ole / wmf 字节（字段多为 base64 字符串）。"""
    ole_b: Optional[bytes] = None
    wmf_b: Optional[bytes] = None
    ole_s = data.get("ole")
    wmf_s = data.get("wmf")
    if isinstance(ole_s, str) and ole_s.strip():
        try:
            ole_b = base64.b64decode(ole_s)
        except Exception as e:
            logger.debug("解码 ole base64 失败: %s", e)
    if isinstance(wmf_s, str) and wmf_s.strip():
        try:
            wmf_b = base64.b64decode(wmf_s)
        except Exception as e:
            logger.debug("解码 wmf base64 失败: %s", e)
    return ole_b, wmf_b
