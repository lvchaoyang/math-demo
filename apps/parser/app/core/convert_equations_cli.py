"""
可选：调用 tools/latexToMathType 编译出的 ConvertEquations.exe（--latex），
从 LaTeX 经 Word/MathType 生成 OLE/WMF（JSON 输出）。

未设置 EXPORT_CONVERT_EQUATIONS_EXE 时，会按仓库相对路径自动查找 Release/Debug 下的 exe。

需本机安装 MathType + Word；InvalidLatex 等对 LaTeX 限制见 C# 端。
"""

from __future__ import annotations

import base64
import json
import logging
import os
import shutil
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_ENV_EXE = "EXPORT_CONVERT_EQUATIONS_EXE"
_DEFAULT_TIMEOUT = 120

# 无 EXPORT_CONVERT_EQUATIONS_EXE 时，自动发现结果只计算一次
_auto_exe_resolved: Optional[Path] = None
_auto_exe_scan_done: bool = False


def _core_dir() -> Path:
    return Path(__file__).resolve().parent


def _iter_repo_roots() -> List[Path]:
    """从 apps/parser/app/core 向上尝试定位 monorepo 根目录。"""
    p = _core_dir()
    roots: List[Path] = []
    for depth in (4, 5, 3, 6):
        if len(p.parents) > depth:
            roots.append(p.parents[depth])
    # 去重保序
    seen = set()
    out: List[Path] = []
    for r in roots:
        k = str(r)
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def _built_convert_equations_paths() -> List[Path]:
    rels = (
        "tools/latexToMathType/ConvertEquations/ConvertEquations/bin/Release/ConvertEquations.exe",
        "tools/latexToMathType/ConvertEquations/ConvertEquations/bin/Debug/ConvertEquations.exe",
    )
    paths: List[Path] = []
    for root in _iter_repo_roots():
        for rel in rels:
            paths.append((root / rel).resolve())
    return paths


def get_convert_equations_exe() -> Optional[Path]:
    """
    解析 ConvertEquations.exe 路径：
    1) EXPORT_CONVERT_EQUATIONS_EXE（显式配置优先）
    2) 仓库 tools/.../bin/Release|Debug/ConvertEquations.exe
    3) PATH 中的 ConvertEquations / ConvertEquations.exe
    """
    global _auto_exe_resolved, _auto_exe_scan_done

    raw = (os.environ.get(_ENV_EXE) or "").strip().strip('"')
    if raw:
        p = Path(raw)
        if p.is_file():
            return p
        logger.warning("%s 指向的文件不存在: %s", _ENV_EXE, raw)
        return None

    if _auto_exe_scan_done:
        return _auto_exe_resolved

    _auto_exe_scan_done = True
    for candidate in _built_convert_equations_paths():
        if candidate.is_file():
            _auto_exe_resolved = candidate
            logger.info("ConvertEquations 自动发现: %s", candidate)
            return _auto_exe_resolved

    for name in ("ConvertEquations.exe", "ConvertEquations"):
        w = shutil.which(name)
        if w:
            pth = Path(w)
            if pth.is_file():
                _auto_exe_resolved = pth
                logger.info("ConvertEquations 从 PATH 发现: %s", pth)
                return _auto_exe_resolved

    return None


def _strip_bom(s: str) -> str:
    if s.startswith("\ufeff"):
        return s[1:]
    return s


def _balanced_json_objects(s: str) -> List[str]:
    """从文本中提取所有大括号平衡的 {...} 片段（应对多行或前后杂质）。"""
    chunks: List[str] = []
    depth = 0
    start = -1
    in_str = False
    esc = False
    for i, c in enumerate(s):
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
            continue
        if c == '"':
            in_str = True
            continue
        if c == "{":
            if depth == 0:
                start = i
            depth += 1
        elif c == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start >= 0:
                    chunks.append(s[start : i + 1])
                    start = -1
    return chunks


def _parse_json_stdout(stdout: str) -> Optional[Dict[str, Any]]:
    """
    解析 ConvertEquations 输出中的 JSON。
    C# 端可能向 stdout 写入调试行（如 mml:、异常信息），故不能只取首行。
    """
    s = _strip_bom((stdout or "").strip())
    if not s:
        return None

    # 1) 整段即 JSON
    if s.startswith("{"):
        try:
            return json.loads(s)
        except json.JSONDecodeError:
            pass

    # 2) 逐行：常见「最后一行是 JSON」
    lines = [ln.strip() for ln in s.splitlines() if ln.strip()]
    for line in reversed(lines):
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                continue

    # 3) 首行 JSON（旧行为兼容）
    if lines:
        try:
            return json.loads(lines[0])
        except json.JSONDecodeError:
            pass

    # 4) 平衡括号扫描，从后往前试（真正的 payload 往往在末尾）
    for chunk in reversed(_balanced_json_objects(s)):
        try:
            return json.loads(chunk)
        except json.JSONDecodeError:
            continue

    # 5) 宽松：第一个 { 到最后一个 }（旧正则的加强版）
    i0 = s.find("{")
    i1 = s.rfind("}")
    if i0 >= 0 and i1 > i0:
        try:
            return json.loads(s[i0 : i1 + 1])
        except json.JSONDecodeError:
            pass

    return None


def _get_str_ci(d: Dict[str, Any], name: str) -> str:
    for k, v in d.items():
        if isinstance(k, str) and k.lower() == name.lower() and isinstance(v, str):
            return v
    return ""


def decode_math_type_model(data: Dict[str, Any]) -> Tuple[Optional[bytes], Optional[bytes]]:
    """从 MathTypeModel JSON 解码 ole / wmf 字节（字段多为 base64 字符串）。"""
    ole_b: Optional[bytes] = None
    wmf_b: Optional[bytes] = None
    ole_s = _get_str_ci(data, "ole").strip()
    wmf_s = _get_str_ci(data, "wmf").strip()
    if ole_s:
        try:
            ole_b = base64.b64decode(ole_s)
        except Exception as e:
            logger.debug("解码 ole base64 失败: %s", e)
    if wmf_s:
        try:
            wmf_b = base64.b64decode(wmf_s)
        except Exception as e:
            logger.debug("解码 wmf base64 失败: %s", e)
    return ole_b, wmf_b


def run_latex_to_mathtype_payload(
    latex: str, timeout: Optional[int] = None
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """执行 ConvertEquations --latex <latex>，解析 stdout/stderr 中的 JSON。返回 (payload, error_message)。"""
    exe = get_convert_equations_exe()
    if exe is None:
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

    combined = (proc.stdout or "") + "\n" + (proc.stderr or "")
    data = _parse_json_stdout(proc.stdout or "")
    if not data:
        data = _parse_json_stdout(combined)
    if not data:
        logger.debug("ConvertEquations 无有效 JSON 输出: %s", (proc.stdout or "")[:400])
        return None, "no_json"
    return data, None
