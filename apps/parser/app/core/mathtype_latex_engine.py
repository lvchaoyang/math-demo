"""
MathType → LaTeX 可插拔引擎

- 无官方 Python MTEF→LaTeX 库时，用环境变量接入自建/第三方可执行程序（如基于 MathType SDK 的 CLI）。
- 默认 auto：已配置外部命令则优先外部，否则启发式扫描 OLE 流。
"""
from __future__ import annotations

import os
import re
import shlex
import shutil
import subprocess
import logging
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)


def get_latex_mode() -> str:
    """MATHTYPE_LATEX_MODE: auto | heuristic | external | none"""
    raw = (os.environ.get("MATHTYPE_LATEX_MODE") or "auto").strip().lower()
    if raw in ("off", "disabled", "0", "false"):
        return "none"
    if raw in ("auto", "heuristic", "external", "none"):
        return raw
    logger.warning("未知 MATHTYPE_LATEX_MODE=%s，回退 auto", raw)
    return "auto"


def is_latex_extraction_disabled() -> bool:
    return get_latex_mode() == "none"


def should_try_external_first() -> bool:
    mode = get_latex_mode()
    return mode in ("auto", "external")


def should_fallback_heuristic() -> bool:
    mode = get_latex_mode()
    if mode == "external":
        return False
    if mode == "heuristic":
        return True
    if mode == "auto":
        return True
    return False


def _get_external_timeout_sec() -> int:
    try:
        return max(1, min(120, int(os.environ.get("MATHTYPE_LATEX_TIMEOUT", "15"))))
    except ValueError:
        return 15


def _build_external_argv(cmd_line: str, ole_path: str) -> Optional[List[str]]:
    """
    MATHTYPE_LATEX_CMD 示例（整条命令，含参数）：
      - mytool --format tex "{ole_path}"
      - C:\\Tools\\mtef2tex.exe
    未写 {ole_path} 时，默认在末尾追加 ole_path。
    Windows 下使用非 posix 的 shlex 以便处理反斜杠路径。
    """
    cmd_line = (cmd_line or "").strip()
    if not cmd_line:
        return None
    posix = os.name != "nt"
    try:
        parts = shlex.split(cmd_line, posix=posix)
    except ValueError as e:
        logger.warning("MATHTYPE_LATEX_CMD 解析失败: %s", e)
        return None
    if not parts:
        return None
    expanded: List[str] = []
    placeholder = "{ole_path}"
    for p in parts:
        if placeholder in p:
            expanded.append(p.replace(placeholder, ole_path))
        else:
            expanded.append(p)
    if not any(placeholder in x for x in parts) and ole_path not in expanded:
        expanded.append(ole_path)
    return expanded


def try_external_latex(ole_path: str) -> Tuple[Optional[str], str]:
    """
    调用外部程序，从 stdout 读取 LaTeX（去首尾空白）。
    返回 (latex 或 None, 状态码)。
    """
    cmd_line = (os.environ.get("MATHTYPE_LATEX_CMD") or "").strip()
    if not cmd_line:
        return None, "external_not_configured"

    exe0 = None
    try:
        argv = _build_external_argv(cmd_line, ole_path)
        if not argv:
            return None, "external_bad_cmd"
        exe0 = argv[0]
        if not shutil.which(exe0) and not os.path.isfile(exe0):
            return None, "external_exe_not_found"

        timeout = _get_external_timeout_sec()
        env = os.environ.copy()
        env.setdefault("OLE_PATH", ole_path)

        proc = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding="utf-8",
            errors="replace",
        )
        if proc.returncode != 0:
            err = (proc.stderr or proc.stdout or "").strip()[:500]
            logger.debug("MATHTYPE_LATEX_CMD 非零退出 %s: %s", proc.returncode, err)
            return None, "external_nonzero_exit"

        out = (proc.stdout or "").strip()
        if not out:
            return None, "external_empty_stdout"

        # 单行或多行，去掉 BOM
        if out.startswith("\ufeff"):
            out = out[1:].strip()
        # 若输出被日志包裹，尝试取最后一行看起来像 TeX 的
        latex = _pick_latex_from_output(out)
        if not latex:
            return None, "external_no_latex_in_output"
        return latex, "ok_external"
    except subprocess.TimeoutExpired:
        return None, "external_timeout"
    except FileNotFoundError:
        return None, "external_exe_not_found"
    except Exception as e:
        logger.warning("外部 LaTeX 转换异常: %s", e)
        return None, "external_exception"


def _pick_latex_from_output(text: str) -> Optional[str]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        return None
    # 优先：整段已是 TeX
    if _looks_tex_like(text.strip()):
        return text.strip()
    # 否则从最后一行往前找
    for line in reversed(lines):
        if _looks_tex_like(line):
            return line
    # 最后一行兜底
    return lines[-1] if lines else None


def _looks_tex_like(s: str) -> bool:
    if len(s) < 2:
        return False
    score = 0
    if "\\" in s:
        score += 2
    if re.search(r"\\[a-zA-Z]+", s):
        score += 2
    if "{" in s and "}" in s:
        score += 1
    if any(c in s for c in "^_"):
        score += 1
    return score >= 3
