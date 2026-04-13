"""
可选：在单次 HTTP 请求结束后清理本机残留子进程（主要为 ConvertEquations / MathType 链路透出的 WINWORD.EXE）。

默认不执行任何操作；设置环境变量后才运行，避免误杀用户正在编辑的 Word。
"""

from __future__ import annotations

import logging
import os
import subprocess

logger = logging.getLogger(__name__)


def _should_run_cleanup(phase: str) -> bool:
    """
    EXPORT_KILL_WINWORD_AFTER:
      - 未设置 / 空 / 0 / false / off：不清理
      - 1 / true / yes / all / both：parse / export / convert 后都尝试
      - parse：仅 /parse、/parse/v2 完成后
      - export：仅 /export 完成后
      - convert：仅 LaTeX->MathType（ConvertEquations）完成后
    """
    raw = (os.environ.get("EXPORT_KILL_WINWORD_AFTER") or "").strip().lower()
    if not raw or raw in ("0", "false", "no", "off"):
        return False
    if raw in ("1", "true", "yes", "all", "both"):
        return True
    if raw == "parse":
        return phase == "parse"
    if raw == "export":
        return phase == "export"
    if raw == "convert":
        return phase == "convert"
    return False


def optional_kill_winword_after_request(phase: str) -> None:
    """
    Windows：taskkill /IM WINWORD.EXE /F，用于释放 ConvertEquations 可能遗留的 Word 实例。

    会结束本机所有 Word 窗口；请勿在需要保留 Word 编辑时开启。
    """
    if os.name != "nt":
        return
    if not _should_run_cleanup(phase):
        return
    try:
        r = subprocess.run(
            ["taskkill", "/IM", "WINWORD.EXE", "/F"],
            capture_output=True,
            text=True,
            timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if r.returncode == 0:
            logger.info("EXPORT_KILL_WINWORD_AFTER：已结束 WINWORD.EXE（phase=%s）", phase)
        else:
            err = (r.stderr or r.stdout or "").strip()[:300]
            logger.debug(
                "EXPORT_KILL_WINWORD_AFTER：taskkill 未成功（可能无残留 Word）phase=%s err=%s",
                phase,
                err,
            )
    except subprocess.TimeoutExpired:
        logger.warning("EXPORT_KILL_WINWORD_AFTER：taskkill 超时（phase=%s）", phase)
    except OSError as e:
        logger.debug("EXPORT_KILL_WINWORD_AFTER：taskkill 失败 %s", e)
