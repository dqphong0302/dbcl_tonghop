#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hệ thống ghi log có cấu trúc cho phần mềm Tổng hợp Khảo sát.

Mục đích (phục vụ tiêu chí "Tính khoa học" của SKKN):
- Ghi lại toàn bộ quá trình xử lý ra file log để làm MINH CHỨNG VẬN HÀNH.
- Mỗi lần chạy tạo 1 file log riêng theo ngày, có timestamp từng dòng.
- Cho phép vừa ghi ra file, vừa đẩy ra callback của GUI (log realtime).

Cách dùng:
    from app_logging import get_logger
    log = get_logger()
    log.info("Bắt đầu xử lý")
    log.warning("Sheet 2 thiếu dữ liệu")
    log.error("Không đọc được file X")
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

# ───────────────────────────────────────────────────────────────
# Xác định thư mục lưu log (ghi cạnh app, fallback về home nếu read-only)
# ───────────────────────────────────────────────────────────────

APP_NAME = "dbcl_tonghop"


def _resolve_log_dir() -> Path:
    """Tìm thư mục ghi log phù hợp, an toàn cho cả bản đóng gói PyInstaller."""
    candidates = []

    # 1) Cạnh file thực thi / script
    try:
        if getattr(sys, "frozen", False):
            base = Path(sys.executable).resolve().parent
        else:
            base = Path(__file__).resolve().parent
        candidates.append(base / "logs")
    except Exception:
        pass

    # 2) Fallback: thư mục home người dùng
    candidates.append(Path.home() / f".{APP_NAME}" / "logs")

    for cand in candidates:
        try:
            cand.mkdir(parents=True, exist_ok=True)
            # Kiểm tra quyền ghi
            test = cand / ".write_test"
            test.write_text("ok", encoding="utf-8")
            test.unlink(missing_ok=True)
            return cand
        except Exception:
            continue

    # Cùng đường: dùng thư mục tạm
    import tempfile
    fallback = Path(tempfile.gettempdir()) / f"{APP_NAME}_logs"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


LOG_DIR = _resolve_log_dir()

# Tên file log theo ngày
_LOG_FILE = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d')}.log"

_LOGGER_NAME = "dbcl"
_INITIALIZED = False


class _GuiHandler(logging.Handler):
    """Handler đẩy bản ghi log tới callback của GUI (nếu có)."""

    def __init__(self, callback):
        super().__init__()
        self._callback = callback

    def emit(self, record):
        try:
            msg = record.getMessage()
            tag = {
                logging.INFO: "info",
                logging.WARNING: "warning",
                logging.ERROR: "error",
                logging.CRITICAL: "error",
                logging.DEBUG: "dim",
            }.get(record.levelno, None)
            self._callback(msg, tag)
        except Exception:
            pass


def get_logger() -> logging.Logger:
    """Trả về logger dùng chung, khởi tạo file handler 1 lần."""
    global _INITIALIZED
    logger = logging.getLogger(_LOGGER_NAME)
    if _INITIALIZED:
        return logger

    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler (UTF-8)
    try:
        fh = logging.FileHandler(_LOG_FILE, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass

    # Console handler
    try:
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.INFO)
        ch.setFormatter(fmt)
        logger.addHandler(ch)
    except Exception:
        pass

    _INITIALIZED = True
    logger.info("=" * 60)
    logger.info("Khởi động phiên làm việc mới - log: %s", _LOG_FILE)
    return logger


def attach_gui_callback(callback):
    """Gắn callback GUI để log hiển thị realtime trên giao diện.

    callback(msg: str, tag: str|None) -> None
    """
    logger = get_logger()
    # Tránh gắn trùng
    for h in logger.handlers:
        if isinstance(h, _GuiHandler):
            logger.removeHandler(h)
    handler = _GuiHandler(callback)
    handler.setLevel(logging.INFO)
    logger.addHandler(handler)
    return handler


def detach_gui_handler(handler):
    """Gỡ GUI handler khi đóng cửa sổ / chuyển tab."""
    try:
        get_logger().removeHandler(handler)
    except Exception:
        pass


def get_log_file_path() -> str:
    """Đường dẫn file log hiện tại (để hiển thị/mở từ GUI)."""
    return str(_LOG_FILE)


def get_log_dir() -> str:
    return str(LOG_DIR)
