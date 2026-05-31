#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Đo lường HIỆU QUẢ phần mềm (Round 2 - tiêu chí then chốt của SKKN).

Sinh dữ liệu "TRƯỚC / SAU" để chứng minh hiệu quả:
  - Đo thời gian xử lý thực tế của phần mềm (giây/báo cáo).
  - Ước lượng thời gian & công sức làm THỦ CÔNG tương đương.
  - Tính thời gian tiết kiệm, tỷ lệ giảm, năng suất tăng.

Các tham số làm tay (MANUAL_*) là ƯỚC LƯỢNG mặc định, có thể chỉnh lại
bằng số liệu thực tế của đơn vị (xem hàm set_manual_estimates).
"""

import json
import os
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional


# ───────────────────────────────────────────────────────────────
# THAM SỐ ƯỚC LƯỢNG QUY TRÌNH THỦ CÔNG (có thể chỉnh theo thực tế)
# ───────────────────────────────────────────────────────────────

# Thời gian trung bình làm THỦ CÔNG cho 1 báo cáo học phần (phút)
MANUAL_MINUTES_PER_REPORT = 45.0
# Thời gian trung bình tổng hợp & làm sạch ý kiến thủ công cho 1 file (phút)
MANUAL_MINUTES_PER_FEEDBACK_FILE = 30.0
# Tỷ lệ sai sót khi nhập/tính tay (tham khảo, dùng để thuyết minh)
MANUAL_ERROR_RATE = 0.05

_ESTIMATE_FILE = Path.home() / ".dbcl_tonghop_benchmark.json"


def set_manual_estimates(minutes_per_report: Optional[float] = None,
                         minutes_per_feedback: Optional[float] = None):
    """Cập nhật tham số ước lượng làm tay (lưu lại để dùng về sau)."""
    global MANUAL_MINUTES_PER_REPORT, MANUAL_MINUTES_PER_FEEDBACK_FILE
    if minutes_per_report is not None:
        MANUAL_MINUTES_PER_REPORT = float(minutes_per_report)
    if minutes_per_feedback is not None:
        MANUAL_MINUTES_PER_FEEDBACK_FILE = float(minutes_per_feedback)
    try:
        _ESTIMATE_FILE.write_text(json.dumps({
            "minutes_per_report": MANUAL_MINUTES_PER_REPORT,
            "minutes_per_feedback": MANUAL_MINUTES_PER_FEEDBACK_FILE,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def load_manual_estimates():
    """Nạp tham số ước lượng đã lưu (nếu có)."""
    global MANUAL_MINUTES_PER_REPORT, MANUAL_MINUTES_PER_FEEDBACK_FILE
    try:
        if _ESTIMATE_FILE.exists():
            data = json.loads(_ESTIMATE_FILE.read_text(encoding="utf-8"))
            MANUAL_MINUTES_PER_REPORT = float(
                data.get("minutes_per_report", MANUAL_MINUTES_PER_REPORT))
            MANUAL_MINUTES_PER_FEEDBACK_FILE = float(
                data.get("minutes_per_feedback", MANUAL_MINUTES_PER_FEEDBACK_FILE))
    except Exception:
        pass


# ───────────────────────────────────────────────────────────────
# BỘ ĐẾM THỜI GIAN
# ───────────────────────────────────────────────────────────────

@dataclass
class BenchmarkResult:
    n_reports: int = 0              # số báo cáo Word tạo ra
    n_feedback_files: int = 0       # số file feedback đã xử lý
    n_feedbacks: int = 0            # số ý kiến đã làm sạch
    auto_seconds: float = 0.0       # thời gian phần mềm chạy thực tế

    @property
    def manual_minutes(self) -> float:
        return (self.n_reports * MANUAL_MINUTES_PER_REPORT
                + self.n_feedback_files * MANUAL_MINUTES_PER_FEEDBACK_FILE)

    @property
    def auto_minutes(self) -> float:
        return self.auto_seconds / 60.0

    @property
    def saved_minutes(self) -> float:
        return max(0.0, self.manual_minutes - self.auto_minutes)

    @property
    def reduction_pct(self) -> float:
        if self.manual_minutes <= 0:
            return 0.0
        return self.saved_minutes / self.manual_minutes * 100

    @property
    def speedup(self) -> float:
        if self.auto_minutes <= 0:
            return 0.0
        return self.manual_minutes / self.auto_minutes

    def to_dict(self) -> Dict:
        return {
            **asdict(self),
            "manual_minutes": round(self.manual_minutes, 1),
            "auto_minutes": round(self.auto_minutes, 2),
            "saved_minutes": round(self.saved_minutes, 1),
            "saved_hours": round(self.saved_minutes / 60.0, 2),
            "reduction_pct": round(self.reduction_pct, 1),
            "speedup": round(self.speedup, 1),
        }

    def summary_lines(self) -> List[str]:
        d = self.to_dict()
        return [
            f"Số báo cáo tạo tự động: {self.n_reports}",
            f"Số file ý kiến xử lý: {self.n_feedback_files} "
            f"({self.n_feedbacks} ý kiến đã làm sạch)",
            f"Thời gian phần mềm: {d['auto_minutes']} phút "
            f"({self.auto_seconds:.1f} giây)",
            f"Ước lượng làm thủ công: {d['manual_minutes']} phút "
            f"(~{round(d['manual_minutes']/60,1)} giờ)",
            f"Thời gian tiết kiệm: {d['saved_minutes']} phút "
            f"(~{d['saved_hours']} giờ), giảm {d['reduction_pct']}%",
            f"Nhanh gấp ~{d['speedup']} lần so với làm tay",
        ]


class Timer:
    """Bộ đếm thời gian đơn giản (context manager)."""

    def __init__(self):
        self.elapsed = 0.0
        self._start = None

    def __enter__(self):
        self._start = time.perf_counter()
        return self

    def __exit__(self, *exc):
        self.elapsed = time.perf_counter() - self._start
        return False


def build_result(n_reports: int, auto_seconds: float,
                 n_feedback_files: int = 0, n_feedbacks: int = 0) -> BenchmarkResult:
    """Tạo BenchmarkResult từ số liệu một phiên chạy."""
    load_manual_estimates()
    return BenchmarkResult(
        n_reports=n_reports,
        n_feedback_files=n_feedback_files,
        n_feedbacks=n_feedbacks,
        auto_seconds=auto_seconds,
    )
