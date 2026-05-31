#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lưu & so sánh LỊCH SỬ khảo sát qua các kỳ (Round 3 - SKKN).

Phục vụ minh chứng "cải tiến liên tục" cho công tác ĐBCL:
  - Sau mỗi đợt xử lý, lưu lại bản ghi gọn (điểm TB, % hài lòng, phân bố...).
  - Truy xuất lịch sử theo học phần để vẽ biểu đồ xu hướng giữa các kỳ.

Dữ liệu lưu dạng JSON tại thư mục cấu hình của người dùng (offline, an toàn).
Mỗi bản ghi định danh bằng (học phần + kỳ/đợt khảo sát).
"""

import json
import os
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from analytics import SurveyAnalysis

HISTORY_FILE = Path.home() / ".dbcl_tonghop_history.json"


@dataclass
class HistoryRecord:
    course: str                 # tên học phần/lớp
    period: str                 # kỳ/đợt khảo sát (vd "HKI 25-26")
    timestamp: str              # thời điểm ghi nhận
    overall_average: float
    overall_satisfied_pct: float
    n_responses: int
    distribution: List[int]

    def key(self) -> str:
        return f"{self.course}||{self.period}"


def _load_raw() -> List[Dict]:
    try:
        if HISTORY_FILE.exists():
            data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
    except Exception:
        pass
    return []


def _save_raw(records: List[Dict]):
    try:
        HISTORY_FILE.write_text(
            json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _infer_period(analysis: SurveyAnalysis, meta: Optional[Dict] = None) -> str:
    """Suy ra tên kỳ/đợt khảo sát từ metadata nếu có."""
    if meta:
        dot = meta.get("dot_khao_sat") or meta.get("thoi_gian_ks")
        if dot:
            return str(dot)[:60]
    return datetime.now().strftime("%Y-%m-%d")


def record_analysis(analysis: SurveyAnalysis, meta: Optional[Dict] = None,
                    period: Optional[str] = None) -> HistoryRecord:
    """Ghi nhận 1 kết quả phân tích vào lịch sử (ghi đè nếu trùng course+period)."""
    rec = HistoryRecord(
        course=analysis.label,
        period=period or _infer_period(analysis, meta),
        timestamp=datetime.now().isoformat(timespec="seconds"),
        overall_average=round(analysis.overall_average, 3),
        overall_satisfied_pct=round(analysis.overall_satisfied_pct, 1),
        n_responses=analysis.n_responses,
        distribution=list(analysis.distribution),
    )
    records = _load_raw()
    # Ghi đè nếu đã có cùng key
    records = [r for r in records
               if f"{r.get('course')}||{r.get('period')}" != rec.key()]
    records.append(asdict(rec))
    _save_raw(records)
    return rec


def get_history(course: Optional[str] = None) -> List[Dict]:
    """Lấy lịch sử (lọc theo học phần nếu chỉ định), sắp theo thời gian ghi."""
    records = _load_raw()
    if course:
        records = [r for r in records if r.get("course") == course]
    records.sort(key=lambda r: r.get("timestamp", ""))
    return records


def list_courses() -> List[str]:
    """Danh sách học phần đã có trong lịch sử."""
    seen = []
    for r in _load_raw():
        c = r.get("course")
        if c and c not in seen:
            seen.append(c)
    return sorted(seen)


def history_for_trend(course: str) -> List[Dict]:
    """Chuẩn bị dữ liệu cho charts.chart_trend()."""
    out = []
    for r in get_history(course):
        out.append({
            "period": r.get("period", ""),
            "overall_average": r.get("overall_average", 0),
            "overall_satisfied_pct": r.get("overall_satisfied_pct", 0),
        })
    return out


def clear_history():
    """Xóa toàn bộ lịch sử (dùng khi cần làm lại từ đầu)."""
    _save_raw([])
