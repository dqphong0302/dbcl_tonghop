#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phân tích thống kê kết quả khảo sát (Round 2 - phục vụ "Tính hiệu quả" SKKN).

Tính các chỉ số định lượng cho mỗi học phần/lớp dựa trên dữ liệu đã trích
xuất bởi excel_to_word.extract_data_from_excel():
  - Điểm trung bình, độ lệch chuẩn (SD) toàn học phần.
  - Tỷ lệ hài lòng (mức 4-5), tỷ lệ chưa hài lòng (mức 1-2).
  - Phân bố tổng theo 5 mức Likert.
  - Xếp hạng tiêu chí: cao nhất / thấp nhất (gợi ý điểm cần cải tiến).

Không phụ thuộc matplotlib (thuần Python) để dễ test & chạy mọi nơi.
"""

import math
from dataclasses import dataclass, field
from typing import List, Dict, Optional


SCALE_MIN, SCALE_MAX = 1, 5


@dataclass
class QuestionStat:
    number: str
    text: str
    total: int
    counts: List[int]            # 5 phần tử: mức 1..5
    average: float
    satisfied_pct: float          # % mức 4-5
    dissatisfied_pct: float       # % mức 1-2
    neutral_pct: float            # % mức 3


@dataclass
class SurveyAnalysis:
    label: str                              # tên lớp/học phần
    n_questions: int = 0
    n_responses: int = 0                    # tổng phiếu (theo câu đầu)
    overall_average: float = 0.0
    overall_sd: float = 0.0
    overall_satisfied_pct: float = 0.0
    distribution: List[int] = field(default_factory=lambda: [0] * 5)
    distribution_pct: List[float] = field(default_factory=lambda: [0.0] * 5)
    questions: List[QuestionStat] = field(default_factory=list)
    lowest: Optional[QuestionStat] = None
    highest: Optional[QuestionStat] = None

    def to_dict(self) -> Dict:
        return {
            "label": self.label,
            "n_questions": self.n_questions,
            "n_responses": self.n_responses,
            "overall_average": round(self.overall_average, 3),
            "overall_sd": round(self.overall_sd, 3),
            "overall_satisfied_pct": round(self.overall_satisfied_pct, 1),
            "distribution": self.distribution,
            "distribution_pct": [round(p, 1) for p in self.distribution_pct],
            "lowest": self.lowest.number if self.lowest else None,
            "highest": self.highest.number if self.highest else None,
        }


def _safe_counts(counts: List) -> List[int]:
    """Chuẩn hóa list counts về đúng 5 phần tử số nguyên."""
    out = []
    for i in range(5):
        try:
            out.append(int(counts[i]))
        except (IndexError, ValueError, TypeError):
            out.append(0)
    return out


def analyze_question(q: Dict) -> QuestionStat:
    """Tính chỉ số cho 1 câu hỏi (từ dict trong data['survey_questions'])."""
    counts = _safe_counts(q.get("counts", []))
    total = q.get("total", 0) or sum(counts)
    total = int(total) if total else 0

    if total > 0:
        avg = sum((i + 1) * counts[i] for i in range(5)) / total
        satisfied = (counts[3] + counts[4]) / total * 100
        dissatisfied = (counts[0] + counts[1]) / total * 100
        neutral = counts[2] / total * 100
    else:
        avg = satisfied = dissatisfied = neutral = 0.0

    return QuestionStat(
        number=str(q.get("number", "")),
        text=str(q.get("text", "")),
        total=total,
        counts=counts,
        average=round(avg, 2),
        satisfied_pct=round(satisfied, 1),
        dissatisfied_pct=round(dissatisfied, 1),
        neutral_pct=round(neutral, 1),
    )


def _weighted_sd(distribution: List[int], mean: float) -> float:
    """Độ lệch chuẩn của phân bố Likert gộp (theo tần suất)."""
    n = sum(distribution)
    if n <= 1:
        return 0.0
    var = sum(distribution[i] * ((i + 1) - mean) ** 2 for i in range(5)) / n
    return math.sqrt(var)


def _clean_label(meta: Dict) -> str:
    """Tạo nhãn học phần sạch & ổn định để gom nhóm xu hướng giữa các kỳ.

    Ưu tiên tên môn học (bỏ mã số ở đầu), sau đó tới lớp (đã khử trùng lặp),
    cuối cùng tới tên phiếu khảo sát.
    """
    import re
    mon = str(meta.get("mon_hoc", "") or "").strip()
    if mon:
        # Bỏ tiền tố "Môn học:" và mã số (vd "101700821501 - Tên môn")
        mon = re.sub(r"^\s*môn học\s*:?\s*", "", mon, flags=re.IGNORECASE)
        mon = re.sub(r"^\s*[\d.]+\s*[-–]\s*", "", mon).strip()
        if mon:
            return mon[:80]
    lop = str(meta.get("lop_hoc", "") or "").strip()
    if lop:
        # Khử trùng lặp khi gộp nhiều sheet (vd "D25, D25, D25" -> "D25")
        parts = [p.strip() for p in lop.split(",") if p.strip()]
        uniq = list(dict.fromkeys(parts))
        if uniq:
            return ", ".join(uniq)[:80]
    return str(meta.get("phieu_khao_sat", "") or "Khảo sát")[:80]


def analyze_survey(data: Dict, label: Optional[str] = None) -> SurveyAnalysis:
    """Phân tích toàn bộ 1 báo cáo khảo sát đã trích xuất.

    data: dict trả về từ extract_data_from_excel hoặc merge_sheets_data.
    """
    meta = data.get("meta", {})
    lbl = label or _clean_label(meta)

    questions = [analyze_question(q) for q in data.get("survey_questions", [])]
    analysis = SurveyAnalysis(label=lbl, n_questions=len(questions), questions=questions)

    if not questions:
        return analysis

    # Phân bố gộp toàn bộ câu hỏi
    dist = [0] * 5
    for qs in questions:
        for i in range(5):
            dist[i] += qs.counts[i]
    analysis.distribution = dist
    grand_total = sum(dist)

    if grand_total > 0:
        analysis.distribution_pct = [c / grand_total * 100 for c in dist]
        analysis.overall_average = sum((i + 1) * dist[i] for i in range(5)) / grand_total
        analysis.overall_satisfied_pct = (dist[3] + dist[4]) / grand_total * 100
        analysis.overall_sd = _weighted_sd(dist, analysis.overall_average)

    analysis.n_responses = questions[0].total if questions else 0

    # Câu cao nhất / thấp nhất theo điểm TB (chỉ tính câu có dữ liệu)
    valid = [q for q in questions if q.total > 0]
    if valid:
        analysis.lowest = min(valid, key=lambda q: q.average)
        analysis.highest = max(valid, key=lambda q: q.average)

    return analysis


def suggest_improvements(analysis: SurveyAnalysis, top_n: int = 3) -> List[QuestionStat]:
    """Gợi ý các tiêu chí cần cải tiến (điểm TB thấp nhất).

    Phục vụ điền sẵn 'Phần 4. Đề xuất kế hoạch cải tiến' trong báo cáo Word.
    """
    valid = [q for q in analysis.questions if q.total > 0]
    valid.sort(key=lambda q: (q.average, -q.dissatisfied_pct))
    return valid[:top_n]


def compare_analyses(current: SurveyAnalysis,
                     previous: SurveyAnalysis) -> Dict:
    """So sánh 2 kỳ: chênh lệch điểm TB & tỷ lệ hài lòng (cho phần xu hướng)."""
    return {
        "avg_delta": round(current.overall_average - previous.overall_average, 3),
        "satisfied_delta": round(
            current.overall_satisfied_pct - previous.overall_satisfied_pct, 1),
        "current_avg": round(current.overall_average, 2),
        "previous_avg": round(previous.overall_average, 2),
        "improved": current.overall_average >= previous.overall_average,
    }
