#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phân loại ý kiến mở theo CHỦ ĐỀ & SẮC THÁI - hoàn toàn OFFLINE (Round 3 - SKKN).

Đây là điểm SÁNG TẠO ("tính mới") của phần mềm: thay vì chỉ liệt kê ý kiến,
phần mềm tự động:
  - Gom ý kiến vào các CHỦ ĐỀ (giảng viên, nội dung học phần, cơ sở vật chất,
    lượng giá/đánh giá, tài liệu, thời lượng, khác) bằng từ điển từ khóa TV.
  - Phân loại SẮC THÁI: tích cực (khen) / góp ý (đề xuất, cần cải thiện) /
    trung tính.
  - Tổng hợp định lượng: "X% ý kiến liên quan phương pháp giảng dạy"...

Không dùng mạng/LLM => an toàn dữ liệu sinh viên, chạy mọi máy.
"""

import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List

# ───────────────────────────────────────────────────────────────
# TỪ ĐIỂN CHỦ ĐỀ (keyword tiếng Việt, không dấu/có dấu đều khớp lower)
# ───────────────────────────────────────────────────────────────

TOPIC_KEYWORDS: Dict[str, List[str]] = {
    "Giảng viên & Phương pháp giảng dạy": [
        "giảng viên", "giáo viên", "thầy", "cô", "gv ", "giảng dạy", "bài giảng",
        "phương pháp", "cách dạy", "nhiệt tình", "tâm huyết", "dạy hay",
        "truyền đạt", "giảng bài", "tương tác", "vui tính", "dễ hiểu",
    ],
    "Nội dung & Chương trình học phần": [
        "nội dung", "kiến thức", "chương trình", "môn học", "học phần",
        "bài học", "chủ đề", "cập nhật", "thực tế", "ứng dụng", "lý thuyết",
        "thực hành", "bài tập",
    ],
    "Tài liệu học tập": [
        "tài liệu", "giáo trình", "slide", "bài đọc", "đề cương", "sách",
        "tư liệu", "học liệu",
    ],
    "Lượng giá & Đánh giá": [
        "lượng giá", "đánh giá", "thi", "kiểm tra", "điểm", "bài thi",
        "chấm", "trắc nghiệm", "tự luận", "khiếu nại", "công bằng", "kết quả",
    ],
    "Thời lượng & Tổ chức lớp": [
        "thời lượng", "thời gian", "tiết học", "buổi học", "lịch học",
        "phân bổ", "số buổi", "tổ chức", "sắp xếp", "giờ học",
    ],
    "Cơ sở vật chất": [
        "phòng học", "cơ sở vật chất", "trang thiết bị", "máy chiếu",
        "âm thanh", "điều hòa", "wifi", "thiết bị", "sân", "dụng cụ",
        "phòng tập", "nhà thi đấu",
    ],
}

# ───────────────────────────────────────────────────────────────
# TỪ KHÓA SẮC THÁI
# ───────────────────────────────────────────────────────────────

POSITIVE_WORDS = [
    "tốt", "hay", "hài lòng", "tuyệt", "thích", "bổ ích", "hữu ích", "nhiệt tình",
    "tâm huyết", "dễ hiểu", "cảm ơn", "yêu thích", "tuyệt vời", "xuất sắc",
    "vui", "hấp dẫn", "chất lượng", "ổn", "rất tốt", "giỏi", "thân thiện",
]

SUGGEST_WORDS = [
    "nên", "cần", "mong", "đề xuất", "góp ý", "cải thiện", "cải tiến",
    "hy vọng", "hi vọng", "giá như", "thêm", "tăng", "giảm", "bớt", "hơn",
    "chưa", "khó", "thiếu", "quá", "hạn chế", "khắc phục", "điều chỉnh",
    "cần thiết", "đề nghị",
]


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip().lower())


# ───────────────────────────────────────────────────────────────
# KẾT QUẢ PHÂN LOẠI
# ───────────────────────────────────────────────────────────────

@dataclass
class ClassifiedFeedback:
    text: str
    topics: List[str]
    sentiment: str  # "tích cực" | "góp ý" | "trung tính"


@dataclass
class ClassificationResult:
    items: List[ClassifiedFeedback] = field(default_factory=list)
    topic_counts: Dict[str, int] = field(default_factory=dict)
    sentiment_counts: Dict[str, int] = field(default_factory=dict)

    @property
    def total(self) -> int:
        return len(self.items)

    def topic_percentages(self) -> Dict[str, float]:
        if not self.total:
            return {}
        return {k: round(v / self.total * 100, 1)
                for k, v in self.topic_counts.items()}

    def sentiment_percentages(self) -> Dict[str, float]:
        if not self.total:
            return {}
        return {k: round(v / self.total * 100, 1)
                for k, v in self.sentiment_counts.items()}

    def by_topic(self) -> Dict[str, List[ClassifiedFeedback]]:
        groups: Dict[str, List[ClassifiedFeedback]] = defaultdict(list)
        for item in self.items:
            for t in item.topics:
                groups[t].append(item)
        return dict(groups)


# ───────────────────────────────────────────────────────────────
# LÕI PHÂN LOẠI
# ───────────────────────────────────────────────────────────────

OTHER_TOPIC = "Khác"


def classify_topic(text: str) -> List[str]:
    """Trả về danh sách chủ đề khớp (có thể nhiều); rỗng -> ['Khác']."""
    norm = _norm(text)
    matched = []
    for topic, kws in TOPIC_KEYWORDS.items():
        if any(kw in norm for kw in kws):
            matched.append(topic)
    return matched if matched else [OTHER_TOPIC]


def classify_sentiment(text: str) -> str:
    """Phân loại sắc thái dựa trên từ khóa khen/góp ý."""
    norm = _norm(text)
    pos = sum(1 for w in POSITIVE_WORDS if w in norm)
    sug = sum(1 for w in SUGGEST_WORDS if w in norm)
    if sug > pos:
        return "góp ý"
    if pos > 0:
        return "tích cực"
    return "trung tính"


def classify_feedbacks(feedbacks: List[str]) -> ClassificationResult:
    """Phân loại toàn bộ danh sách ý kiến (đã làm sạch)."""
    result = ClassificationResult()
    topic_counter: Dict[str, int] = defaultdict(int)
    sentiment_counter: Dict[str, int] = defaultdict(int)

    for fb in feedbacks:
        fb = (fb or "").strip()
        if not fb:
            continue
        topics = classify_topic(fb)
        sentiment = classify_sentiment(fb)
        result.items.append(ClassifiedFeedback(fb, topics, sentiment))
        # Đếm chủ đề: mỗi ý kiến tính 1 lần cho mỗi chủ đề khớp
        for t in topics:
            topic_counter[t] += 1
        sentiment_counter[sentiment] += 1

    # Sắp xếp chủ đề theo số lượng giảm dần
    result.topic_counts = dict(
        sorted(topic_counter.items(), key=lambda x: -x[1]))
    result.sentiment_counts = dict(
        sorted(sentiment_counter.items(), key=lambda x: -x[1]))
    return result


def summarize(result: ClassificationResult) -> List[str]:
    """Tạo các dòng tóm tắt định lượng để log/đưa vào báo cáo."""
    lines = [f"Tổng số ý kiến phân loại: {result.total}"]
    if result.total == 0:
        return lines
    lines.append("Theo chủ đề:")
    for topic, pct in result.topic_percentages().items():
        cnt = result.topic_counts[topic]
        lines.append(f"  • {topic}: {cnt} ý kiến ({pct}%)")
    lines.append("Theo sắc thái:")
    for s, pct in result.sentiment_percentages().items():
        cnt = result.sentiment_counts[s]
        lines.append(f"  • {s.capitalize()}: {cnt} ý kiến ({pct}%)")
    return lines
