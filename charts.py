#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sinh biểu đồ trực quan từ kết quả phân tích khảo sát (Round 2 - SKKN).

Dùng matplotlib (backend 'Agg', không cần GUI) để xuất PNG:
  - Biểu đồ tròn/cột phân bố 5 mức Likert.
  - Biểu đồ cột điểm trung bình theo từng tiêu chí.
  - Biểu đồ xu hướng giữa các kỳ (nếu có dữ liệu lịch sử).

Các ảnh PNG vừa nhúng vào báo cáo Word, vừa lưu riêng để đưa vào SKKN.
Xử lý font tiếng Việt: ưu tiên các font có sẵn hỗ trợ Unicode.
"""

import os
from typing import List, Optional

import matplotlib
matplotlib.use("Agg")  # không cần màn hình
import matplotlib.pyplot as plt
from matplotlib import font_manager

from analytics import SurveyAnalysis

# ───────────────────────────────────────────────────────────────
# CẤU HÌNH MÀU & FONT
# ───────────────────────────────────────────────────────────────

# Bảng màu hài hòa (xanh dương -> xanh lá), tránh màu thuần
LEVEL_COLORS = ["#ef4444", "#f59e0b", "#facc15", "#60a5fa", "#22c55e"]
ACCENT = "#4f46e5"
ACCENT2 = "#0ea5e9"
GRID = "#e5e7eb"

LEVEL_LABELS = ["Mức 1", "Mức 2", "Mức 3", "Mức 4", "Mức 5"]
LEVEL_LABELS_FULL = [
    "1. Hoàn toàn\nkhông đồng ý",
    "2. Không\nđồng ý",
    "3. Đồng ý\nmột phần",
    "4. Đồng ý",
    "5. Hoàn toàn\nđồng ý",
]

_FONT_READY = False


def _setup_font():
    """Chọn font hỗ trợ tiếng Việt có sẵn trên hệ thống."""
    global _FONT_READY
    if _FONT_READY:
        return
    candidates = [
        "Arial", "Helvetica", "DejaVu Sans", "Liberation Sans",
        "Times New Roman", "Tahoma", "Verdana", "Segoe UI",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = next((c for c in candidates if c in available), "DejaVu Sans")
    plt.rcParams["font.family"] = chosen
    plt.rcParams["axes.unicode_minus"] = False
    _FONT_READY = True


def _finalize(fig, out_path: str, dpi: int = 150) -> str:
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    fig.tight_layout()
    fig.savefig(out_path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path


# ───────────────────────────────────────────────────────────────
# BIỂU ĐỒ PHÂN BỐ 5 MỨC
# ───────────────────────────────────────────────────────────────

def chart_distribution(analysis: SurveyAnalysis, out_path: str) -> Optional[str]:
    """Biểu đồ cột phân bố tổng số lượt trả lời theo 5 mức Likert."""
    _setup_font()
    dist = analysis.distribution
    if sum(dist) == 0:
        return None

    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    bars = ax.bar(LEVEL_LABELS_FULL, dist, color=LEVEL_COLORS,
                  edgecolor="white", linewidth=1.2, zorder=3)

    total = sum(dist)
    for bar, val in zip(bars, dist):
        pct = val / total * 100 if total else 0
        ax.annotate(f"{val}\n({pct:.1f}%)",
                    xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    xytext=(0, 4), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")

    ax.set_title(f"Phân bố mức độ đồng ý - {analysis.label}",
                 fontsize=13, fontweight="bold", pad=14)
    ax.set_ylabel("Số lượt trả lời", fontsize=10)
    ax.grid(axis="y", color=GRID, zorder=0)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.margins(y=0.18)
    return _finalize(fig, out_path)


# ───────────────────────────────────────────────────────────────
# BIỂU ĐỒ ĐIỂM TRUNG BÌNH THEO TIÊU CHÍ
# ───────────────────────────────────────────────────────────────

def chart_question_averages(analysis: SurveyAnalysis, out_path: str) -> Optional[str]:
    """Biểu đồ cột ngang điểm TB từng tiêu chí, tô đậm câu thấp nhất."""
    _setup_font()
    qs = [q for q in analysis.questions if q.total > 0]
    if not qs:
        return None

    labels = [f"C{q.number}" for q in qs]
    values = [q.average for q in qs]
    min_avg = min(values)

    colors = [ACCENT if v > min_avg else "#ef4444" for v in values]

    height = max(4.0, len(qs) * 0.32 + 1.2)
    fig, ax = plt.subplots(figsize=(8.0, height))
    y = range(len(qs))
    bars = ax.barh(list(y), values, color=colors, edgecolor="white",
                   linewidth=0.8, zorder=3)
    ax.set_yticks(list(y))
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlim(0, 5.4)

    for bar, val in zip(bars, values):
        ax.annotate(f"{val:.2f}",
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(4, 0), textcoords="offset points",
                    ha="left", va="center", fontsize=8, fontweight="bold")

    # Đường tham chiếu điểm TB chung
    ax.axvline(analysis.overall_average, color=ACCENT2, linestyle="--",
               linewidth=1.4, zorder=4,
               label=f"TB chung = {analysis.overall_average:.2f}")
    ax.legend(loc="lower right", fontsize=9)

    ax.set_title(f"Điểm trung bình theo tiêu chí - {analysis.label}",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_xlabel("Điểm trung bình (thang 1-5)", fontsize=10)
    ax.grid(axis="x", color=GRID, zorder=0)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    return _finalize(fig, out_path)


# ───────────────────────────────────────────────────────────────
# BIỂU ĐỒ TỶ LỆ HÀI LÒNG (DONUT)
# ───────────────────────────────────────────────────────────────

def chart_satisfaction(analysis: SurveyAnalysis, out_path: str) -> Optional[str]:
    """Biểu đồ donut: hài lòng (4-5) / trung lập (3) / chưa hài lòng (1-2)."""
    _setup_font()
    dist = analysis.distribution
    total = sum(dist)
    if total == 0:
        return None

    satisfied = dist[3] + dist[4]
    neutral = dist[2]
    dissatisfied = dist[0] + dist[1]
    sizes = [satisfied, neutral, dissatisfied]
    labels = ["Hài lòng\n(mức 4-5)", "Trung lập\n(mức 3)", "Chưa hài lòng\n(mức 1-2)"]
    colors = ["#22c55e", "#facc15", "#ef4444"]

    # Bỏ phần bằng 0 để biểu đồ sạch
    data = [(s, l, c) for s, l, c in zip(sizes, labels, colors) if s > 0]
    sizes, labels, colors = zip(*data)

    fig, ax = plt.subplots(figsize=(6.0, 5.0))
    wedges, _texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors, autopct="%1.1f%%",
        startangle=90, pctdistance=0.78,
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=10),
    )
    for at in autotexts:
        at.set_fontweight("bold")
        at.set_fontsize(10)

    ax.text(0, 0, f"{analysis.overall_satisfied_pct:.0f}%\nhài lòng",
            ha="center", va="center", fontsize=14, fontweight="bold",
            color="#15803d")
    ax.set_title(f"Tỷ lệ hài lòng - {analysis.label}",
                 fontsize=13, fontweight="bold", pad=12)
    return _finalize(fig, out_path)


# ───────────────────────────────────────────────────────────────
# BIỂU ĐỒ XU HƯỚNG GIỮA CÁC KỲ
# ───────────────────────────────────────────────────────────────

def chart_trend(history: List[dict], out_path: str,
                title: str = "Xu hướng qua các kỳ khảo sát") -> Optional[str]:
    """Biểu đồ đường: điểm TB & % hài lòng qua các kỳ.

    history: list dict, mỗi phần tử có 'period', 'overall_average',
             'overall_satisfied_pct' (thứ tự thời gian tăng dần).
    """
    _setup_font()
    if len(history) < 2:
        return None

    periods = [h.get("period", str(i + 1)) for i, h in enumerate(history)]
    avgs = [h.get("overall_average", 0) for h in history]
    sats = [h.get("overall_satisfied_pct", 0) for h in history]

    fig, ax1 = plt.subplots(figsize=(8.0, 4.6))
    ax1.plot(periods, avgs, "-o", color=ACCENT, linewidth=2.2,
             markersize=7, label="Điểm TB", zorder=3)
    ax1.set_ylabel("Điểm trung bình (1-5)", color=ACCENT, fontsize=10)
    ax1.tick_params(axis="y", labelcolor=ACCENT)
    ax1.set_ylim(1, 5.2)
    for x, v in zip(periods, avgs):
        ax1.annotate(f"{v:.2f}", (x, v), textcoords="offset points",
                     xytext=(0, 9), ha="center", fontsize=8,
                     fontweight="bold", color=ACCENT)

    ax2 = ax1.twinx()
    ax2.plot(periods, sats, "--s", color="#22c55e", linewidth=2.0,
             markersize=6, label="% hài lòng", zorder=3)
    ax2.set_ylabel("% hài lòng (mức 4-5)", color="#16a34a", fontsize=10)
    ax2.tick_params(axis="y", labelcolor="#16a34a")
    ax2.set_ylim(0, 105)

    ax1.set_title(title, fontsize=13, fontweight="bold", pad=12)
    ax1.grid(axis="y", color=GRID, zorder=0)
    ax1.set_axisbelow(True)
    lines1, lbl1 = ax1.get_legend_handles_labels()
    lines2, lbl2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, lbl1 + lbl2, loc="lower center",
               ncol=2, fontsize=9)
    return _finalize(fig, out_path)


def generate_all_charts(analysis: SurveyAnalysis, out_dir: str,
                        prefix: str = "chart") -> dict:
    """Sinh toàn bộ biểu đồ cho 1 báo cáo, trả về dict {tên: đường_dẫn}."""
    os.makedirs(out_dir, exist_ok=True)
    result = {}
    safe = "".join(ch if ch.isalnum() else "_" for ch in prefix)[:40]
    mapping = {
        "distribution": chart_distribution,
        "averages": chart_question_averages,
        "satisfaction": chart_satisfaction,
    }
    for key, fn in mapping.items():
        path = os.path.join(out_dir, f"{safe}_{key}.png")
        try:
            res = fn(analysis, path)
            if res:
                result[key] = res
        except Exception:
            pass
    return result
