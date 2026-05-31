#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Test cho Round 3: feedback_classifier và history_store."""

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import feedback_classifier as fc
import history_store as hs
import analytics


# ─── feedback_classifier ────────────────────────────────────────

def test_classify_topic_teacher():
    topics = fc.classify_topic("Giảng viên dạy rất nhiệt tình và dễ hiểu")
    assert "Giảng viên & Phương pháp giảng dạy" in topics


def test_classify_topic_facility():
    topics = fc.classify_topic("Sân tập nắng quá, cần thêm mái che")
    assert "Cơ sở vật chất" in topics


def test_classify_topic_other():
    topics = fc.classify_topic("Xyz abc 123")
    assert topics == [fc.OTHER_TOPIC]


def test_classify_sentiment_positive():
    assert fc.classify_sentiment("Thầy dạy rất hay và bổ ích") == "tích cực"


def test_classify_sentiment_suggest():
    assert fc.classify_sentiment("Nên tăng thêm thời lượng thực hành") == "góp ý"


def test_classify_feedbacks_counts():
    fbs = [
        "Giảng viên nhiệt tình, dạy hay",
        "Nên cung cấp thêm tài liệu học tập",
        "Phòng học cần thêm máy chiếu",
    ]
    res = fc.classify_feedbacks(fbs)
    assert res.total == 3
    pcts = res.topic_percentages()
    assert sum(res.sentiment_counts.values()) == 3
    assert all(0 <= v <= 100 for v in pcts.values())


def test_classify_empty():
    res = fc.classify_feedbacks([])
    assert res.total == 0
    assert res.topic_percentages() == {}


def test_summarize_lines():
    res = fc.classify_feedbacks(["Thầy dạy hay", "Cần thêm bài tập"])
    lines = fc.summarize(res)
    assert any("Tổng số ý kiến" in l for l in lines)


# ─── history_store ──────────────────────────────────────────────

@pytest.fixture
def temp_history(tmp_path, monkeypatch):
    """Cô lập file lịch sử vào thư mục tạm."""
    fake = tmp_path / "history.json"
    monkeypatch.setattr(hs, "HISTORY_FILE", fake)
    yield fake


def _mk_analysis(label, avg, sat):
    return analytics.SurveyAnalysis(
        label=label, overall_average=avg, overall_satisfied_pct=sat,
        n_responses=50, distribution=[1, 1, 5, 40, 53])


def test_record_and_get(temp_history):
    hs.record_analysis(_mk_analysis("Môn A", 4.2, 88), period="HK1")
    hs.record_analysis(_mk_analysis("Môn A", 4.4, 91), period="HK2")
    history = hs.get_history("Môn A")
    assert len(history) == 2
    assert hs.list_courses() == ["Môn A"]


def test_record_overwrite_same_key(temp_history):
    hs.record_analysis(_mk_analysis("Môn B", 4.0, 80), period="HK1")
    hs.record_analysis(_mk_analysis("Môn B", 4.5, 90), period="HK1")
    history = hs.get_history("Môn B")
    assert len(history) == 1
    assert history[0]["overall_average"] == 4.5


def test_history_for_trend(temp_history):
    hs.record_analysis(_mk_analysis("Môn C", 4.1, 85), period="HK1")
    hs.record_analysis(_mk_analysis("Môn C", 4.3, 89), period="HK2")
    trend = hs.history_for_trend("Môn C")
    assert len(trend) == 2
    assert all("overall_average" in t for t in trend)


def test_clear_history(temp_history):
    hs.record_analysis(_mk_analysis("Môn D", 4.0, 80), period="HK1")
    hs.clear_history()
    assert hs.list_courses() == []
