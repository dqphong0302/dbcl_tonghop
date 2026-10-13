#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kiểm thử bộ phân loại ý kiến mở offline (feedback_classifier.py).
"""
import os
import sys
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import feedback_classifier as fc


def test_classify_topic_teaching():
    res = fc.classify_topic("Thầy cô giảng bài rất nhiệt tình và tận tâm")
    assert "Giảng viên & Phương pháp giảng dạy" in res


def test_classify_topic_materials():
    res = fc.classify_topic("Slide bài giảng và tài liệu học tập rất chi tiết")
    assert "Tài liệu học tập" in res


def test_classify_topic_facilities():
    res = fc.classify_topic("Phòng học cần lắp thêm điều hòa và cải thiện âm thanh")
    assert "Cơ sở vật chất" in res


def test_classify_topic_unknown():
    res = fc.classify_topic("Em không có thêm nhận xét")
    assert res == ["Khác"]


def test_classify_sentiment_positive():
    sentiment = fc.classify_sentiment("Giáo viên dạy rất hay, bài giảng bổ ích và dễ hiểu")
    assert sentiment == "tích cực"


def test_classify_sentiment_suggestion():
    sentiment = fc.classify_sentiment("Nên thêm thời gian thực hành và giảm bớt bài tập lý thuyết")
    assert sentiment == "góp ý"


def test_classify_sentiment_neutral():
    sentiment = fc.classify_sentiment("Sinh viên học tại cơ sở 1 theo thời khóa biểu")
    assert sentiment == "trung tính"


def test_classify_feedbacks_batch():
    feedbacks = [
        "Thầy dạy rất hay và nhiệt tình",
        "Slide bài giảng cần bổ sung thêm câu hỏi ôn tập",
        "Phòng học hơi nóng, cần bật điều hòa sớm hơn",
        "Không có gì thêm"
    ]
    result = fc.classify_feedbacks(feedbacks)
    assert result.total == 4
    assert len(result.items) == 4
    topic_pct = result.topic_percentages()
    assert isinstance(topic_pct, dict)
    sentiment_pct = result.sentiment_percentages()
    assert isinstance(sentiment_pct, dict)
    assert "tích cực" in sentiment_pct or "góp ý" in sentiment_pct
    summary_lines = fc.summarize(result)
    assert len(summary_lines) > 2


def test_classify_feedbacks_empty():
    result = fc.classify_feedbacks([])
    assert result.total == 0
    assert result.topic_percentages() == {}
    assert result.sentiment_percentages() == {}
    summary = fc.summarize(result)
    assert "Tổng số ý kiến phân loại: 0" in summary[0]
