#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Test cho các module phân tích Round 2: analytics, benchmark, charts."""

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import analytics
import benchmark
import charts
import excel_to_word as e2w

DULIEU = os.path.join(ROOT, "dulieu")
SURVEY_FILE = os.path.join(DULIEU, "CTXH25 Giáo dục thể chất 1 25-26 HKI.xls")


# ─── analytics ──────────────────────────────────────────────────

def _sample_analysis():
    data = e2w.extract_data_from_excel(SURVEY_FILE)
    return analytics.analyze_survey(data)


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_analyze_survey_basic():
    a = _sample_analysis()
    assert a.n_questions == 22
    assert 1.0 <= a.overall_average <= 5.0
    assert 0.0 <= a.overall_satisfied_pct <= 100.0
    assert len(a.distribution) == 5
    assert sum(a.distribution) > 0


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_distribution_pct_sums_100():
    a = _sample_analysis()
    assert abs(sum(a.distribution_pct) - 100.0) < 0.5


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_lowest_highest_consistent():
    a = _sample_analysis()
    assert a.lowest.average <= a.highest.average


@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_suggest_improvements_sorted():
    a = _sample_analysis()
    sug = analytics.suggest_improvements(a, top_n=3)
    assert len(sug) <= 3
    avgs = [q.average for q in sug]
    assert avgs == sorted(avgs)


def test_analyze_empty_survey():
    a = analytics.analyze_survey({"meta": {}, "survey_questions": []})
    assert a.n_questions == 0
    assert a.overall_average == 0.0


def test_compare_analyses():
    cur = analytics.SurveyAnalysis(label="cur", overall_average=4.5,
                                   overall_satisfied_pct=92.0)
    prev = analytics.SurveyAnalysis(label="prev", overall_average=4.2,
                                    overall_satisfied_pct=88.0)
    cmp = analytics.compare_analyses(cur, prev)
    assert cmp["improved"] is True
    assert abs(cmp["avg_delta"] - 0.3) < 1e-6


# ─── benchmark ──────────────────────────────────────────────────

def test_benchmark_savings():
    r = benchmark.build_result(n_reports=10, auto_seconds=5.0,
                               n_feedback_files=3, n_feedbacks=50)
    d = r.to_dict()
    assert d["manual_minutes"] > 0
    assert d["saved_minutes"] > 0
    assert 0 <= d["reduction_pct"] <= 100
    assert r.speedup > 1


def test_benchmark_timer():
    with benchmark.Timer() as t:
        sum(range(1000))
    assert t.elapsed >= 0


def test_benchmark_zero_reports():
    r = benchmark.build_result(n_reports=0, auto_seconds=0.0)
    assert r.manual_minutes == 0
    assert r.reduction_pct == 0


# ─── charts ─────────────────────────────────────────────────────

@pytest.mark.skipif(not os.path.exists(SURVEY_FILE), reason="thiếu file mẫu")
def test_generate_charts(tmp_path):
    a = _sample_analysis()
    res = charts.generate_all_charts(a, str(tmp_path), prefix="test")
    assert "distribution" in res
    for path in res.values():
        assert os.path.exists(path)
        assert os.path.getsize(path) > 1000


def test_chart_trend_needs_two_points(tmp_path):
    out = charts.chart_trend([{"period": "A", "overall_average": 4.0,
                               "overall_satisfied_pct": 80}],
                             str(tmp_path / "t.png"))
    assert out is None
