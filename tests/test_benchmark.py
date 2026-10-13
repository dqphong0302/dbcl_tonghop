#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kiểm thử module đo lường hiệu năng và ước lượng thời gian (benchmark.py).
"""
import os
import sys
import time
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import benchmark


def test_timer_context():
    with benchmark.Timer() as t:
        time.sleep(0.05)
    assert t.elapsed >= 0.04


def test_benchmark_result_calculations():
    # 2 báo cáo, 1 file feedback, 30 giây máy chạy
    res = benchmark.BenchmarkResult(
        n_reports=2,
        n_feedback_files=1,
        n_feedbacks=15,
        auto_seconds=30.0
    )
    # manual = 2 * 45 + 1 * 30 = 120 phút
    assert res.manual_minutes == 120.0
    # auto = 0.5 phút
    assert res.auto_minutes == 0.5
    # saved = 120 - 0.5 = 119.5 phút
    assert res.saved_minutes == 119.5
    # reduction = 119.5 / 120 * 100 > 99%
    assert res.reduction_pct > 99.0
    # speedup = 120 / 0.5 = 240
    assert res.speedup == 240.0

    d = res.to_dict()
    assert d["n_reports"] == 2
    assert d["manual_minutes"] == 120.0
    assert d["speedup"] == 240.0

    lines = res.summary_lines()
    assert len(lines) == 6
    assert any("tiết kiệm" in line for line in lines)


def test_build_result_helper():
    res = benchmark.build_result(n_reports=1, auto_seconds=6.0)
    assert res.n_reports == 1
    assert res.auto_seconds == 6.0
    assert res.manual_minutes == benchmark.MANUAL_MINUTES_PER_REPORT
