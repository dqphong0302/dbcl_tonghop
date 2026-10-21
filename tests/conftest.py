import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture(autouse=True)
def _no_ai(monkeypatch):
    """Không gọi AI qua mạng khi chạy test (cấu hình cá nhân có thể đang bật AI)."""
    import excel_to_word
    monkeypatch.setitem(excel_to_word.CONFIG, "ai_enabled", False)
