#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tóm tắt ý kiến sinh viên thành đoạn văn bằng AI (API tương thích OpenAI, vd 9router).

Chỉ dùng khi bật trong Cấu hình, có API key và kết nối được máy chủ. Mọi lỗi
(không có mạng, hết thời gian, model lỗi...) đều trả về None để nơi gọi
quay về cách xử lý offline như cũ — không bao giờ làm hỏng việc tạo báo cáo.
"""

import json
import re
import urllib.error
import urllib.request
from typing import List, Optional, Tuple

try:
    from app_logging import get_logger
    log = get_logger()
except Exception:
    import logging
    log = logging.getLogger("dbcl")

DEFAULT_PROMPT = (
    "Bạn là chuyên viên đảm bảo chất lượng giáo dục. Dưới đây là các ý kiến đóng góp "
    "của sinh viên sau khi kết thúc học phần. Hãy tổng hợp và tóm tắt đánh giá của sinh viên "
    "thành MỘT đoạn văn tiếng Việt, văn phong hành chính, khách quan (khoảng 80–150 từ): "
    "nêu những điểm sinh viên hài lòng, những vấn đề/đề xuất được nhắc đến nhiều nhất. "
    "Bỏ qua các câu trả lời kiểu \"không có ý kiến\". Không bịa thêm thông tin, không dùng "
    "gạch đầu dòng hay định dạng markdown, không nêu tên cá nhân."
)


def _config():
    """Đọc cấu hình AI hiện tại từ excel_to_word.CONFIG (import muộn để tránh vòng lặp)."""
    from excel_to_word import CONFIG
    return CONFIG


def is_configured(cfg=None) -> bool:
    cfg = cfg or _config()
    return bool(cfg.get("ai_enabled") and cfg.get("ai_api_key") and cfg.get("ai_base_url")
                and cfg.get("ai_model"))


def _chat(cfg, messages, timeout) -> str:
    url = cfg["ai_base_url"].rstrip("/") + "/chat/completions"
    body = json.dumps({"model": cfg["ai_model"], "messages": messages,
                       "temperature": 0.3, "stream": False}).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST", headers={
        "Content-Type": "application/json",
        "Authorization": f"Bearer {cfg['ai_api_key']}",
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return _parse_reply(resp.read().decode("utf-8"))


def _parse_reply(raw: str) -> str:
    """Đọc nội dung trả lời. Chấp nhận cả JSON thường có rác phía sau (vd "data: [DONE]")
    lẫn dạng streaming SSE ("data: {...}" từng dòng)."""
    raw = raw.strip()
    if raw.startswith("{"):
        data, _ = json.JSONDecoder().raw_decode(raw)
        if "error" in data:
            raise RuntimeError(str(data["error"].get("message", data["error"]))[:300])
        return data["choices"][0]["message"].get("content") or ""
    parts = []
    for line in raw.splitlines():
        line = line.strip()
        if not line.startswith("data:") or line == "data: [DONE]":
            continue
        chunk = json.loads(line[5:])
        choice = (chunk.get("choices") or [{}])[0]
        parts.append((choice.get("delta") or choice.get("message") or {}).get("content") or "")
    if not parts:
        raise ValueError(f"Phản hồi không hợp lệ từ máy chủ AI: {raw[:120]}")
    return "".join(parts)


def _clean_paragraph(text: str) -> str:
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S)   # model có suy luận
    text = re.sub(r"[*_#`>]+", "", text)                           # ký hiệu markdown
    text = re.sub(r"^\s*[-•]\s*", "", text, flags=re.M)
    return re.sub(r"\s+", " ", text).strip()


def summarize_feedbacks(opinions: List[str], context: str = "") -> Optional[str]:
    """Trả về đoạn văn tóm tắt, hoặc None nếu AI tắt / lỗi / không có mạng."""
    cfg = _config()
    if not opinions or not is_configured(cfg):
        return None
    prompt = (cfg.get("ai_prompt") or "").strip() or DEFAULT_PROMPT
    listing = "\n".join(f"- {o}" for o in opinions)
    user = (f"{context}\n\n" if context else "") + f"Các ý kiến của sinh viên:\n{listing}"
    try:
        content = _chat(cfg, [{"role": "system", "content": prompt},
                              {"role": "user", "content": user}],
                        timeout=float(cfg.get("ai_timeout", 60)))
        summary = _clean_paragraph(content)
        if len(summary) < 20:
            log.warning("AI trả về nội dung quá ngắn, dùng cách xử lý offline")
            return None
        log.info("AI (%s) đã tóm tắt %d ý kiến", cfg["ai_model"], len(opinions))
        return summary
    except Exception as e:
        log.warning("Không dùng được AI (%s) — chuyển sang xử lý offline: %s",
                    cfg.get("ai_model"), _describe_error(e))
        return None


def _describe_error(e: Exception) -> str:
    if isinstance(e, urllib.error.HTTPError):
        try:
            detail = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
        except Exception:
            detail = ""
        return f"HTTP {e.code} {detail}".strip()[:300]
    if isinstance(e, urllib.error.URLError):
        return f"không kết nối được máy chủ ({e.reason})"
    return str(e)[:300]


def test_connection(cfg) -> Tuple[bool, str]:
    """Gửi một yêu cầu ngắn để kiểm tra URL / key / model. Trả về (thành công, thông báo)."""
    if not (cfg.get("ai_api_key") and cfg.get("ai_base_url") and cfg.get("ai_model")):
        return False, "Chưa nhập đủ URL, API key và model."
    try:
        reply = _chat(cfg, [{"role": "user", "content": "Trả lời đúng một từ: OK"}], timeout=30)
        return True, f"Kết nối thành công với model “{cfg['ai_model']}” (trả lời: {reply.strip()[:40]})"
    except Exception as e:
        return False, _describe_error(e)
