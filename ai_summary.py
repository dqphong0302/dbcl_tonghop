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
import ssl
import time
import urllib.error
import urllib.request
from typing import Dict, List, Optional, Tuple

try:
    from app_logging import get_logger
    log = get_logger()
except Exception:
    import logging
    log = logging.getLogger("dbcl")

_RULES = (
    " Bỏ qua các câu trả lời kiểu \"không có ý kiến\". Không bịa thêm thông tin, không dùng "
    "gạch đầu dòng hay định dạng markdown, không nêu tên cá nhân."
)

DEFAULT_PROMPT = (
    "Bạn là chuyên viên đảm bảo chất lượng giáo dục. Dưới đây là {so_y_kien} ý kiến đóng góp "
    "của sinh viên sau khi kết thúc học phần {mon_hoc}. Hãy tổng hợp và tóm tắt đánh giá của "
    "sinh viên thành MỘT đoạn văn tiếng Việt, văn phong hành chính, khách quan (khoảng 80–150 từ): "
    "nêu những điểm sinh viên hài lòng, những vấn đề/đề xuất được nhắc đến nhiều nhất." + _RULES
)

DEFAULT_FEEDBACK_PROMPT = (
    "Bạn là chuyên viên đảm bảo chất lượng giáo dục. Dưới đây là {so_y_kien} ý kiến đóng góp "
    "của sinh viên, gộp từ nhiều lớp/học phần. Hãy tổng hợp thành MỘT đoạn văn tiếng Việt, "
    "văn phong hành chính, khách quan (khoảng 120–200 từ): nhận xét chung, các nhóm vấn đề "
    "được phản ánh nhiều nhất (giảng viên, nội dung, lượng giá, thời lượng, cơ sở vật chất...) "
    "và các đề xuất cải tiến chính." + _RULES
)

DEFAULT_PROMPTS = {"report": DEFAULT_PROMPT, "feedback": DEFAULT_FEEDBACK_PROMPT}

# Mẫu prompt chọn nhanh trong trang Cấu hình
PRESETS = {
    "Đoạn văn hành chính (mặc định)": DEFAULT_PROMPT,
    "Ngắn gọn 2–3 câu": (
        "Tóm tắt {so_y_kien} ý kiến của sinh viên về học phần {mon_hoc} thành 2–3 câu tiếng Việt "
        "ngắn gọn, khách quan: mức độ hài lòng chung và vấn đề nổi bật nhất." + _RULES),
    "Điểm mạnh – Hạn chế – Đề xuất": (
        "Bạn là chuyên viên đảm bảo chất lượng. Từ {so_y_kien} ý kiến của sinh viên về học phần "
        "{mon_hoc}, viết MỘT đoạn văn tiếng Việt (100–180 từ) lần lượt nêu: điểm mạnh được sinh viên "
        "ghi nhận, các hạn chế/vấn đề còn tồn tại, và đề xuất cải tiến cụ thể cho bộ môn." + _RULES),
    "Báo cáo gửi Bộ môn (có định lượng)": (
        "Viết MỘT đoạn văn tiếng Việt văn phong báo cáo gửi Bộ môn, tổng hợp {so_y_kien} ý kiến "
        "của sinh viên lớp {lop} về học phần {mon_hoc}. Dùng các cụm định lượng ước lượng như "
        "\"đa số\", \"một số\", \"ít\" sinh viên khi nêu từng nhận xét; kết thúc bằng 1 câu "
        "kiến nghị." + _RULES),
}

PLACEHOLDERS = {
    "{mon_hoc}": "tên học phần",
    "{lop}": "mã lớp",
    "{so_y_kien}": "số ý kiến",
}


def fill_placeholders(prompt: str, variables: Optional[Dict[str, str]] = None) -> str:
    """Thay {mon_hoc}, {lop}, {so_y_kien}… — chỉ thay biến đã biết, giữ nguyên ngoặc khác."""
    for key, value in (variables or {}).items():
        prompt = prompt.replace("{" + key + "}", str(value))
    for ph in PLACEHOLDERS:          # biến không có dữ liệu → bỏ trống
        prompt = prompt.replace(ph, "")
    return re.sub(r"[ \t]{2,}", " ", prompt).replace(" .", ".").strip()


def get_prompt(kind: str = "report", cfg=None) -> str:
    cfg = cfg or _config()
    key = "ai_prompt" if kind == "report" else "ai_prompt_feedback"
    return (cfg.get(key) or "").strip() or DEFAULT_PROMPTS[kind]


def _config():
    """Đọc cấu hình AI hiện tại từ excel_to_word.CONFIG (import muộn để tránh vòng lặp)."""
    from excel_to_word import CONFIG
    return CONFIG


def is_configured(cfg=None) -> bool:
    cfg = cfg or _config()
    return bool(cfg.get("ai_enabled") and cfg.get("ai_api_key") and cfg.get("ai_base_url")
                and cfg.get("ai_model"))


# Sau lỗi mạng (không kết nối/quá thời gian), tạm bỏ qua AI một lúc để các file sau
# không phải chờ timeout lần lượt.
_OFFLINE_BACKOFF = 300
_offline_until = 0.0


def _ssl_context():
    """Dùng bộ chứng chỉ certifi nếu có — bản đóng gói (PyInstaller, nhất là macOS)
    thường không tìm thấy chứng chỉ hệ thống."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _chat(cfg, messages, timeout) -> str:
    url = cfg["ai_base_url"].rstrip("/") + "/chat/completions"
    body = json.dumps({"model": cfg["ai_model"], "messages": messages,
                       "temperature": 0.3, "stream": False}).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST", headers={
        "Content-Type": "application/json",
        "Authorization": f"Bearer {cfg['ai_api_key']}",
    })
    with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as resp:
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


def summarize_feedbacks(opinions: List[str], kind: str = "report",
                        variables: Optional[Dict[str, str]] = None,
                        prompt: Optional[str] = None) -> Optional[str]:
    """Trả về đoạn văn tóm tắt, hoặc None nếu AI tắt / lỗi / không có mạng.

    kind: "report" (Phần 3 báo cáo) hoặc "feedback" (Feedbacks.docx) — chọn prompt tương ứng.
    prompt: truyền vào để thử prompt chưa lưu (nút "Thử prompt" trong Cấu hình).
    """
    cfg = _config()
    global _offline_until
    if not opinions or not is_configured(cfg) or time.time() < _offline_until:
        return None
    try:
        return run_summary(cfg, opinions, prompt or get_prompt(kind, cfg), variables)
    except Exception as e:
        if isinstance(e, (urllib.error.URLError, TimeoutError, OSError)) and \
                not isinstance(e, urllib.error.HTTPError):
            _offline_until = time.time() + _OFFLINE_BACKOFF
        log.warning("Không dùng được AI (%s) — chuyển sang xử lý offline: %s",
                    cfg.get("ai_model"), _describe_error(e))
        return None


def run_summary(cfg, opinions: List[str], prompt: str,
                variables: Optional[Dict[str, str]] = None) -> Optional[str]:
    """Gọi AI tóm tắt; ném lỗi nếu gọi thất bại (dùng cho nút Thử prompt để hiện lỗi)."""
    variables = {"so_y_kien": len(opinions), **(variables or {})}
    listing = "\n".join(f"- {o}" for o in opinions)
    content = _chat(cfg, [{"role": "system", "content": fill_placeholders(prompt, variables)},
                          {"role": "user", "content": f"Các ý kiến của sinh viên:\n{listing}"}],
                    timeout=float(cfg.get("ai_timeout", 60)))
    summary = _clean_paragraph(content)
    if len(summary) < 20:
        log.warning("AI trả về nội dung quá ngắn, dùng cách xử lý offline")
        return None
    log.info("AI (%s) đã tóm tắt %d ý kiến", cfg["ai_model"], len(opinions))
    return summary


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
