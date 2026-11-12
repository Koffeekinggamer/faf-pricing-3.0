"""Plain text cleanup and ReportLab Paragraph escaping."""

from __future__ import annotations

import re

_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffd]")
_X000D_RE = re.compile(r"_x000D_", re.I)
_LITERAL_BR_RE = re.compile(r"<br\s*/?>", re.I)
_STRIP_TAG_RE = re.compile(r"</?\s*br\s*/?\s*>", re.I)


def clean_pdf_plain_text(text: str) -> str:
    if not text:
        return ""
    if not isinstance(text, str):
        text = str(text)
    text = _X000D_RE.sub("\n", text)
    text = _LITERAL_BR_RE.sub("\n", text)
    text = _STRIP_TAG_RE.sub("\n", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u00a0", " ")
    text = text.replace("&nbsp;", " ")
    text = _CONTROL_RE.sub("", text)
    return re.sub(r"[ \t]+\n", "\n", text).strip()


def escape_for_paragraph_markup(text: str) -> str:
    """Escape user text; real newlines become ReportLab line breaks."""
    text = clean_pdf_plain_text(text)
    text = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    return text.replace("\n", "<br/>")

