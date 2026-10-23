"""Outlook(Word·Excel) 이 만든 메일 HTML 을 Qt 의 HTML 엔진(QTextBrowser)이 알아듣는 형태로 고친다.

Qt 는 테두리 CSS 를 '굵기 모양 색' 순서(1px solid black)로만 읽는다. Word 는 'solid windowtext 1.0pt',
Excel 은 '.5pt solid windowtext' 처럼 쓰고 Windows 시스템 색 이름(windowtext)을 써서 표 테두리가 통째로 사라진다.
여기서 border 선언을 Qt 순서로 다시 쓴다. 메일 원본(.msg)은 건드리지 않고 미리보기에서만 바꾼다.

shape(): 본문 HTML 의 구조만 남긴 진단용 사본 (글자는 x 로, 링크·주석·문서 속성 제거). 내용은 남지 않는다.
"""

from __future__ import annotations

import re

# style 속성·<style> 블록 안의 border / border-top 등 (border-collapse, border-top-width 같은 세부 속성은 제외)
_BORDER = re.compile(r"(?<![\w-])(border(?:-(?:top|right|bottom|left))?)(\s*:\s*)([^;\"'}<>]+)", re.IGNORECASE)
_WIDTH = re.compile(r"^(\d*\.?\d+)(pt|px|cm|mm|in|em)?$", re.IGNORECASE)
_STYLES = {"none", "hidden", "dotted", "dashed", "solid", "double", "groove", "ridge", "inset", "outset"}
_KEYWORD_WIDTHS = {"thin": 1, "medium": 2, "thick": 3}
_SYSTEM_COLORS = {"windowtext", "windowframe", "buttontext", "auto", "currentcolor"}
_PX_PER = {"pt": 96 / 72, "px": 1, "cm": 96 / 2.54, "mm": 96 / 25.4, "in": 96, "em": 13, "": 1}


def _px(token: str) -> int | None:
    if token.lower() in _KEYWORD_WIDTHS:
        return _KEYWORD_WIDTHS[token.lower()]
    m = _WIDTH.match(token)
    if not m:
        return None
    value = float(m[1]) * _PX_PER[(m[2] or "").lower()]
    return 0 if value == 0 else max(1, round(value))  # 0.5pt 처럼 가는 선도 1px 로 보이게


def _border_value(value: str) -> str:
    important = "!important" in value.lower()
    tokens = value.replace("!important", "").replace("!IMPORTANT", "").split()
    width, style, color = None, None, None
    for token in tokens:
        low = token.lower()
        if low in _STYLES:
            style = low
        elif (px := _px(token)) is not None:
            width = px
        else:
            color = "#000000" if low in _SYSTEM_COLORS else token
    if style in (None, "none", "hidden") or width == 0:
        return "0px none" + (" !important" if important else "")
    return f"{width if width is not None else 1}px {style} {color or '#000000'}" + (" !important" if important else "")


def normalize(html: str) -> str:
    """테두리 선언을 Qt 가 읽는 순서로."""
    return _BORDER.sub(lambda m: m[1] + m[2] + _border_value(m[3]), html)


# --- 진단용 구조 사본 ---

_COMMENT = re.compile(r"<!--.*?-->", re.DOTALL)
_XML = re.compile(r"<xml>.*?</xml>", re.DOTALL | re.IGNORECASE)
_DROP_ATTR = re.compile(r"\s(href|src|alt|title|name|lang|o:\w+|v:\w+)\s*=\s*(\"[^\"]*\"|'[^']*'|[^\s>]+)", re.IGNORECASE)
_TAG_OR_TEXT = re.compile(r"(<style\b.*?</style>|<[^>]+>)|([^<]+)", re.DOTALL | re.IGNORECASE)


def shape(html: str) -> str:
    """태그·class·style 만 남긴다. 글자는 x(공백은 유지), 링크·이미지 주소·주석·문서 속성(작성자 등)은 지운다."""
    html = _XML.sub("", _COMMENT.sub("", html))
    out = []
    for m in _TAG_OR_TEXT.finditer(html):
        if m[1]:
            out.append(m[1] if m[1].lower().startswith("<style") else _DROP_ATTR.sub("", m[1]))
        else:
            out.append(re.sub(r"[^\s&;]", "x", m[2]))
    return "".join(out)
