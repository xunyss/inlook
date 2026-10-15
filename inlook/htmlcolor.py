"""어두운 배경에서 메일 본문을 읽을 수 있게 HTML 색을 보정한다 (Outlook 다크 모드와 같은 방식).

- 글자색이 어두우면 밝기를 뒤집어 밝게 (색상은 유지: 짙은 파랑 → 밝은 파랑)
- 배경색이 밝으면 밝기를 뒤집어 어둡게
색이 지정되지 않은 글자·배경은 뷰어의 기본값(어두운 배경·밝은 글자)을 따른다.
"""

from __future__ import annotations

import re

from PySide6.QtGui import QColor

_TEXT_PROPS = {"color"}
_BG_PROPS = {"background", "background-color"}

# style 속성·<style> 블록 안의 CSS 선언
_CSS_DECL = re.compile(r"(?<![\w-])(color|background-color|background)(\s*:\s*)([^;\"'}<>]+)", re.IGNORECASE)
# HTML 속성: <font color=..>, <td bgcolor=..>
_ATTR = re.compile(r"(?<![\w-])(color|bgcolor)(\s*=\s*)([\"']?)([^\"'\s>]+)\3", re.IGNORECASE)
# CSS 값 안의 색 토큰
_TOKEN = re.compile(r"#[0-9a-fA-F]{3,8}\b|rgba?\([^)]*\)|\b[a-zA-Z]+\b")
_RGB = re.compile(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)")
_NOT_COLORS = {"none", "inherit", "initial", "unset", "transparent", "currentcolor", "url", "repeat",
               "no", "top", "left", "right", "bottom", "center", "fixed", "scroll", "auto", "important"}


def _parse(token: str) -> QColor | None:
    if token.lower() in _NOT_COLORS:
        return None
    m = _RGB.match(token)
    color = QColor(int(m[1]), int(m[2]), int(m[3])) if m else QColor(token)
    return color if color.isValid() and color.alpha() > 0 else None


def _flip(color: QColor) -> str:
    h, s, l, _ = color.getHslF()
    out = QColor.fromHslF(max(h, 0.0), s, 1.0 - l)
    return out.name()


def _adjust(token: str, *, text: bool) -> str:
    color = _parse(token)
    if color is None:
        return token
    light = color.lightnessF()
    if text and light < 0.5:
        return _flip(color)
    if not text and light > 0.5:
        return _flip(color)
    return token


def _adjust_value(value: str, *, text: bool) -> str:
    return _TOKEN.sub(lambda m: _adjust(m.group(0), text=text), value)


def darken_html(html: str) -> str:
    def css(m: re.Match) -> str:
        prop = m.group(1).lower()
        return m.group(1) + m.group(2) + _adjust_value(m.group(3), text=prop in _TEXT_PROPS)

    def attr(m: re.Match) -> str:
        text = m.group(1).lower() == "color"
        return m.group(1) + m.group(2) + m.group(3) + _adjust(m.group(4), text=text) + m.group(3)

    return _ATTR.sub(attr, _CSS_DECL.sub(css, html))
