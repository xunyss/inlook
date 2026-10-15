from PySide6.QtGui import QColor

from mailer.htmlcolor import darken_html


def _light(name: str) -> float:
    return QColor(name).lightnessF()


def _colors(html: str) -> list[str]:
    import re

    return re.findall(r"#[0-9a-f]{6}", html)


def test_dark_text_becomes_light_keeping_hue():
    out = darken_html('<p style="color:#1F497D">본문</p>')
    (color,) = _colors(out)
    assert _light(color) > 0.5
    assert abs(QColor(color).hslHueF() - QColor("#1F497D").hslHueF()) < 0.02


def test_light_text_is_kept():
    html = '<span style="color: #FFC000">경고</span>'
    assert darken_html(html) == html


def test_light_background_becomes_dark():
    out = darken_html('<td bgcolor="#FFFFFF" style="background-color: rgb(255, 255, 204)">x</td>')
    colors = _colors(out)
    assert len(colors) == 2 and all(_light(c) < 0.5 for c in colors)


def test_font_color_attribute_and_named_colors():
    out = darken_html('<font color="navy">a</font><p style="color:black;font-size:10pt">b</p>')
    colors = _colors(out)
    assert len(colors) == 2 and all(_light(c) > 0.5 for c in colors)
    assert "font-size:10pt" in out


def test_non_colors_untouched():
    html = '<div style="background: url(x.png) no-repeat; color: inherit">t</div>'
    assert darken_html(html) == html
    assert darken_html("색깔 설명 텍스트") == "색깔 설명 텍스트"
