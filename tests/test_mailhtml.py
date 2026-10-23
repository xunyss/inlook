from __future__ import annotations

from inlook.htmlcolor import darken_html
from inlook.mailhtml import normalize, shape


def test_word_and_excel_borders_become_qt_order():
    assert normalize("<td style='border:solid windowtext 1.0pt;border-left:none'>") == \
        "<td style='border:1px solid #000000;border-left:0px none'>"
    assert normalize(".xl65 {border:.5pt solid windowtext; text-align:center;}") == \
        ".xl65 {border:1px solid #000000; text-align:center;}"
    assert normalize("border-bottom:2.25pt double #4472C4") == "border-bottom:3px double #4472C4"
    assert normalize("border:none") == "border:0px none"


def test_other_border_properties_untouched():
    text = "border-collapse:collapse;mso-border-alt:solid windowtext .5pt;border-top-width:1px"
    assert normalize(text) == text


def test_dark_theme_lightens_border_colors():
    out = darken_html(normalize("<td style='border:solid windowtext 1.0pt'>"))
    assert "#000000" not in out and "1px solid #ffffff" in out


def test_shape_keeps_structure_only():
    html = (
        "<html><head><!--[if gte mso 9]><xml><o:Author>홍길동</o:Author></xml><![endif]-->"
        "<style>p.MsoNormal {font-size:10.0pt}</style></head>"
        "<body><p class=MsoNormal style='font-size:9pt'>고객사 기밀 123</p>"
        "<a href='https://intra.example.com/x'>링크</a><img src='cid:image001.png' alt='로고'></body></html>"
    )
    out = shape(html)
    for secret in ("홍길동", "고객사", "기밀", "123", "intra.example.com", "image001", "로고", "링크"):
        assert secret not in out
    assert "<p class=MsoNormal style='font-size:9pt'>xxx xx xxx</p>" in out
    assert "p.MsoNormal {font-size:10.0pt}" in out
