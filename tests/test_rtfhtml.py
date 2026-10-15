from __future__ import annotations

import os
import struct

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from inlook.rtfhtml import _PREBUF, decompress_rtf, html_from_compressed_rtf, html_from_rtf  # noqa: E402


def compress(raw: bytes) -> bytes:
    """테스트용 LZFu 압축기 (욕심쟁이 일치, 겹침 없는 참조만)."""
    dictionary = bytearray(4096)
    dictionary[:len(_PREBUF)] = _PREBUF
    write = len(_PREBUF)
    tokens: list[tuple[bool, bytes]] = []
    i = 0
    while i < len(raw):
        snapshot = bytes(dictionary)
        match = None
        for length in range(min(17, len(raw) - i), 1, -1):
            offset = snapshot.find(raw[i:i + length])
            if offset >= 0 and offset != write and not (offset <= write < offset + length) \
                    and not (write <= offset < write + length) and offset + length <= 4096:
                match = (offset, length)
                break
        if match:
            offset, length = match
            tokens.append((True, struct.pack(">H", (offset << 4) | (length - 2))))
            chunk = raw[i:i + length]
        else:
            tokens.append((False, raw[i:i + 1]))
            chunk = raw[i:i + 1]
        for b in chunk:
            dictionary[write] = b
            write = (write + 1) % 4096
        i += len(chunk)
    tokens.append((True, struct.pack(">H", write << 4)))  # 끝 표시
    body = b""
    for start in range(0, len(tokens), 8):
        group = tokens[start:start + 8]
        control = sum(1 << bit for bit, (ref, _) in enumerate(group) if ref)
        body += bytes([control]) + b"".join(data for _, data in group)
    return struct.pack("<IIII", len(body) + 12, len(raw), 0x75465A4C, 0) + body


RTF = (
    b"{\\rtf1\\ansi\\ansicpg949\\fromhtml1 \\fbidis \\deff0{\\fonttbl{\\f0\\fswiss Arial;}}\r\n"
    b"{\\*\\htmltag19 <html>}{\\*\\htmltag34 <head>}{\\*\\htmltag241 <style>p\\{color:#1F497D\\}</style>}\r\n"
    b"{\\*\\htmltag41 </head>}{\\*\\htmltag50 <body>}\\htmlrtf {\\htmlrtf0 \r\n"
    b"{\\*\\htmltag64 <p>}\\htmlrtf {\\htmlrtf0 \\'b0\\'a1\\'b3\\'aa \\u54620?\\u44544? text\\htmlrtf\\par\\htmlrtf0 \r\n"
    b"{\\*\\htmltag72 </p>}\r\n"
    b"{\\*\\htmltag84 <![if !vml]>}{\\*\\htmltag84 <img src=\"cid:image001.png@01DB\">}{\\*\\htmltag84 <![endif]>}\r\n"
    b"\\htmlrtf }\\htmlrtf0 {\\*\\htmltag58 </body>}{\\*\\htmltag27 </html>}}"
)


def test_decompress_roundtrip_with_references():
    raw = RTF * 3  # 반복이 있어 참조 토큰이 생긴다
    packed = compress(raw)
    assert len(packed) < len(raw)
    assert decompress_rtf(packed) == raw


def test_uncompressed_mela():
    packed = struct.pack("<IIII", len(RTF) + 12, len(RTF), 0x414C454D, 0) + RTF
    assert decompress_rtf(packed) == RTF


def test_html_is_extracted_from_encapsulated_rtf():
    html = html_from_compressed_rtf(compress(RTF))
    assert html.startswith("<html><head><style>p{color:#1F497D}</style>")
    assert "가나 한글 text" in html
    assert '<![if !vml]><img src="cid:image001.png@01DB"><![endif]>' in html
    assert "Arial" not in html and "fbidis" not in html


def test_plain_rtf_is_not_html():
    assert html_from_rtf(b"{\\rtf1\\ansi Hello\\par}") is None


def test_vml_conditional_image_is_requested_by_viewer(qtbot):
    from PySide6.QtCore import QUrl

    from inlook.ui import MailBody

    requested: list[str] = []

    class Spy(MailBody):
        def loadResource(self, kind, url: QUrl):  # noqa: N802
            requested.append(url.toString())
            return super().loadResource(kind, url)

    body = Spy()
    qtbot.addWidget(body)
    body.setHtml(html_from_rtf(RTF))
    body.show()
    qtbot.wait(50)
    assert "cid:image001.png@01DB" in requested


def test_meeting_request_html_comes_from_msg_and_is_cached(tmp_path):
    from inlook.collect import run_collect
    from inlook.fake import FakeMailSource
    from inlook.library import Library
    from inlook.msgfile import read_html
    from tests.cfb import write_cfb

    meeting_msg = {
        "__substg1.0_0037001F": "회의".encode("utf-16-le"),
        "__substg1.0_10090102": compress(RTF),
        "__properties_version1.0": b"\0" * 32,
    }
    path = tmp_path / "m.msg"
    write_cfb(path, meeting_msg)
    assert "가나 한글 text" in read_html(path)

    # COM 이 HTMLBody 를 주지 않은 회의 요청: 수집 때 저장한 .msg 에서 채운다
    source = FakeMailSource()
    target = next(r for r in source._by_id.values() if r.header.message_class.startswith("IPM.Schedule"))
    target.html_body = ""
    source.save_msg = lambda record, p: write_cfb(p, meeting_msg)
    root = tmp_path / "archive"
    run_collect(lambda: source, root)
    lib = Library(root)
    row = lib.search("분기 리뷰")[0]
    assert "가나 한글 text" in lib.get(row.id).html_body

    # 이미 HTML 없이 저장된 메일: 미리보기 때 채우고 DB 에 남긴다
    lib.db.execute("UPDATE mails SET html_body = '' WHERE id = ?", (row.id,))
    lib.db.commit()
    filled = lib.fill_html(lib.get(row.id))
    assert "cid:image001.png@01DB" in filled.html_body
    assert lib.get(row.id).html_body == filled.html_body
    lib.close()
