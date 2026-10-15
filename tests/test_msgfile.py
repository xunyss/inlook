from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QBuffer, QUrl  # noqa: E402
from PySide6.QtGui import QColor, QImage  # noqa: E402

from inlook.msgfile import inline_images, read_attachments  # noqa: E402
from tests.cfb import fake_msg  # noqa: E402


def _png(color: str, size: int = 40) -> bytes:
    image = QImage(size, size, QImage.Format.Format_RGB32)
    image.fill(QColor(color))
    buf = QBuffer()
    buf.open(QBuffer.OpenModeFlag.WriteOnly)
    image.save(buf, "PNG")
    return bytes(buf.data())


def test_reads_attachments_with_content_id(tmp_path):
    big = os.urandom(5000)  # 여러 섹터에 걸치는 첨부
    path = tmp_path / "a.msg"
    fake_msg(path, [
        ("image001.png", "<image001.png@01DB0000.00000000>", _png("red")),
        ("보고서.docx", "", big),
    ])
    atts = read_attachments(path)
    assert [a.filename for a in atts] == ["image001.png", "보고서.docx"]
    assert atts[0].content_id == "image001.png@01DB0000.00000000"
    assert atts[1].data == big

    images = inline_images(path)
    assert QImage.fromData(images["image001.png@01DB0000.00000000"]).width() == 40
    assert "image001.png" in images


def test_preview_shows_cid_image(qtbot, tmp_path):
    from inlook.ui import MailBody

    path = tmp_path / "b.msg"
    fake_msg(path, [("image001.png", "image001.png@01DB", _png("blue", 64))])
    body = MailBody()
    qtbot.addWidget(body)
    body.images = inline_images(path)
    image = body.loadResource(2, QUrl("cid:image001.png@01DB"))
    assert isinstance(image, QImage) and image.width() == 64
    assert body.loadResource(2, QUrl("cid:missing@x")) is None
    assert body.loadResource(2, QUrl("http://example.com/track.png")) is None
