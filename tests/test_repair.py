from __future__ import annotations

import json
from datetime import datetime

from inlook.library import Library
from inlook.msgfile import read_details, read_message
from inlook.repair import run_repair
from inlook.source import MailHeader, MailRecord
from inlook.store import Store, msg_relpath
from tests.cfb import fake_mail_msg

WHEN = datetime(2026, 10, 20, 9, 0)
X500 = "/O=EXAMPLE/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=GILDONG"
HTML = "<p>안녕</p><img src='cid:image001.png@01DB0000'>"
ATTACHMENTS = [
    ("image001.png", "image001.png@01DB0000", b"png"),  # 본문이 cid: 로 가리킴
    ("image002.png", "image002.png@01DB0000", b"png"),  # 숨김 표시
    ("보고서.pdf", "", b"%PDF"),
    ("도표.png", "chart@01DB0000", b"png"),  # Content-ID 는 있지만 본문이 안 가리킴 → 진짜 첨부
]


def _msg(path, *, smtp: str = "", representing: str = ""):
    extra = {0x0C1F: X500}
    if representing:
        extra[0x5D02] = representing
    fake_mail_msg(path, subject="제목", received=WHEN, message_id="<m@example.com>",
                  sender_email=smtp, html=HTML, attachments=ATTACHMENTS, hidden=(1,), extra=extra)


def test_details_skip_inline_images_and_find_smtp(tmp_path):
    path = tmp_path / "m.msg"
    _msg(path, representing="gildong.hong@example.com")
    details = read_details(path)
    assert details.attachments == ["보고서.pdf", "도표.png"]
    assert details.sender_smtp == "gildong.hong@example.com"
    msg = read_message(path)
    assert msg.attachments == ["보고서.pdf", "도표.png"]
    assert msg.sender_email == "gildong.hong@example.com"


def test_details_without_smtp_keeps_empty(tmp_path):
    path = tmp_path / "m.msg"
    _msg(path)
    assert read_details(path).sender_smtp == ""
    assert read_message(path).sender_email == X500  # 가져오기는 있는 값이라도 남긴다 (화면에서는 숨김)


def _archive(tmp_path, **msg_kwargs):
    root = tmp_path / "archive"
    store = Store(root)
    rel = msg_relpath("<m@example.com>", WHEN)
    (root / rel).parent.mkdir(parents=True)
    _msg(root / rel, **msg_kwargs)
    header = MailHeader(entry_id="E", store_id="S", message_class="IPM.Note", subject="제목",
                        sender_name="홍길동", received=WHEN, size=1, message_id="<m@example.com>", search_key=None)
    record = MailRecord(header=header, sender_email=X500, to="", cc="", sent=WHEN, body="본문", html_body="",
                        attachments=[a[0] for a in ATTACHMENTS])
    store.add("<m@example.com>", "message_id", record, rel, ["받은 편지함"], None)
    store.close()
    return root


def _row(root):
    store = Store(root)
    attachments, sender = store.db.execute("SELECT attachments, sender_email FROM mails").fetchone()
    store.close()
    return json.loads(attachments), sender


def test_repair_dry_run_then_apply(tmp_path, capsys):
    root = _archive(tmp_path, smtp="gildong.hong@example.com")
    assert run_repair(root, dry_run=True) == 0
    out = capsys.readouterr().out
    assert "본문 이미지 2개" in out and "SMTP 주소로 바꿈 1건" in out and "바뀌는 메일 1건" in out
    assert "보고서" not in out and "gildong" not in out
    assert _row(root) == ([a[0] for a in ATTACHMENTS], X500)

    assert run_repair(root) == 0
    assert _row(root) == (["보고서.pdf", "도표.png"], "gildong.hong@example.com")
    capsys.readouterr()
    run_repair(root, dry_run=True)
    assert "바뀌는 메일 0건" in capsys.readouterr().out


def test_repair_keeps_x500_when_msg_has_no_smtp(tmp_path, capsys):
    root = _archive(tmp_path)
    run_repair(root)
    assert "SMTP 주소 없음 1건" in capsys.readouterr().out
    assert _row(root) == (["보고서.pdf", "도표.png"], X500)
    lib = Library(root)
    assert lib.search()[0].has_attachments
    assert lib.search("image001") == []  # 본문 이미지 이름으로는 검색되지 않는다
    lib.close()
