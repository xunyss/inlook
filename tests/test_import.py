from __future__ import annotations

from datetime import datetime

from inlook.collect import run_collect
from inlook.fake import FakeMailSource
from inlook.importer import run_import
from inlook.library import Library
from inlook.msgfile import read_message
from tests.cfb import fake_mail_msg

NOW = datetime(2026, 10, 14, 9, 30)


def _backup(tmp_path):
    src = tmp_path / "backup"
    (src / "고객A").mkdir(parents=True)
    (src / "고객B" / "회의").mkdir(parents=True)
    fake_mail_msg(src / "루트 메일.msg", subject="루트 메일", received=datetime(2026, 3, 13, 13, 48),
                  message_id="<root@example.com>", html="<p style='color:#1F497D'>안녕</p>")
    fake_mail_msg(src / "고객A" / "a1.msg", subject="고객A 메일", received=datetime(2026, 5, 26, 18, 59),
                  message_id="<a1@example.com>", attachments=[("보고서.pdf", "", b"%PDF-fake")])
    fake_mail_msg(src / "고객B" / "회의" / "b1.msg", subject="회의 메일", received=datetime(2026, 6, 2, 9, 0),
                  search_key=bytes.fromhex("00AB12"))
    # 같은 메일이 두 폴더에 있음
    fake_mail_msg(src / "고객B" / "a1 사본.msg", subject="고객A 메일", received=datetime(2026, 5, 26, 18, 59),
                  message_id="<a1@example.com>")
    # Message-ID·SearchKey 둘 다 없음
    fake_mail_msg(src / "고객A" / "noid.msg", subject="아이디 없음", received=datetime(2026, 4, 1, 10, 0))
    (src / "고객A" / "broken.msg").write_bytes(b"not an ole file")
    (src / "고객A" / "메모.txt").write_text("무시")
    return src


def test_read_message_properties(tmp_path):
    path = tmp_path / "m.msg"
    fake_mail_msg(path, subject="제목", received=datetime(2026, 3, 13, 13, 48, 5), message_id=" <x@example.com> ",
                  search_key=b"\x01\xab", html="<meta charset=euc-kr><p>한글</p>", html_charset="euc-kr")
    msg = read_message(path)
    assert msg.subject == "제목" and msg.sender_name == "홍길동"
    assert msg.received == datetime(2026, 3, 13, 13, 48, 5)
    assert msg.message_id == "<x@example.com>" and msg.search_key == "01AB"
    assert "한글" in msg.html_body and msg.body == "본문"


def test_import_labels_dedup_and_files(tmp_path, capsys):
    src = _backup(tmp_path)
    root = tmp_path / "archive"
    assert run_import(src, root) == 1  # broken.msg 1건 실패
    out = capsys.readouterr().out
    assert "가져옴 4건" in out and "실패 1건" in out
    for secret in ("루트 메일", "고객A", "고객B", "회의 메일"):
        assert secret not in out

    lib = Library(root)
    labels = lib.label_counts()
    assert labels == {"받은 편지함": 1, "받은 편지함/고객A": 2, "받은 편지함/고객B": 1, "받은 편지함/고객B/회의": 1}
    a1 = lib.get(lib.search("고객A 메일")[0].id)
    assert a1.labels == ["받은 편지함/고객A", "받은 편지함/고객B"]
    assert a1.attachments == ["보고서.pdf"] and a1.msg_path.is_file()
    assert a1.msg_path.read_bytes() == (src / "고객A" / "a1.msg").read_bytes()
    assert a1.received.startswith("2026-05-26 18:59")
    assert lib.get(lib.search("루트 메일")[0].id).html_body
    assert lib.db.execute("SELECT COUNT(*) FROM syncs").fetchone()[0] == 0  # 동기화 기준 시각에 영향 없음
    lib.close()

    # 다시 돌려도 중복으로 들어가지 않는다
    assert run_import(src, root) == 1
    assert "가져옴 0건, 이미 있음 4건" in capsys.readouterr().out


def test_import_skips_mail_already_collected_and_deleted(tmp_path):
    root = tmp_path / "archive"
    run_collect(lambda: FakeMailSource(NOW), root)
    lib = Library(root)
    existing = lib.get(lib.search("주간 업무 공유 #3")[0].id)
    message_id = lib.db.execute("SELECT message_id FROM mails WHERE id = ?", (existing.id,)).fetchone()[0]
    before = lib.count()
    victim = lib.search("News Letter 2")[0].id
    victim_mid = lib.db.execute("SELECT message_id FROM mails WHERE id = ?", (victim,)).fetchone()[0]
    lib.delete_mails([victim])

    src = tmp_path / "backup"
    src.mkdir()
    fake_mail_msg(src / "dup.msg", subject="x", received=NOW, message_id=message_id)
    fake_mail_msg(src / "deleted.msg", subject="y", received=NOW, message_id=victim_mid)
    fake_mail_msg(src / "new.msg", subject="새 메일", received=NOW, message_id="<new@example.com>")
    assert run_import(src, root) == 0
    assert lib.count() == before - 1 + 1
    assert lib.get(existing.id).labels == existing.labels  # 수집된 메일의 라벨은 그대로
    lib.close()


def test_dry_run_and_bad_paths(tmp_path, capsys):
    src = _backup(tmp_path)
    root = tmp_path / "archive"
    assert run_import(src, root, dry_run=True) == 0
    assert "새 메일 4건" in capsys.readouterr().out
    assert Library(root).count() == 0
    assert run_import(tmp_path / "없음", root) == 2
    assert run_import(root, root) == 2


def test_parent_label_can_be_empty(tmp_path):
    src = _backup(tmp_path)
    root = tmp_path / "archive"
    run_import(src, root, parent_label="")
    assert set(Library(root).label_counts()) == {"가져온 메일", "고객A", "고객B", "고객B/회의"}
