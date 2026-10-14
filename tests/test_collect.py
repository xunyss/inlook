from __future__ import annotations

import sqlite3
from dataclasses import replace
from datetime import datetime, timedelta
from pathlib import Path

import pytest

from mailer.collect import run_collect
from mailer.fake import FakeMailSource
from mailer.store import DB_NAME, Store

NOW = datetime(2026, 10, 14, 9, 30)


def _db(root: Path) -> sqlite3.Connection:
    return sqlite3.connect(root / DB_NAME)


def _target_mails(source: FakeMailSource) -> int:
    excluded = {"지운 편지함"}
    return sum(len(v) for k, v in source._items.items() if k[0] not in excluded)


def test_first_sync_saves_everything_oldest_first(tmp_path, capsys):
    source = FakeMailSource(NOW)
    assert run_collect(lambda: source, tmp_path) == 0

    db = _db(tmp_path)
    assert db.execute("SELECT COUNT(*) FROM mails").fetchone()[0] == _target_mails(source)
    received = [r[0] for r in db.execute("SELECT received FROM mails ORDER BY id")]
    assert received == sorted(received)  # 곧 사라질 오래된 메일부터
    assert db.execute("SELECT status FROM syncs").fetchall() == [("ok",)]

    paths = [r[0] for r in db.execute("SELECT msg_path FROM mails")]
    assert all((tmp_path / p).is_file() for p in paths)
    assert len(set(paths)) == len(paths)
    for p in paths:
        name = Path(p).name
        assert len(name) == len("20261014-093215_3f9a1c2b7e04.msg")
        assert p.startswith("msg/")


def test_labels_are_outlook_folder_paths(tmp_path):
    source = FakeMailSource(NOW)
    run_collect(lambda: source, tmp_path)
    labels = {r[0] for r in _db(tmp_path).execute("SELECT name FROM labels")}
    assert "받은 편지함/프로젝트A/회의록" in labels
    assert "보낸 편지함" in labels
    assert "지운 편지함" not in labels


def test_second_sync_is_incremental_and_keeps_user_labels(tmp_path):
    source = FakeMailSource(NOW)
    run_collect(lambda: source, tmp_path)
    store = Store(tmp_path)
    key = store.db.execute("SELECT key FROM mails WHERE received >= ? LIMIT 1", ("2026-10-14",)).fetchone()[0]
    # 사용자가 mailer 에서 라벨을 모두 뗐다
    store.db.execute("DELETE FROM mail_labels WHERE mail_id = (SELECT id FROM mails WHERE key = ?)", (key,))
    store.db.commit()
    before = store.count()
    store.close()

    new = source._mail(("받은 편지함",), 0, "새로 온 메일", received=NOW + timedelta(hours=1))
    assert run_collect(lambda: source, tmp_path) == 0

    store = Store(tmp_path)
    assert store.count() == before + 1
    assert store.has(new.header.message_id)
    assert store.labels_of(key) == []  # 동기화가 라벨을 되살리지 않는다
    store.close()


def test_same_mail_in_two_folders_is_one_mail_with_two_labels(tmp_path):
    source = FakeMailSource(NOW)
    original = source._items[("받은 편지함",)][0]
    copy = replace(original, header=replace(original.header, entry_id="FAKECOPY"))
    source._items[("보관",)].append(copy)
    source._by_id["FAKECOPY"] = copy
    run_collect(lambda: source, tmp_path)

    store = Store(tmp_path)
    assert store.labels_of(original.header.message_id) == ["받은 편지함", "보관"]
    store.close()


def test_dry_run_saves_nothing(tmp_path):
    assert run_collect(lambda: FakeMailSource(NOW), tmp_path, dry_run=True) == 0
    assert Store(tmp_path).count() == 0
    assert not (tmp_path / "msg").exists()


def test_failed_sync_is_not_used_as_cutoff(tmp_path):
    flaky = FakeMailSource(NOW)
    calls = {"n": 0}
    real_save = flaky.save_msg

    def save(record, path):
        calls["n"] += 1
        if calls["n"] == 3:
            raise OSError("disk hiccup")
        real_save(record, path)

    flaky.save_msg = save
    assert run_collect(lambda: flaky, tmp_path) == 1
    store = Store(tmp_path)
    assert store.last_ok_sync() is None
    missing = _target_mails(flaky) - store.count()
    assert missing == 1
    store.close()

    # 다시 돌리면 빠진 1건을 채운다
    assert run_collect(lambda: FakeMailSource(NOW), tmp_path) == 0
    assert Store(tmp_path).count() == _target_mails(flaky)


def test_consecutive_failures_stop_the_run(tmp_path, capsys):
    assert run_collect(lambda: FakeMailSource(NOW, fail_save=True), tmp_path) == 1
    out = capsys.readouterr().out
    assert "연속 5건 실패" in out
    assert out.count("[E109]") >= 5
    assert Store(tmp_path).count() == 0


def test_ctrl_c_stops_cleanly_and_resumes(tmp_path, capsys):
    source = FakeMailSource(NOW)
    real_save = source.save_msg
    calls = {"n": 0}

    def save(record, path):
        calls["n"] += 1
        if calls["n"] == 10:
            raise KeyboardInterrupt
        real_save(record, path)

    source.save_msg = save
    assert run_collect(lambda: source, tmp_path) == 130
    assert "[STOP]" in capsys.readouterr().out
    store = Store(tmp_path)
    assert store.count() == 9
    assert store.db.execute("SELECT status FROM syncs").fetchall() == [("stopped",)]
    store.close()

    assert run_collect(lambda: FakeMailSource(NOW), tmp_path) == 0
    assert Store(tmp_path).count() == _target_mails(source)


def test_output_has_no_names_or_subjects(tmp_path, capsys):
    source = FakeMailSource(NOW, fail_save=True)
    run_collect(lambda: source, tmp_path)
    run_collect(lambda: FakeMailSource(NOW), tmp_path)
    out = capsys.readouterr().out
    for folder in source.folders():
        for name in folder.path:
            assert name not in out
    for word in ("주간 업무", "프로젝트A", "홍길동", "example.com"):
        assert word not in out


def test_rejects_root_too_long_for_windows(tmp_path, capsys):
    root = tmp_path / ("x" * 240)
    assert run_collect(lambda: FakeMailSource(NOW), root) == 2
    assert "[E111]" in capsys.readouterr().out
