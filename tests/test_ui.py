from __future__ import annotations

import os
from datetime import datetime

import pytest
from PySide6.QtCore import Qt

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from inlook.fake import FakeMailSource  # noqa: E402
from inlook import theme  # noqa: E402
from inlook.ui import ROLE_COUNT, MailModel, MainWindow, SyncDialog, friendly_date, split_people  # noqa: E402

NOW = datetime(2026, 10, 14, 9, 30)


@pytest.fixture
def window(qtbot, tmp_path):
    w = MainWindow(tmp_path, lambda: FakeMailSource(NOW), tmp_path / "logs")
    qtbot.addWidget(w)
    yield w
    w.lib.close()


def _sync(qtbot, window):
    dialog = SyncDialog(window, window.make_source, window.root, window.log_dir)
    qtbot.waitUntil(lambda: not dialog.running, timeout=10000)
    window.refresh_all()
    return dialog


def test_empty_archive_opens_and_warns(window):
    assert window.model.rowCount() == 0
    assert "없음" in window.info.text()
    assert window.info.objectName() == "stale"  # 빨간 경고


def test_sync_then_list_search_and_preview(qtbot, window):
    dialog = _sync(qtbot, window)
    assert dialog.result_code == 0
    assert "저장 106건" in dialog.log.toPlainText()
    assert window.model.rowCount() == 106
    assert window.info.objectName() == ""
    assert window.count_label.text() == "106건"

    window.search.setText("회의록")
    window.refresh_list()
    assert 0 < window.model.rowCount() < 106

    window.table.selectRow(0)
    assert "회의록" in window.preview.title.text()
    assert window.open_act.isEnabled() and window.label_act.isEnabled()
    assert "검색 결과" in window.count_label.text()


def test_label_tree_filters(qtbot, window):
    _sync(qtbot, window)
    flags = Qt.MatchFlag.MatchExactly | Qt.MatchFlag.MatchRecursive
    parent = window.tree.findItems("프로젝트A", flags)
    assert parent and parent[0].data(0, ROLE_COUNT) == 10
    assert "하위 라벨 포함 15건" in parent[0].toolTip(0)
    found = window.tree.findItems("회의록", flags)
    assert found and found[0].data(0, ROLE_COUNT) == 5
    assert not window.table.isColumnHidden(MailModel.LABEL)
    window.tree.setCurrentItem(found[0])
    found[0].setSelected(True)
    assert window.model.rowCount() == 5
    assert window.table.isColumnHidden(MailModel.LABEL)  # 라벨 하나를 고르면 라벨 열은 숨긴다


def test_sync_can_be_cancelled(qtbot, window):
    dialog = SyncDialog(window, window.make_source, window.root, window.log_dir)
    dialog._stop()
    qtbot.waitUntil(lambda: not dialog.running, timeout=10000)
    assert dialog.result_code in (0, 130)


def test_delete_from_ui(qtbot, window, monkeypatch):
    from PySide6.QtWidgets import QMessageBox

    _sync(qtbot, window)
    monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: QMessageBox.StandardButton.Yes)
    window.table.selectRow(0)
    window.delete_selected()
    assert window.model.rowCount() == 105
    assert "105" in window.info.text()


def _actions(menu):
    for action in menu.actions():
        if action.menu():
            yield from _actions(action.menu())
        elif action.text():
            yield action


def test_change_label_from_submenu(qtbot, window, monkeypatch):
    from PySide6.QtWidgets import QInputDialog, QMenu

    _sync(qtbot, window)
    window.table.selectRow(0)
    mail_id = window.selected_ids()[0]
    before = window.lib.get(mail_id).labels

    menu = window._label_submenu(QMenu(), [mail_id])
    actions = {a.text(): a for a in _actions(menu)}
    assert {"받은 편지함", "프로젝트A", "회의록", "보낸 편지함", "새 라벨…"} <= set(actions)
    assert [a.text() for a in _actions(menu) if a.isChecked()] == [b.rsplit("/", 1)[-1] for b in before]

    actions["회의록"].trigger()
    assert window.lib.get(mail_id).labels == ["받은 편지함/프로젝트A/회의록"]

    monkeypatch.setattr(QInputDialog, "getText", lambda *a, **k: ("새 라벨", True))
    {a.text(): a for a in _actions(window._label_submenu(QMenu(), [mail_id]))}["새 라벨…"].trigger()
    assert window.lib.get(mail_id).labels == ["새 라벨"]


def test_friendly_date():
    from datetime import date

    today = date(2026, 10, 23)
    assert friendly_date("2026-10-23 14:51:09", today) == "오늘 14:51"
    assert friendly_date("2026-10-22 09:05:00", today) == "어제 09:05"
    assert friendly_date("2026-10-21 18:41:00", today) == "10-21 (수)"
    assert friendly_date("2025-11-03 08:00:00", today) == "2025-11-03"
    assert friendly_date("", today) == ""


def test_split_people():
    assert split_people("A; B, C ;; ") == ["A", "B", "C"]


def test_theme_toggle_is_remembered(qtbot, tmp_path):
    settings = tmp_path / "settings.json"
    theme.save_mode(settings, theme.LIGHT)
    w = MainWindow(tmp_path / "a", lambda: FakeMailSource(NOW), None, settings)
    qtbot.addWidget(w)
    assert w.mode == theme.LIGHT and w.theme_act.text() == "어둡게"
    w.toggle_theme()
    assert w.mode == theme.DARK and w.theme_act.text() == "밝게"
    assert theme.load_mode(settings) == theme.DARK
    w.lib.close()
    settings.write_text("not json")
    assert theme.load_mode(settings) in (theme.LIGHT, theme.DARK)  # 깨진 설정 파일은 무시


def test_recipients_are_collapsed(qtbot, window):
    from dataclasses import replace

    _sync(qtbot, window)
    window.table.selectRow(0)
    mail = replace(window.preview._mail, to="; ".join(f"사람{i}" for i in range(10)), cc="")
    window.preview.show_mail(mail)
    assert "외 7명" in window.preview.recipients.text() and "사람9" not in window.preview.recipients.text()
    window.preview._toggle_recipients("more")
    assert "사람9" in window.preview.recipients.text()


def test_multi_select_shows_count(qtbot, window):
    _sync(qtbot, window)
    window.table.selectAll()
    assert "106건 선택됨" in window.preview.title.text()
    assert not window.open_act.isEnabled() and window.delete_act.isEnabled()


def test_list_headers_and_row_hover(qtbot, window):
    from PySide6.QtCore import QPointF
    from PySide6.QtGui import QMouseEvent

    _sync(qtbot, window)
    headers = [window.model.headerData(i, Qt.Orientation.Horizontal) for i in range(window.model.columnCount())]
    assert headers == ["보낸 사람", "제목", "라벨", "첨부", "일시"]
    window.show()
    y = window.table.rowViewportPosition(3) + 5
    move = QMouseEvent(QMouseEvent.Type.MouseMove, QPointF(50, y), QPointF(50, y),
                       Qt.MouseButton.NoButton, Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier)
    window.table.mouseMoveEvent(move)
    assert window.table.hover_row == 3  # 칸이 아니라 줄 단위
    window.table.leaveEvent(None)
    assert window.table.hover_row == -1
