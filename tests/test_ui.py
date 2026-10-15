from __future__ import annotations

import os
from datetime import datetime

import pytest
from PySide6.QtCore import Qt

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from mailer.fake import FakeMailSource  # noqa: E402
from mailer.ui import MainWindow, SyncDialog  # noqa: E402

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
    assert "c00" in window.info.styleSheet()


def test_sync_then_list_search_and_preview(qtbot, window):
    dialog = _sync(qtbot, window)
    assert dialog.result_code == 0
    assert "저장 106건" in dialog.log.toPlainText()
    assert window.model.rowCount() == 106
    assert window.info.styleSheet() == ""

    window.search.setText("회의록")
    window.refresh_list()
    assert 0 < window.model.rowCount() < 106

    window.table.selectRow(0)
    assert "회의록" in window.preview.header.text()
    assert window.preview.open_btn.isEnabled()


def test_label_tree_filters(qtbot, window):
    _sync(qtbot, window)
    parent = window.tree.findItems("프로젝트A (10)", Qt.MatchFlag.MatchExactly | Qt.MatchFlag.MatchRecursive)
    assert parent and "하위 라벨 포함 15건" in parent[0].toolTip(0)
    found = window.tree.findItems("회의록 (5)", Qt.MatchFlag.MatchExactly | Qt.MatchFlag.MatchRecursive)
    assert found
    window.tree.setCurrentItem(found[0])
    found[0].setSelected(True)
    assert window.model.rowCount() == 5


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
