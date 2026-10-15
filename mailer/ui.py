"""PySide6 데스크톱 뷰어: 검색, 라벨, 미리보기, 동기화.

화면 구성: [라벨 트리] | [메일 목록] | [미리보기]
동기화는 작업 스레드에서 collect 엔진을 돌린다. COM 은 그 스레드 안에서 초기화한다
(pythoncom.CoInitialize). UI 스레드에서는 Outlook 에 접근하지 않는다.
"""

from __future__ import annotations

import html
import sys
import threading
import traceback
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QFont, QFontDatabase, QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .collect import run_collect
from .library import ALL, UNLABELED, Library, MailRow
from .source import MailSource

SYNC_WARN_DAYS = 21  # 보존 기간(약 1개월)보다 넉넉히 앞에서 경고
ROLE_LABEL = Qt.ItemDataRole.UserRole


# ---------------------------------------------------------------- 메일 목록


class MailModel(QAbstractTableModel):
    HEADERS = ("수신", "보낸 사람", "제목", "라벨", "첨부")

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[MailRow] = []

    def set_rows(self, rows: list[MailRow]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):  # noqa: N802
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self.HEADERS[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            col = index.column()
            if col == 0:
                return row.received[:16]
            if col == 1:
                return row.sender
            if col == 2:
                return row.subject or "(제목 없음)"
            if col == 3:
                return ", ".join(_short(label) for label in row.labels)
            if col == 4:
                return "📎" if row.has_attachments else ""
        if role == Qt.ItemDataRole.ToolTipRole and index.column() == 3:
            return "\n".join(row.labels)
        return None


def _short(label: str) -> str:
    return label.rsplit("/", 1)[-1]


# ---------------------------------------------------------------- 동기화


class SyncWorker(QThread):
    text = Signal(str)
    progress = Signal(int, int)
    finished_with = Signal(int)

    def __init__(self, make_source: Callable[[], MailSource], root: Path, log_dir: Path | None) -> None:
        super().__init__()
        self.make_source = make_source
        self.root = root
        self.log_dir = log_dir
        self.cancel = threading.Event()

    def run(self) -> None:
        com = None
        if sys.platform == "win32":
            import pythoncom  # type: ignore[import-not-found]

            com = pythoncom
            com.CoInitialize()
        code = 2
        try:
            code = run_collect(
                self.make_source,
                self.root,
                log_dir=self.log_dir,
                echo=self.text.emit,
                progress=self.progress.emit,
                cancel=self.cancel,
            )
        except Exception:
            self.text.emit("\n[FAIL] 예기치 못한 오류\n" + traceback.format_exc())
        finally:
            if com is not None:
                com.CoUninitialize()
            self.finished_with.emit(code)


class SyncDialog(QDialog):
    def __init__(self, parent: QWidget, make_source, root: Path, log_dir: Path | None) -> None:
        super().__init__(parent)
        self.setWindowTitle("동기화")
        self.resize(820, 520)
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, False)
        self.running = True
        self.result_code: int | None = None

        self.log = QPlainTextEdit(readOnly=True)
        self.log.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        self.log.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.bar = QProgressBar()
        self.bar.setRange(0, 0)  # 목록 조회 중에는 진행률을 모른다
        self.status = QLabel("Outlook 에서 새 메일을 찾는 중… 창을 닫지 말고 기다려 주세요.")
        self.stop_btn = QPushButton("중단")
        self.stop_btn.clicked.connect(self._stop)
        self.close_btn = QPushButton("닫기")
        self.close_btn.setEnabled(False)
        self.close_btn.clicked.connect(self.accept)

        buttons = QHBoxLayout()
        buttons.addWidget(self.status, 1)
        buttons.addWidget(self.stop_btn)
        buttons.addWidget(self.close_btn)
        layout = QVBoxLayout(self)
        layout.addWidget(self.log, 1)
        layout.addWidget(self.bar)
        layout.addLayout(buttons)

        self.worker = SyncWorker(make_source, root, log_dir)
        self.worker.text.connect(self._append)
        self.worker.progress.connect(self._progress)
        self.worker.finished_with.connect(self._done)
        self.worker.start()

    def _append(self, text: str) -> None:
        self.log.moveCursor(QTextCursor.MoveOperation.End)
        self.log.insertPlainText(text)
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _progress(self, done: int, total: int) -> None:
        self.bar.setRange(0, max(total, 1))
        self.bar.setValue(done)
        self.status.setText(f"새 메일 저장 중 {done}/{total}")

    def _stop(self) -> None:
        self.worker.cancel.set()
        self.stop_btn.setEnabled(False)
        self.status.setText("지금 저장 중인 메일까지 마치고 멈춥니다…")

    def _done(self, code: int) -> None:
        self.worker.wait()
        self.running = False
        self.result_code = code
        self.bar.setRange(0, 1)
        self.bar.setValue(1)
        self.stop_btn.setEnabled(False)
        self.close_btn.setEnabled(True)
        self.status.setText({0: "완료", 1: "일부 실패 — 다시 동기화하면 빠진 것만 재시도합니다",
                             130: "중단됨 — 다시 동기화하면 이어서 합니다"}.get(code, "실패 — 로그를 확인하세요"))

    def reject(self) -> None:  # Esc 로 닫히지 않게
        if not self.running:
            super().reject()

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.running:
            event.ignore()
        else:
            super().closeEvent(event)


# ---------------------------------------------------------------- 미리보기


class Preview(QWidget):
    def __init__(self, open_msg: Callable[[], None], edit_labels: Callable[[], None]) -> None:
        super().__init__()
        self.header = QLabel()
        self.header.setTextFormat(Qt.TextFormat.RichText)
        self.header.setWordWrap(True)
        self.header.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.open_btn = QPushButton(".msg 열기")
        self.open_btn.clicked.connect(open_msg)
        self.label_btn = QPushButton("라벨…")
        self.label_btn.clicked.connect(edit_labels)
        self.body = QTextBrowser()
        self.body.setOpenLinks(False)  # 링크를 눌러도 이 창에서 이동하지 않는다
        self.body.anchorClicked.connect(self._link)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(self.label_btn)
        buttons.addWidget(self.open_btn)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.header)
        layout.addLayout(buttons)
        layout.addWidget(self.body, 1)
        self.clear()

    def clear(self) -> None:
        self.header.setText("")
        self.body.clear()
        self.open_btn.setEnabled(False)
        self.label_btn.setEnabled(False)

    def show_mail(self, mail) -> None:
        def row(name: str, value: str) -> str:
            return f"<tr><td style='color:gray;padding-right:8px'>{name}</td><td>{html.escape(value)}</td></tr>"

        sender = mail.sender_name + (f" <{mail.sender_email}>" if mail.sender_email else "")
        rows = [row("보낸 사람", sender), row("받은 시각", mail.received)]
        if mail.to:
            rows.append(row("받는 사람", mail.to))
        if mail.cc:
            rows.append(row("참조", mail.cc))
        if mail.attachments:
            rows.append(row("첨부", ", ".join(mail.attachments)))
        rows.append(row("라벨", ", ".join(mail.labels) or "(없음)"))
        self.header.setText(
            f"<div style='font-size:14pt;font-weight:bold'>{html.escape(mail.subject or '(제목 없음)')}</div>"
            f"<table>{''.join(rows)}</table>"
        )
        if mail.html_body:
            self.body.setHtml(mail.html_body)
        else:
            self.body.setPlainText(mail.body)
        self.open_btn.setEnabled(mail.msg_path.exists())
        self.label_btn.setEnabled(True)

    def _link(self, url: QUrl) -> None:
        if url.scheme() in ("http", "https", "mailto"):
            QDesktopServices.openUrl(url)


# ---------------------------------------------------------------- 메인 창


class MainWindow(QMainWindow):
    def __init__(self, root: Path, make_source: Callable[[], MailSource], log_dir: Path | None) -> None:
        super().__init__()
        self.root = root
        self.make_source = make_source
        self.log_dir = log_dir
        self.lib = Library(root)
        self.label_filter: str | None = ALL
        self.setWindowTitle(f"mailer — {root}")
        self.resize(1400, 860)

        # 위: 검색 + 동기화
        self.search = QLineEdit(placeholderText="검색 (제목·보낸 사람·받는 사람·본문·첨부 파일명, 띄어쓰기로 여러 단어)")
        self.search.setClearButtonEnabled(True)
        self._debounce = QTimer(self, singleShot=True, interval=250)
        self._debounce.timeout.connect(self.refresh_list)
        self.search.textChanged.connect(self._debounce.start)
        self.sync_btn = QPushButton("동기화")
        self.sync_btn.clicked.connect(self.sync)
        self.info = QLabel()
        top = QHBoxLayout()
        top.addWidget(self.search, 1)
        top.addWidget(self.info)
        top.addWidget(self.sync_btn)

        # 왼쪽: 라벨
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemSelectionChanged.connect(self._label_selected)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_menu)

        # 가운데: 목록
        self.model = MailModel()
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for col, width in ((0, 120), (1, 150), (3, 140), (4, 36)):
            self.table.setColumnWidth(col, width)
        self.table.selectionModel().selectionChanged.connect(self._mail_selected)
        self.table.doubleClicked.connect(lambda _: self.open_msg())
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_menu)

        # 오른쪽: 미리보기
        self.preview = Preview(self.open_msg, self.edit_labels)

        split = QSplitter()
        split.addWidget(self.tree)
        split.addWidget(self.table)
        split.addWidget(self.preview)
        split.setSizes([240, 620, 540])

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addLayout(top)
        layout.addWidget(split, 1)
        self.setCentralWidget(central)

        self.refresh_all()

    # --- 갱신 ---

    def refresh_all(self) -> None:
        self.refresh_labels()
        self.refresh_list()
        self.refresh_info()

    def refresh_info(self) -> None:
        last = self.lib.last_ok_sync()
        text = f"메일 {self.lib.count():,}건 · 마지막 동기화 {last[:16] if last else '없음'}"
        stale = last is None or (datetime.now() - datetime.fromisoformat(last)).days >= SYNC_WARN_DAYS
        self.info.setText(text)
        self.info.setStyleSheet("color: #c00; font-weight: bold" if stale else "")
        self.info.setToolTip("보존 기간(약 1개월)이 지나기 전에 동기화하세요." if stale else "")

    def refresh_labels(self) -> None:
        current = self.label_filter
        self.tree.blockSignals(True)
        self.tree.clear()
        all_item = QTreeWidgetItem([f"전체 메일 ({self.lib.count():,})"])
        all_item.setData(0, ROLE_LABEL, ("all", None))
        none_item = QTreeWidgetItem([f"라벨 없음 ({self.lib.unlabeled_count():,})"])
        none_item.setData(0, ROLE_LABEL, ("none", UNLABELED))
        self.tree.addTopLevelItem(all_item)
        self.tree.addTopLevelItem(none_item)
        nodes: dict[str, QTreeWidgetItem] = {}
        counts = self.lib.label_counts()  # 그 라벨에 직접 붙은 수 (눌렀을 때 목록에 보이는 수)
        totals = self.lib.tree_counts()  # 하위 라벨 포함
        for name in sorted(counts):
            parts = name.split("/")
            for depth in range(1, len(parts) + 1):
                path = "/".join(parts[:depth])
                if path in nodes:
                    continue
                total, own = totals.get(path, 0), counts.get(path, 0)
                item = QTreeWidgetItem([f"{parts[depth - 1]} ({own:,})"])
                if total != own:
                    item.setToolTip(0, f"이 라벨 {own:,}건 · 하위 라벨 포함 {total:,}건")
                item.setData(0, ROLE_LABEL, ("label", path))
                parent = nodes.get("/".join(parts[:depth - 1]))
                (parent.addChild(item) if parent else self.tree.addTopLevelItem(item))
                nodes[path] = item
        self.tree.expandAll()
        target = all_item
        if current == UNLABELED:
            target = none_item
        elif current in nodes:
            target = nodes[current]
        else:
            self.label_filter = ALL
        target.setSelected(True)
        self.tree.setCurrentItem(target)
        self.tree.blockSignals(False)

    def refresh_list(self) -> None:
        selected = set(self.selected_ids())
        rows = self.lib.search(self.search.text(), self.label_filter)
        self.model.set_rows(rows)
        if selected:
            sm = self.table.selectionModel()
            for i, row in enumerate(rows):
                if row.id in selected:
                    sm.select(self.model.index(i, 0), sm.SelectionFlag.Select | sm.SelectionFlag.Rows)
        if not self.selected_ids():
            self.preview.clear()

    # --- 선택 ---

    def selected_ids(self) -> list[int]:
        return [self.model.rows[i.row()].id for i in self.table.selectionModel().selectedRows()]

    def _label_selected(self) -> None:
        items = self.tree.selectedItems()
        if not items:
            return
        _, value = items[0].data(0, ROLE_LABEL)
        self.label_filter = value
        self.refresh_list()

    def _mail_selected(self) -> None:
        ids = self.selected_ids()
        if len(ids) != 1:
            self.preview.clear()
            if ids:
                self.preview.label_btn.setEnabled(True)
                self.preview.header.setText(f"{len(ids)}건 선택됨")
            return
        mail = self.lib.get(ids[0])
        if mail:
            self.preview.show_mail(mail)

    # --- 동작 ---

    def open_msg(self) -> None:
        ids = self.selected_ids()
        if len(ids) != 1:
            return
        mail = self.lib.get(ids[0])
        if mail is None or not mail.msg_path.exists():
            QMessageBox.warning(self, "mailer", ".msg 파일을 찾을 수 없습니다.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(mail.msg_path)))

    def edit_labels(self) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        menu = self._label_menu(ids)
        menu.exec(self.preview.label_btn.mapToGlobal(self.preview.label_btn.rect().bottomLeft()))

    def _label_menu(self, ids: list[int]) -> QMenu:
        menu = QMenu(self)
        menu.addAction("라벨 붙이기…", lambda: self._add_label(ids))
        current = sorted({label for r in self.model.rows if r.id in ids for label in r.labels})
        if current:
            remove = menu.addMenu("라벨 떼기")
            for name in current:
                remove.addAction(name, lambda n=name: self._remove_label(ids, n))
        return menu

    def _add_label(self, ids: list[int]) -> None:
        names = sorted(self.lib.label_counts())
        name, ok = QInputDialog.getItem(
            self, "라벨 붙이기", f"{len(ids)}건에 붙일 라벨 ('/' 로 하위 라벨):", names, 0, True
        )
        if ok and name.strip():
            self.lib.add_label(ids, name)
            self.refresh_labels()
            self.refresh_list()
            self._mail_selected()

    def _remove_label(self, ids: list[int], name: str) -> None:
        self.lib.remove_label(ids, name)
        self.refresh_labels()
        self.refresh_list()
        self._mail_selected()

    def _table_menu(self, pos) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        menu = self._label_menu(ids)
        if len(ids) == 1:
            menu.addSeparator()
            menu.addAction(".msg 열기 (Outlook)", self.open_msg)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _tree_menu(self, pos) -> None:
        item = self.tree.itemAt(pos)
        if item is None:
            return
        kind, name = item.data(0, ROLE_LABEL)
        if kind != "label":
            return
        menu = QMenu(self)
        menu.addAction("이름 바꾸기…", lambda: self._rename_label(name))
        menu.addAction("라벨 삭제 (메일은 남음)", lambda: self._delete_label(name))
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def _rename_label(self, name: str) -> None:
        new, ok = QInputDialog.getText(self, "라벨 이름 바꾸기", "새 이름 ('/' 로 하위 라벨):", text=name)
        if ok and new.strip() and new.strip() != name:
            self.label_filter = self.lib.rename_label(name, new)
            self.refresh_labels()
            self.refresh_list()

    def _delete_label(self, name: str) -> None:
        answer = QMessageBox.question(
            self, "라벨 삭제", f"'{name}' 라벨(하위 포함)을 삭제할까요?\n메일은 지워지지 않습니다."
        )
        if answer == QMessageBox.StandardButton.Yes:
            self.lib.delete_label(name)
            self.refresh_labels()
            self.refresh_list()

    def sync(self) -> None:
        dialog = SyncDialog(self, self.make_source, self.root, self.log_dir)
        dialog.exec()
        self.refresh_all()

    def closeEvent(self, event) -> None:  # noqa: N802
        self.lib.close()
        super().closeEvent(event)


def run_app(root: Path, make_source: Callable[[], MailSource], log_dir: Path | None) -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("mailer")
    if sys.platform == "win32":
        app.setFont(QFont("Malgun Gothic", 9))

    def excepthook(kind, value, tb) -> None:
        text = "".join(traceback.format_exception(kind, value, tb))
        if log_dir is not None:
            log_dir.mkdir(parents=True, exist_ok=True)
            with (log_dir / "app-error.txt").open("a", encoding="utf-8") as f:
                f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S}\n{text}\n")
        QMessageBox.critical(None, "mailer 오류", text[-2000:])

    sys.excepthook = excepthook
    try:
        window = MainWindow(root, make_source, log_dir)
    except Exception as exc:
        QMessageBox.critical(None, "mailer", f"보존 경로를 열 수 없습니다: {root}\n\n{exc}")
        return 2
    window.show()
    return app.exec()
