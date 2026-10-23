"""PySide6 데스크톱 뷰어: 검색, 라벨, 미리보기, 동기화.

화면 구성: [툴바] / [라벨] | [메일 목록] | [미리보기] / [상태 표시줄]
색·아이콘·스타일시트는 theme.py. 툴바에서 밝은/어두운 테마를 바꾸고, 고른 값은 설정 파일에 기억한다.
동기화는 작업 스레드에서 collect 엔진을 돌린다. COM 은 그 스레드 안에서 초기화한다
(pythoncom.CoInitialize). UI 스레드에서는 Outlook 에 접근하지 않는다.
"""

from __future__ import annotations

import hashlib
import html
import sys
import threading
import traceback
from collections.abc import Callable
from datetime import date, datetime, timedelta
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QPoint, QRect, QSize, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QColor, QIcon, QDesktopServices, QFont, QFontDatabase, QImage, QKeySequence, QPainter, QPalette, QShortcut, QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLayout,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableView,
    QTextBrowser,
    QToolBar,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import theme
from .collect import run_collect
from .htmlcolor import darken_html
from .mailhtml import normalize as normalize_html
from .mailhtml import image_sizes
from .mailhtml import shape as html_shape
from .msgfile import inline_images
from .library import ALL, UNLABELED, Library, MailRow
from .source import MailSource

SYNC_WARN_DAYS = 21  # 보존 기간(약 1개월)보다 넉넉히 앞에서 경고
ROLE_LABEL = Qt.ItemDataRole.UserRole
ROLE_COUNT = Qt.ItemDataRole.UserRole + 1
WEEKDAYS = "월화수목금토일"
RECIPIENTS_SHOWN = 3  # 받는 사람·참조는 처음에 이만큼만 보이고 나머지는 '외 N명'


def friendly_date(received: str, today: date | None = None) -> str:
    """목록용 짧은 날짜: 오늘 14:51 / 어제 14:51 / 10-21 (화) / 2025-11-03."""
    try:
        when = datetime.strptime(received[:16], "%Y-%m-%d %H:%M")
    except ValueError:
        return received
    today = today or date.today()
    if when.date() == today:
        return f"오늘 {when:%H:%M}"
    if when.date() == today - timedelta(days=1):
        return f"어제 {when:%H:%M}"
    if when.year == today.year:
        return f"{when:%m-%d} ({WEEKDAYS[when.weekday()]})"
    return f"{when:%Y-%m-%d}"


def split_people(text: str) -> list[str]:
    return [p.strip() for p in text.replace(",", ";").split(";") if p.strip()]


# ---------------------------------------------------------------- 메일 목록


class MailModel(QAbstractTableModel):
    HEADERS = ("보낸 사람", "제목", "라벨", "첨부", "일시")  # 보낸 편지함도 있어서 '받은 시각' 이 아니라 '일시'
    SENDER, SUBJECT, LABEL, CLIP, DATE = range(5)

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[MailRow] = []
        self.clip_icon = None
        self.muted = QColor("gray")

    def set_rows(self, rows: list[MailRow]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):  # noqa: N802
        if orientation != Qt.Orientation.Horizontal:
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            return self.HEADERS[section]
        if role == Qt.ItemDataRole.TextAlignmentRole:
            align = {self.DATE: Qt.AlignmentFlag.AlignRight, self.CLIP: Qt.AlignmentFlag.AlignHCenter}
            return int(align.get(section, Qt.AlignmentFlag.AlignLeft) | Qt.AlignmentFlag.AlignVCenter)
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            if col == self.SENDER:
                return row.sender
            if col == self.SUBJECT:
                return row.subject or "(제목 없음)"
            if col == self.LABEL:
                return ", ".join(_short(label) for label in row.labels)
            if col == self.DATE:
                return friendly_date(row.received)
        if role == Qt.ItemDataRole.DecorationRole and col == self.CLIP and row.has_attachments:
            return self.clip_icon
        if role == Qt.ItemDataRole.ToolTipRole:
            if col == self.LABEL:
                return "\n".join(row.labels)
            if col == self.DATE:
                return row.received
            if col == self.CLIP and row.has_attachments:
                return "첨부 있음"
        if role == Qt.ItemDataRole.TextAlignmentRole and col == self.DATE:
            return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        if role == Qt.ItemDataRole.ForegroundRole and col in (self.LABEL, self.DATE):
            return self.muted
        return None


def _menu_text(text: str) -> str:
    return text.replace("&", "&&")  # 메뉴에서 & 는 단축키 표시라서


def _short(label: str) -> str:
    return label.rsplit("/", 1)[-1]


class MailTable(QTableView):
    """마우스가 올라간 줄 전체를 옅게 칠한다 (스타일시트의 :hover 는 칸 하나만 칠해서)."""

    def __init__(self) -> None:
        super().__init__()
        self.hover_row = -1
        self.setMouseTracking(True)

    def _set_hover(self, row: int) -> None:
        if row != self.hover_row:
            self.hover_row = row
            self.viewport().update()

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        self._set_hover(self.rowAt(event.position().toPoint().y()))
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._set_hover(-1)
        super().leaveEvent(event)


class RowDelegate(QStyledItemDelegate):
    """목록: 마우스가 올라간 줄 배경, 첨부 아이콘은 칸 가운데."""

    def __init__(self, table: MailTable) -> None:
        super().__init__(table)
        self.table = table
        self.hover = QColor("#eef2f8")
        self.grid = QColor("#e9ecf0")

    def paint(self, painter: QPainter, option, index) -> None:
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        if index.row() == self.table.hover_row and not selected:
            painter.fillRect(option.rect, self.hover)
        self._paint_cell(painter, option, index, selected)
        if index.column() != MailModel.DATE and not selected:  # 열 구분선 (마지막 열 제외)
            painter.save()
            painter.setPen(self.grid)
            painter.drawLine(option.rect.topRight(), option.rect.bottomRight())
            painter.restore()

    def _paint_cell(self, painter: QPainter, option, index, selected: bool) -> None:
        if index.column() != MailModel.CLIP:
            super().paint(painter, option, index)
            return
        icon = index.data(Qt.ItemDataRole.DecorationRole)
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.icon = QIcon()
        opt.features &= ~QStyleOptionViewItem.ViewItemFeature.HasDecoration
        widget = opt.widget
        (widget.style() if widget else QApplication.style()).drawControl(
            QStyle.ControlElement.CE_ItemViewItem, opt, painter, widget
        )
        if isinstance(icon, QIcon) and not icon.isNull():
            size = self.table.iconSize()
            rect = QRect(0, 0, size.width(), size.height())
            rect.moveCenter(option.rect.center())
            mode = QIcon.Mode.Selected if selected and option.state & QStyle.StateFlag.State_Active else QIcon.Mode.Normal
            icon.paint(painter, rect, Qt.AlignmentFlag.AlignCenter, mode)


class CountDelegate(QStyledItemDelegate):
    """라벨 트리: 이름 오른쪽에 건수 배지. 한 칸에 그려서 선택 표시가 끊기지 않게 한다."""

    def __init__(self, parent) -> None:
        super().__init__(parent)
        self.colors = theme.colors(theme.LIGHT)

    def sizeHint(self, option, index) -> QSize:  # noqa: N802
        size = super().sizeHint(option, index)
        return QSize(size.width() + 40, max(size.height(), 28))

    def paint(self, painter: QPainter, option, index) -> None:
        count = index.data(ROLE_COUNT)
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        badge = QRect()
        if count is not None:
            text = f"{count:,}"
            fm = opt.fontMetrics
            width = fm.horizontalAdvance(text) + 14
            badge = QRect(opt.rect.right() - width - 6, opt.rect.center().y() - 9, width, 18)
            room = badge.left() - opt.rect.left() - opt.decorationSize.width() - 16
            opt.text = fm.elidedText(opt.text, Qt.TextElideMode.ElideRight, max(room, 20))
        widget = opt.widget
        style = widget.style() if widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, widget)
        if count is not None:
            painter.save()
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(self.colors.raised))
            painter.drawRoundedRect(badge, 9, 9)
            painter.setPen(QColor(self.colors.muted))
            painter.drawText(badge, Qt.AlignmentFlag.AlignCenter, f"{count:,}")
            painter.restore()


class FlowLayout(QLayout):
    """칩(첨부·라벨)을 줄 바꿈하며 늘어놓는다."""

    def __init__(self, parent=None, spacing: int = 6) -> None:
        super().__init__(parent)
        self._items: list = []
        self._spacing = spacing
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item) -> None:  # noqa: N802
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, i):  # noqa: N802
        return self._items[i] if 0 <= i < len(self._items) else None

    def takeAt(self, i):  # noqa: N802
        return self._items.pop(i) if 0 <= i < len(self._items) else None

    def clear(self) -> None:
        while self._items:
            item = self._items.pop()
            if item.widget():
                item.widget().hide()  # deleteLater 는 늦게 지워져서, 그 사이 새 칩과 겹쳐 보인다
                item.widget().deleteLater()

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._place(QRect(0, 0, width, 0), move=False)

    def setGeometry(self, rect) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._place(rect, move=True)

    def sizeHint(self) -> QSize:  # noqa: N802
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # noqa: N802
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        return size

    def _place(self, rect: QRect, move: bool) -> int:
        x, y, line = rect.x(), rect.y(), 0
        for item in self._items:
            hint = item.sizeHint()
            if x + hint.width() > rect.right() and line > 0:
                x, y, line = rect.x(), y + line + self._spacing, 0
            if move:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x += hint.width() + self._spacing
            line = max(line, hint.height())
        return y + line - rect.y()


def card() -> QFrame:
    frame = QFrame(objectName="card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(1, 1, 1, 1)
    layout.setSpacing(0)
    return frame


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


class MailBody(QTextBrowser):
    """본문 표시. 'cid:' 이미지는 .msg 의 첨부에서 꺼내 보여준다. 원격(http) 이미지는 불러오지 않는다.

    Qt 는 큰 이미지를 <img width=..> 크기로 줄여 그릴 때 픽셀을 건너뛰어(nearest) 캡처 이미지의 글자가 깨진다.
    그래서 HTML 에 적힌 표시 크기로 미리 부드럽게 줄여서 넘긴다 (화면 배율만큼 더 크게)."""

    def __init__(self) -> None:
        super().__init__()
        self.images: dict[str, bytes] = {}
        self.sizes: dict[str, tuple[float | None, float | None]] = {}

    def loadResource(self, kind: int, url: QUrl):  # noqa: N802
        if url.scheme().lower() == "cid":
            key = url.path()
            data = self.images.get(key) or self.images.get(key.split("@", 1)[0])
            if data:
                image = QImage.fromData(data)
                if not image.isNull():
                    return self._fit(image, self.sizes.get(key))
            return None
        if url.scheme().lower() in ("http", "https"):
            return None
        return super().loadResource(kind, url)

    def _fit(self, image: QImage, size: tuple[float | None, float | None] | None) -> QImage:
        if not size:
            return image
        width, height = size
        if width is None:
            width = height * image.width() / image.height()
        if height is None:
            height = width * image.height() / image.width()
        ratio = self.devicePixelRatioF()
        target = QSize(max(1, round(width * ratio)), max(1, round(height * ratio)))
        if image.width() <= target.width() and image.height() <= target.height():
            return image  # 줄일 필요 없음
        scaled = image.scaled(target, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        scaled.setDevicePixelRatio(ratio)
        return scaled


_AVATAR_COLORS = ("#5b7be0", "#2f9e8f", "#d0703c", "#9b59b6", "#c2507a", "#3c8dbc", "#6d8b3a", "#b07d2b")


class Avatar(QLabel):
    """보낸 사람 이름 첫 글자 동그라미."""

    def __init__(self) -> None:
        super().__init__(objectName="avatar")
        self.setFixedSize(36, 36)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def set_name(self, name: str) -> None:
        name = name.strip().strip("'\"")
        self.setText(name[:1].upper() if name else "?")
        color = _AVATAR_COLORS[int(hashlib.md5(name.encode("utf-8")).hexdigest(), 16) % len(_AVATAR_COLORS)]
        self.setStyleSheet(f"background: {color};")


class Preview(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.colors = theme.colors(theme.LIGHT)
        self.dark = False
        self._mail = None
        self._expanded = False

        self.title = QLabel(objectName="title", wordWrap=True)
        self.title.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.avatar = Avatar()
        self.sender = QLabel(objectName="sender")
        self.sender.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.when = QLabel(objectName="muted")
        self.recipients = QLabel(objectName="muted", wordWrap=True)
        self.recipients.setTextFormat(Qt.TextFormat.RichText)
        self.recipients.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.LinksAccessibleByMouse
        )
        self.recipients.linkActivated.connect(self._toggle_recipients)
        # 어두운 테마면 본문 색을 보정해서 보여주고, 이 버튼을 누르면 흰 배경 원본으로 (Outlook 의 배경 전환과 같음)
        self.paper_btn = QToolButton(objectName="flat", text="원본 색", checkable=True)
        self.paper_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.paper_btn.setToolTip("본문을 흰 배경(원래 색)으로 보기")
        self.paper_btn.toggled.connect(self._render_body)

        who = QVBoxLayout()
        who.setSpacing(2)
        line = QHBoxLayout()
        line.addWidget(self.sender, 1)
        line.addWidget(self.when)
        who.addLayout(line)
        who.addWidget(self.recipients)
        head = QHBoxLayout()
        head.setSpacing(10)
        head.addWidget(self.avatar, 0, Qt.AlignmentFlag.AlignTop)
        head.addLayout(who, 1)

        self.chips = QWidget()
        self.chip_layout = FlowLayout(self.chips)
        tools = QHBoxLayout()
        tools.addWidget(self.chips, 1)
        tools.addWidget(self.paper_btn, 0, Qt.AlignmentFlag.AlignTop)

        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 14, 18, 10)
        header_layout.setSpacing(10)
        header_layout.addWidget(self.title)
        header_layout.addLayout(head)
        header_layout.addLayout(tools)
        self.header = header

        self.body = MailBody()
        self.body.setOpenLinks(False)  # 링크를 눌러도 이 창에서 이동하지 않는다
        self.body.anchorClicked.connect(self._link)
        self.body.document().setDocumentMargin(16)

        self.empty = QLabel("메일을 고르면 여기에 보입니다", objectName="muted", alignment=Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(header)
        layout.addWidget(QFrame(objectName="divider"))
        layout.addWidget(self.body, 1)
        layout.addWidget(self.empty, 1)
        self.clear()

    def set_theme(self, colors: theme.Colors, dark: bool) -> None:
        self.colors = colors
        self.dark = dark
        self.paper_btn.setIcon(theme.icon("sun", colors.text, 16))
        self.paper_btn.setVisible(dark and self._mail is not None and bool(self._mail.html_body))
        if self._mail is not None:
            self._fill_chips(self._mail)
            self._render_body()

    def clear(self, message: str = "") -> None:
        self._mail = None
        self.header.setVisible(bool(message))
        self.title.setText(message)
        for widget in (self.avatar, self.sender, self.when, self.recipients, self.chips, self.paper_btn):
            widget.setVisible(False)
        self.body.clear()
        self.body.setVisible(False)
        self.empty.setVisible(not message)

    def show_mail(self, mail) -> None:
        self._mail = mail
        self._expanded = False
        self.header.setVisible(True)
        self.empty.setVisible(False)
        self.body.setVisible(True)
        for widget in (self.avatar, self.sender, self.when, self.recipients, self.chips):
            widget.setVisible(True)
        self.title.setText(mail.subject or "(제목 없음)")
        self.avatar.set_name(mail.sender_name or mail.sender_email)
        # 회사 내부(Exchange) 주소 '/O=...' 는 보여주지 않는다 (repair 전 메일)
        email = mail.sender_email if "@" in mail.sender_email else ""
        muted = self.colors.muted
        self.sender.setText(
            f"{html.escape(mail.sender_name or email or '(알 수 없음)')}"
            + (f" <span style='color:{muted};font-weight:400'>&lt;{html.escape(email)}&gt;</span>" if email and mail.sender_name else "")
        )
        self.when.setText(mail.received[:16])
        self.when.setToolTip(mail.received)
        self._fill_recipients()
        self._fill_chips(mail)
        self.paper_btn.setVisible(self.dark and bool(mail.html_body))
        self._render_body()

    def _fill_recipients(self) -> None:
        mail = self._mail
        lines = []
        for name, value in (("받는 사람", mail.to), ("참조", mail.cc)):
            people = split_people(value)
            if not people:
                continue
            shown = people if self._expanded else people[:RECIPIENTS_SHOWN]
            text = ", ".join(html.escape(p) for p in shown)
            if len(people) > len(shown):
                text += f" <a href='more' style='color:{self.colors.accent};text-decoration:none'>외 {len(people) - len(shown)}명</a>"
            lines.append(f"{name}&nbsp;&nbsp;{text}")
        if self._expanded:
            lines.append(f"<a href='less' style='color:{self.colors.accent};text-decoration:none'>접기</a>")
        self.recipients.setText("<br>".join(lines))
        self.recipients.setVisible(bool(lines))

    def _toggle_recipients(self, _link: str) -> None:
        self._expanded = not self._expanded
        self._fill_recipients()

    def _fill_chips(self, mail) -> None:
        self.chip_layout.clear()
        for name in mail.attachments:
            self.chip_layout.addWidget(self._chip(name, name, "paperclip", "chip"))
        for label in mail.labels:
            self.chip_layout.addWidget(self._chip(_short(label), label, "tag", "labelChip"))
        self.chips.setVisible(bool(mail.attachments or mail.labels))
        self.chips.updateGeometry()

    def _chip(self, text: str, tip: str, icon_name: str, kind: str) -> QFrame:
        chip = QFrame(objectName=kind)
        chip.setToolTip(tip)
        row = QHBoxLayout(chip)
        row.setContentsMargins(8, 3, 10, 3)
        row.setSpacing(5)
        icon = QLabel()
        icon.setPixmap(theme.icon(icon_name, self.colors.muted, 13).pixmap(13, 13))
        row.addWidget(icon)
        row.addWidget(QLabel(text))
        return chip

    def _render_body(self) -> None:
        mail = self._mail
        if mail is None:
            return
        self.body.images = {}
        self.body.sizes = image_sizes(mail.html_body)
        if "cid:" in mail.html_body and mail.msg_path.exists():
            try:
                self.body.images = inline_images(mail.msg_path)
            except Exception:
                pass  # 이미지가 안 보일 뿐 본문은 보여준다
        paper = not self.dark or self.paper_btn.isChecked()
        palette = self.body.palette()
        base, text = (self.colors.paper, "#000000") if paper and self.dark else (self.colors.surface, self.colors.text)
        for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled):
            palette.setColor(group, QPalette.ColorRole.Base, QColor(base))
            palette.setColor(group, QPalette.ColorRole.Text, QColor(text))
        self.body.setPalette(palette)
        self.body.setStyleSheet(f"background: {base}; color: {text};")
        if mail.html_body:
            body = normalize_html(mail.html_body)  # Word·Excel 표 테두리를 Qt 가 읽는 형태로
            self.body.setHtml(body if paper else darken_html(body))
        else:
            self.body.setPlainText(mail.body)

    def _link(self, url: QUrl) -> None:
        if url.scheme() in ("http", "https", "mailto"):
            QDesktopServices.openUrl(url)


# ---------------------------------------------------------------- 메인 창

_LABEL_ICONS = {"받은 편지함": "inbox", "보낸 편지함": "send", "정크 메일": "shield-alert", "보관": "archive"}


class MainWindow(QMainWindow):
    def __init__(
        self,
        root: Path,
        make_source: Callable[[], MailSource],
        log_dir: Path | None,
        settings_path: Path | None = None,
    ) -> None:
        super().__init__()
        self.root = root
        self.make_source = make_source
        self.log_dir = log_dir
        self.settings_path = settings_path
        self.lib = Library(root)
        self.label_filter: str | None = ALL
        self.mode = theme.load_mode(settings_path)
        self.colors = theme.colors(self.mode)
        self.setWindowTitle(f"inlook — {root}")
        self.resize(1400, 860)

        # 위: 툴바 (동기화 · 검색 · 메일 동작 · 테마)
        bar = QToolBar(objectName="main", movable=False, floatable=False)
        bar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        bar.setIconSize(QSize(17, 17))
        bar.setContextMenuPolicy(Qt.ContextMenuPolicy.PreventContextMenu)
        self.addToolBar(bar)
        self.sync_act = bar.addAction("동기화", self.sync)
        self.sync_act.setToolTip("Outlook 의 새 메일을 가져와 저장합니다")
        bar.widgetForAction(self.sync_act).setObjectName("primary")
        gap = QWidget()
        gap.setFixedWidth(8)
        bar.addWidget(gap)
        self.search = QLineEdit(objectName="search", placeholderText="메일 검색 — 제목·사람·본문·첨부 파일명, 띄어쓰기로 여러 단어")
        self.search.setClearButtonEnabled(True)
        self.search.setMaximumWidth(720)
        self.search.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._search_icon = QAction(self)
        self.search.addAction(self._search_icon, QLineEdit.ActionPosition.LeadingPosition)
        self._debounce = QTimer(self, singleShot=True, interval=250)
        self._debounce.timeout.connect(self.refresh_list)
        self.search.textChanged.connect(self._debounce.start)
        bar.addWidget(self.search)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        bar.addWidget(spacer)
        self.label_act = bar.addAction("라벨 변경", self.edit_labels)
        self.open_act = bar.addAction(".msg 열기", self.open_msg)
        self.open_act.setToolTip("Outlook 으로 원본 열기 (목록에서 더블클릭)")
        self.delete_act = bar.addAction("삭제", self.delete_selected)
        self.delete_act.setToolTip("선택한 메일 삭제 (Del)")
        bar.addSeparator()
        self.theme_act = bar.addAction("", self.toggle_theme)

        # 왼쪽: 라벨
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setIndentation(14)
        self.tree.setIconSize(QSize(16, 16))
        self.count_delegate = CountDelegate(self.tree)
        self.tree.setItemDelegate(self.count_delegate)
        self.tree.itemSelectionChanged.connect(self._label_selected)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_menu)
        left = card()
        left.layout().addWidget(self.tree)

        # 가운데: 목록
        self.model = MailModel()
        self.table = MailTable()
        self.table.setModel(self.model)
        self.row_delegate = RowDelegate(self.table)
        self.table.setItemDelegate(self.row_delegate)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setIconSize(QSize(14, 14))
        header = self.table.horizontalHeader()
        header.setHighlightSections(False)
        header.setSectionResizeMode(MailModel.SUBJECT, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(MailModel.CLIP, QHeaderView.ResizeMode.Fixed)
        for col, width in ((MailModel.SENDER, 170), (MailModel.LABEL, 130), (MailModel.CLIP, 44), (MailModel.DATE, 120)):
            self.table.setColumnWidth(col, width)
        self.table.selectionModel().selectionChanged.connect(self._mail_selected)
        self.table.doubleClicked.connect(lambda _: self.open_msg())
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_menu)
        delete = QShortcut(QKeySequence(QKeySequence.StandardKey.Delete), self.table)
        delete.setContext(Qt.ShortcutContext.WidgetShortcut)
        delete.activated.connect(self.delete_selected)
        middle = card()
        middle.layout().addWidget(self.table)

        # 오른쪽: 미리보기
        self.preview = Preview()
        right = card()
        right.layout().addWidget(self.preview)

        split = QSplitter()
        split.setHandleWidth(8)
        split.setChildrenCollapsible(False)
        for pane in (left, middle, right):
            split.addWidget(pane)
        split.setSizes([250, 610, 540])

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 4, 8, 0)
        layout.addWidget(split, 1)
        self.setCentralWidget(central)

        # 아래: 상태 표시줄
        self.count_label = QLabel()
        self.info = QLabel()
        self.statusBar().setSizeGripEnabled(False)
        self.statusBar().addWidget(self.count_label, 1)
        self.statusBar().addPermanentWidget(self.info)

        self.apply_theme()
        self.refresh_all()

    # --- 테마 ---

    def apply_theme(self) -> None:
        app = QApplication.instance()
        theme.apply(app, self.mode)
        c = self.colors = theme.colors(self.mode)
        dark = self.mode == theme.DARK
        for act, name in ((self.label_act, "tag"), (self.open_act, "external-link"), (self.delete_act, "trash-2")):
            act.setIcon(theme.icon(name, c.text))
        self.sync_act.setIcon(theme.icon("refresh-cw", c.accent_text))
        self.theme_act.setIcon(theme.icon("sun" if dark else "moon", c.text))
        self.theme_act.setText("밝게" if dark else "어둡게")
        self.theme_act.setToolTip("밝은 테마로" if dark else "어두운 테마로")
        self._search_icon.setIcon(theme.icon("search", c.muted, 16))
        self.model.clip_icon = theme.icon("paperclip", c.muted, 14, selected=c.accent_text)
        self.model.muted = QColor(c.muted)
        self.count_delegate.colors = c
        self.row_delegate.hover = QColor(c.hover)
        self.row_delegate.grid = QColor(c.grid)
        # 라벨 트리의 선택 표시는 강조색 대신 옅은 색 (들여쓰기 부분까지 같은 색으로 칠해지게 팔레트로 지정)
        palette = self.tree.palette()
        palette.setColor(QPalette.ColorRole.Highlight, QColor(c.select_inactive))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(c.text))
        self.tree.setPalette(palette)
        self.preview.set_theme(c, dark)
        if self.tree.topLevelItemCount():
            self.refresh_labels()
        self.table.viewport().update()

    def toggle_theme(self) -> None:
        self.mode = theme.LIGHT if self.mode == theme.DARK else theme.DARK
        theme.save_mode(self.settings_path, self.mode)
        self.apply_theme()

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
        self.info.setObjectName("stale" if stale else "")
        self.info.style().unpolish(self.info)
        self.info.style().polish(self.info)
        self.info.setToolTip("보존 기간(약 1개월)이 지나기 전에 동기화하세요." if stale else "")

    def _label_icon(self, path: str, has_children: bool):
        c = self.colors
        name = _LABEL_ICONS.get(path) or ("folder" if has_children else "tag")
        return theme.icon(name, c.muted, 16)

    def refresh_labels(self) -> None:
        current = self.label_filter
        c = self.colors
        self.tree.blockSignals(True)
        self.tree.clear()
        all_item = QTreeWidgetItem(["전체 메일"])
        all_item.setData(0, ROLE_LABEL, ("all", None))
        all_item.setData(0, ROLE_COUNT, self.lib.count())
        all_item.setIcon(0, theme.icon("mails", c.muted, 16))
        none_item = QTreeWidgetItem(["라벨 없음"])
        none_item.setData(0, ROLE_LABEL, ("none", UNLABELED))
        none_item.setData(0, ROLE_COUNT, self.lib.unlabeled_count())
        none_item.setIcon(0, theme.icon("circle-dashed", c.muted, 16))
        title = QTreeWidgetItem(["라벨"])
        title.setFlags(Qt.ItemFlag.NoItemFlags)
        font = QFont(self.tree.font())
        font.setBold(True)
        title.setFont(0, font)
        title.setForeground(0, QColor(c.muted))
        for item in (all_item, none_item, title):
            self.tree.addTopLevelItem(item)
        nodes: dict[str, QTreeWidgetItem] = {}
        counts = self.lib.label_counts()  # 그 라벨에 직접 붙은 수 (눌렀을 때 목록에 보이는 수)
        totals = self.lib.tree_counts()  # 하위 라벨 포함
        parents = {n.rsplit("/", 1)[0] for n in counts if "/" in n}
        parents |= {"/".join(n.split("/")[:d]) for n in counts for d in range(1, n.count("/") + 1)}
        for name in sorted(counts):
            parts = name.split("/")
            for depth in range(1, len(parts) + 1):
                path = "/".join(parts[:depth])
                if path in nodes:
                    continue
                total, own = totals.get(path, 0), counts.get(path, 0)
                item = QTreeWidgetItem([parts[depth - 1]])
                item.setData(0, ROLE_COUNT, own)
                item.setIcon(0, self._label_icon(path, path in parents))
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
        # 라벨 하나를 골랐으면 라벨 열은 모두 같은 값이라 숨긴다
        self.table.setColumnHidden(MailModel.LABEL, self.label_filter not in (ALL,))
        if selected:
            sm = self.table.selectionModel()
            for i, row in enumerate(rows):
                if row.id in selected:
                    sm.select(self.model.index(i, 0), sm.SelectionFlag.Select | sm.SelectionFlag.Rows)
        if not self.selected_ids():
            self.preview.clear()
        self.count_label.setText(f"{len(rows):,}건" + (" (검색 결과)" if self.search.text().strip() else ""))
        self._update_actions()

    def _update_actions(self) -> None:
        n = len(self.selected_ids())
        self.label_act.setEnabled(n > 0)
        self.delete_act.setEnabled(n > 0)
        self.open_act.setEnabled(n == 1)

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
        self._update_actions()
        ids = self.selected_ids()
        if len(ids) != 1:
            self.preview.clear(f"{len(ids)}건 선택됨" if ids else "")
            return
        mail = self.lib.get(ids[0])
        if mail:
            self.preview.show_mail(self.lib.with_html(mail))

    # --- 동작 ---

    def open_msg(self) -> None:
        ids = self.selected_ids()
        if len(ids) != 1:
            return
        mail = self.lib.get(ids[0])
        if mail is None or not mail.msg_path.exists():
            QMessageBox.warning(self, "inlook", ".msg 파일을 찾을 수 없습니다.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(mail.msg_path)))

    def edit_labels(self) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        button = self.findChild(QToolBar).widgetForAction(self.label_act)
        self._label_submenu(QMenu(self), ids).exec(button.mapToGlobal(button.rect().bottomLeft()))

    # 화면에서는 메일당 라벨 하나만 다룬다 (사용자 결정 2026-10-15). DB 는 여러 개를 담을 수 있는 구조 그대로.

    def _label_submenu(self, menu: QMenu, ids: list[int]) -> QMenu:
        """라벨 트리를 그대로 서브메뉴로. 고르면 바로 바뀐다 (기존 라벨은 떼어짐)."""
        current = {label for r in self.model.rows if r.id in ids for label in r.labels}
        names = sorted(self.lib.label_counts())
        has_children = {n.rsplit("/", 1)[0] for n in names if "/" in n}
        menus: dict[str, QMenu] = {"": menu}

        def parent_menu(path: str) -> QMenu:
            if path in menus:
                return menus[path]
            parent, _, leaf = path.rpartition("/")
            sub = parent_menu(parent).addMenu(_menu_text(leaf))
            menus[path] = sub
            return sub

        for name in names:
            parent, _, leaf = name.rpartition("/")
            target = parent_menu(name) if name in has_children else parent_menu(parent)
            action = target.addAction(_menu_text(leaf))
            action.setCheckable(True)
            action.setChecked(name in current)
            action.triggered.connect(lambda _=False, n=name: self._apply_label(ids, n))
            if name in has_children:
                target.addSeparator()
        menu.addSeparator()
        menu.addAction("새 라벨…", lambda: self._new_label(ids))
        return menu

    def _new_label(self, ids: list[int]) -> None:
        name, ok = QInputDialog.getText(
            self, "새 라벨", f"{len(ids)}건에 붙일 새 라벨 (기존 라벨은 떼어집니다, '/' 로 하위 라벨):"
        )
        if ok and name.strip():
            self._apply_label(ids, name)

    def _apply_label(self, ids: list[int], name: str) -> None:
        self.lib.set_label(ids, name)
        self.refresh_labels()
        self.refresh_list()
        self._mail_selected()

    def _table_menu(self, pos) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        menu = QMenu(self)
        self._label_submenu(menu.addMenu("라벨 변경"), ids)
        menu.addSeparator()
        if len(ids) == 1:
            menu.addAction(".msg 열기 (Outlook)", self.open_msg)
            if self.log_dir is not None:
                menu.addAction("본문 HTML 구조 저장 (진단용, 내용 제외)", self.save_html_shape)
        menu.addAction(f"삭제 ({len(ids)}건)\tDel", self.delete_selected)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def save_html_shape(self) -> None:
        """미리보기 모양이 Outlook 과 다를 때 원인을 찾으려고, 본문 HTML 의 태그·스타일만 남긴 사본을 logs 에 저장한다."""
        ids = self.selected_ids()
        mail = self.lib.get(ids[0]) if len(ids) == 1 else None
        if mail is None:
            return
        mail = self.lib.with_html(mail)
        if not mail.html_body:
            QMessageBox.information(self, "inlook", "이 메일은 HTML 본문이 없습니다 (텍스트 본문).")
            return
        self.log_dir.mkdir(parents=True, exist_ok=True)
        path = self.log_dir / f"html-shape-{mail.id}.html"
        path.write_text(html_shape(mail.html_body), encoding="utf-8")
        QMessageBox.information(
            self, "inlook",
            f"저장했습니다:\n{path}\n\n글자는 x 로 바뀌고 링크·이미지 주소·작성자 정보는 지웠습니다.\n"
            "공유하기 전에 한 번 열어 확인하세요.",
        )

    def delete_selected(self) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        answer = QMessageBox.warning(
            self,
            "메일 삭제",
            f"선택한 {len(ids)}건을 삭제할까요?\n\n"
            ".msg 파일과 DB 기록이 모두 지워지고 되돌릴 수 없습니다.\n"
            "Outlook 에 남아 있어도 다음 동기화에서 다시 가져오지 않습니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        failed = self.lib.delete_mails(ids)
        self.refresh_all()
        if failed:
            QMessageBox.warning(
                self, "메일 삭제",
                f"{len(failed)}건은 .msg 파일을 지우지 못해 남겨 두었습니다.\n"
                "Outlook 등에서 그 파일을 열고 있다면 닫은 뒤 다시 삭제하세요.",
            )

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


def run_app(
    root: Path, make_source: Callable[[], MailSource], log_dir: Path | None, settings_path: Path | None = None
) -> int:
    if sys.platform == "win32":
        # pythonw 로 띄우면 작업 표시줄에 Python 아이콘이 나온다. 이 프로세스에만 앱 ID 를 붙여 inlook 아이콘으로
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("xunyss.inlook")
        except Exception:
            pass
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("inlook")
    app.setWindowIcon(theme.app_icon())
    if sys.platform == "win32":
        app.setFont(QFont("Malgun Gothic", 9))

    def excepthook(kind, value, tb) -> None:
        text = "".join(traceback.format_exception(kind, value, tb))
        if log_dir is not None:
            log_dir.mkdir(parents=True, exist_ok=True)
            with (log_dir / "app-error.txt").open("a", encoding="utf-8") as f:
                f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S}\n{text}\n")
        QMessageBox.critical(None, "inlook 오류", text[-2000:])

    sys.excepthook = excepthook
    try:
        window = MainWindow(root, make_source, log_dir, settings_path)
    except Exception as exc:
        QMessageBox.critical(None, "inlook", f"보존 경로를 열 수 없습니다: {root}\n\n{exc}")
        return 2
    window.show()
    return app.exec()
