"""화면 테마: 밝은/어두운 색 팔레트, 스타일시트, 아이콘.

OS 기본 스타일(Windows 11 등)은 표의 선택 줄을 칸마다 따로 그리고 OS 마다 모양이 달라서,
Fusion 스타일 위에 직접 정한 색을 입힌다. Mac 과 VDI 에서 같은 모양으로 보인다.
아이콘은 Lucide(ISC 라이선스, icons/LICENSE) 의 SVG 를 테마 색으로 칠해서 쓴다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from PySide6.QtCore import QByteArray, QRectF, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPalette, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication, QStyleFactory

ICON_DIR = Path(__file__).resolve().parent / "icons"
APP_ICON = Path(__file__).resolve().parent / "app.svg"  # 앱 아이콘 (Outlook 패러디). app.ico 는 이 파일로 만든다
LIGHT, DARK = "light", "dark"


@dataclass(frozen=True)
class Colors:
    window: str  # 패널 사이 바탕
    surface: str  # 패널(카드)
    raised: str  # 입력창·배지·칩
    border: str
    text: str
    muted: str
    accent: str
    accent_text: str
    hover: str
    select_inactive: str  # 목록에 초점이 없을 때 선택 줄
    danger: str
    paper: str  # 흰 배경 원본 보기의 본문 바탕


COLORS = {
    LIGHT: Colors(
        window="#e9ecf1", surface="#ffffff", raised="#f1f3f6", border="#dde1e7", text="#1f2329",
        muted="#6b7380", accent="#3d6ee0", accent_text="#ffffff", hover="#eef2f8",
        select_inactive="#dbe4f8", danger="#d33a3a", paper="#ffffff",
    ),
    DARK: Colors(
        window="#15171b", surface="#1e2126", raised="#292d33", border="#30343b", text="#e4e7eb",
        muted="#959ca7", accent="#4c7ced", accent_text="#ffffff", hover="#262a30",
        select_inactive="#2b3650", danger="#ff6b6b", paper="#ffffff",
    ),
}


def colors(mode: str) -> Colors:
    return COLORS[mode]


def system_mode() -> str:
    hints = QApplication.styleHints()
    if hasattr(hints, "colorScheme") and hints.colorScheme() == Qt.ColorScheme.Dark:
        return DARK
    return LIGHT


# --- 설정 파일 (테마 선택 기억) ---


def load_mode(path: Path | None) -> str:
    try:
        mode = json.loads(path.read_text(encoding="utf-8")).get("theme") if path else None
    except (OSError, ValueError, AttributeError):
        mode = None
    return mode if mode in COLORS else system_mode()


def save_mode(path: Path | None, mode: str) -> None:
    if path is None:
        return
    try:
        data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
        data["theme"] = mode
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, ValueError):
        pass  # 기억 못 해도 화면은 동작한다


# --- 적용 ---


def apply(app: QApplication, mode: str) -> None:
    c = colors(mode)
    app.setStyle(QStyleFactory.create("Fusion"))
    pal = QPalette()
    roles = {
        QPalette.ColorRole.Window: c.window,
        QPalette.ColorRole.WindowText: c.text,
        QPalette.ColorRole.Base: c.surface,
        QPalette.ColorRole.AlternateBase: c.raised,
        QPalette.ColorRole.Text: c.text,
        QPalette.ColorRole.Button: c.raised,
        QPalette.ColorRole.ButtonText: c.text,
        QPalette.ColorRole.Highlight: c.accent,
        QPalette.ColorRole.HighlightedText: c.accent_text,
        QPalette.ColorRole.ToolTipBase: c.surface,
        QPalette.ColorRole.ToolTipText: c.text,
        QPalette.ColorRole.PlaceholderText: c.muted,
        QPalette.ColorRole.Link: c.accent,
        QPalette.ColorRole.Mid: c.border,
        QPalette.ColorRole.Midlight: c.border,
        QPalette.ColorRole.Dark: c.border,
    }
    for role, value in roles.items():
        pal.setColor(role, QColor(value))
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        pal.setColor(QPalette.ColorGroup.Disabled, role, QColor(c.muted))
    app.setPalette(pal)
    app.setStyleSheet(stylesheet(c))


def stylesheet(c: Colors) -> str:
    return f"""
QMainWindow, QDialog {{ background: {c.window}; }}
QToolTip {{ background: {c.surface}; color: {c.text}; border: 1px solid {c.border}; padding: 4px 6px; }}

QToolBar#main {{ background: {c.window}; border: none; padding: 6px 8px 2px 8px; spacing: 4px; }}
QToolBar#main QToolButton {{ background: transparent; border: none; border-radius: 6px; padding: 5px 8px; color: {c.text}; }}
QToolBar#main QToolButton:hover {{ background: {c.hover}; }}
QToolBar#main QToolButton:pressed {{ background: {c.raised}; }}
QToolBar#main QToolButton:disabled {{ color: {c.muted}; }}
QToolBar#main QToolButton#primary {{ background: {c.accent}; color: {c.accent_text}; padding: 5px 12px; }}
QToolBar#main QToolButton#primary:hover {{ background: {c.accent}; }}
QToolBar#main QToolButton::menu-indicator {{ image: none; width: 0; }}

QLineEdit#search {{ background: {c.surface}; border: 1px solid {c.border}; border-radius: 8px; padding: 6px 10px; color: {c.text}; }}
QLineEdit#search:focus {{ border-color: {c.accent}; }}

QFrame#card {{ background: {c.surface}; border: 1px solid {c.border}; border-radius: 10px; }}
QSplitter::handle {{ background: transparent; }}

QTreeView, QTableView, QTextBrowser, QPlainTextEdit {{ background: {c.surface}; border: none; color: {c.text}; outline: 0; }}
QTreeView {{ padding: 4px; }}
QTreeView::item {{ padding: 4px 2px; border-radius: 6px; }}
QTreeView::item:hover {{ background: {c.hover}; }}
QTreeView::item:selected {{ background: {c.select_inactive}; color: {c.text}; }}
QTreeView::branch, QTreeView::branch:selected {{ background: transparent; }}

QTableView {{ selection-background-color: {c.accent}; selection-color: {c.accent_text}; }}
QTableView::item {{ padding: 0 6px; border: none; }}
QTableView::item:selected:active {{ background: {c.accent}; color: {c.accent_text}; }}
QTableView::item:selected:!active {{ background: {c.select_inactive}; color: {c.text}; }}
QHeaderView {{ background: {c.surface}; border: none; }}
QHeaderView::section {{ background: {c.surface}; color: {c.muted}; border: none;
    border-bottom: 1px solid {c.border}; padding: 6px 6px; font-weight: 600; }}

QLabel#title {{ font-size: 15pt; font-weight: 700; color: {c.text}; }}
QLabel#sender {{ font-weight: 600; color: {c.text}; }}
QLabel#muted, QLabel#sectionTitle {{ color: {c.muted}; }}
QLabel#sectionTitle {{ font-weight: 700; padding: 8px 10px 2px 10px; }}
QFrame#chip {{ background: {c.raised}; border: 1px solid {c.border}; border-radius: 9px; }}
QFrame#labelChip {{ background: {c.select_inactive}; border: 1px solid {c.select_inactive}; border-radius: 9px; }}
QFrame#chip QLabel, QFrame#labelChip QLabel {{ background: transparent; border: none; color: {c.text}; }}
QLabel#avatar {{ border-radius: 18px; color: white; font-weight: 700; }}
QFrame#divider {{ background: {c.border}; max-height: 1px; min-height: 1px; border: none; }}
QToolButton#flat {{ background: transparent; border: none; border-radius: 6px; padding: 4px; color: {c.text}; }}
QToolButton#flat:hover {{ background: {c.hover}; }}
QToolButton#flat:checked {{ background: {c.raised}; }}

QStatusBar {{ background: {c.window}; color: {c.muted}; border: none; }}
QStatusBar::item {{ border: none; }}
QStatusBar QLabel {{ color: {c.muted}; padding: 0 8px; }}
QStatusBar QLabel#stale {{ color: {c.danger}; font-weight: 700; }}

QPushButton {{ background: {c.raised}; border: 1px solid {c.border}; border-radius: 6px; padding: 5px 14px; color: {c.text}; }}
QPushButton:hover {{ background: {c.hover}; }}
QPushButton:disabled {{ color: {c.muted}; }}
QProgressBar {{ background: {c.raised}; border: none; border-radius: 4px; height: 8px; text-align: center; color: transparent; }}
QProgressBar::chunk {{ background: {c.accent}; border-radius: 4px; }}

QMenu {{ background: {c.surface}; border: 1px solid {c.border}; padding: 4px; color: {c.text}; }}
QMenu::item {{ padding: 5px 24px 5px 24px; border-radius: 4px; }}
QMenu::item:selected {{ background: {c.accent}; color: {c.accent_text}; }}
QMenu::separator {{ height: 1px; background: {c.border}; margin: 4px 8px; }}

QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {c.border}; border-radius: 3px; min-height: 30px; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {c.border}; border-radius: 3px; min-width: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
"""


# --- 아이콘 ---


@lru_cache(maxsize=256)
def _pixmap(name: str, color: str, size: int) -> QPixmap:
    svg = (ICON_DIR / f"{name}.svg").read_text(encoding="utf-8").replace("currentColor", color)
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    ratio = 2  # 고배율 화면에서도 선명하게
    pixmap = QPixmap(size * ratio, size * ratio)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter, QRectF(0, 0, size * ratio, size * ratio))
    painter.end()
    pixmap.setDevicePixelRatio(ratio)
    return pixmap


def icon(name: str, color: str, size: int = 18, selected: str | None = None) -> QIcon:
    """selected: 선택된 줄(강조색 바탕) 위에서 쓸 색."""
    result = QIcon(_pixmap(name, color, size))
    if selected:
        result.addPixmap(_pixmap(name, selected, size), QIcon.Mode.Selected)
    return result


def app_icon() -> QIcon:
    """창·작업 표시줄 아이콘. 크기별로 미리 그려 둔다 (작은 크기에서도 선명하게)."""
    renderer = QSvgRenderer(str(APP_ICON))
    result = QIcon()
    for size in (16, 24, 32, 48, 64, 128, 256):
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        renderer.render(painter, QRectF(0, 0, size, size))
        painter.end()
        result.addPixmap(pixmap)
    return result
