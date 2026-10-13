"""Outlook COM 으로 구현한 MailSource. Windows(VDI) 전용.

이 파일은 Mac 에서 테스트할 수 없으므로 최대한 얇게 유지한다.
로직은 collector/probe 쪽에 두고, 여기서는 COM 값을 MailRecord 로 옮기기만 한다.
"""

from __future__ import annotations

import platform
import sys
from collections.abc import Iterator
from datetime import datetime
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

import pywintypes  # type: ignore[import-not-found]
import win32com.client  # type: ignore[import-not-found]

from .source import (
    KIND_DELETED,
    KIND_DRAFTS,
    KIND_INBOX,
    KIND_JUNK,
    KIND_OUTBOX,
    KIND_SENT,
    FolderRef,
    MailRecord,
)

# OlDefaultFolders
_DEFAULT_FOLDERS = {
    6: KIND_INBOX,
    5: KIND_SENT,
    16: KIND_DRAFTS,
    4: KIND_OUTBOX,
    3: KIND_DELETED,
    23: KIND_JUNK,
}
OL_MAIL_ITEM = 0  # OlItemType: Folder.DefaultItemType 가 이 값이면 메일 폴더
OL_MSG_UNICODE = 9  # OlSaveAsType. olMSG(3)는 ANSI 라 한글이 깨질 수 있다
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

COM_ERRORS = (pywintypes.com_error, AttributeError)


def _get(obj, name: str, default=None):
    try:
        return getattr(obj, name)
    except COM_ERRORS:
        return default


def _to_datetime(value) -> datetime | None:
    """pywintypes 시각 → naive local datetime.

    pywin32 는 로컬 시각에 UTC tzinfo 를 붙여서 돌려주므로 tzinfo 만 떼어낸다.
    Outlook 은 '값 없음'을 4501-01-01 로 표현한다.
    """
    if value is None:
        return None
    try:
        if value.year >= 4500:
            return None
        return datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        )
    except (AttributeError, ValueError):
        return None


class OutlookSource:
    def __init__(self) -> None:
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self._kinds: dict[str, str] = {}
        for number, kind in _DEFAULT_FOLDERS.items():
            try:
                self._kinds[self.ns.GetDefaultFolder(number).EntryID] = kind
            except COM_ERRORS:
                pass

    def describe(self) -> dict[str, str]:
        try:
            pywin32_version = version("pywin32")
        except PackageNotFoundError:
            pywin32_version = "?"
        return {
            "python": sys.version.split()[0],
            "pywin32": pywin32_version,
            "os": platform.platform(),
            "outlook": str(_get(self.app, "Version", "?")),
            "stores": str(_get(_get(self.ns, "Stores"), "Count", "?")),
        }

    def folders(self) -> Iterator[FolderRef]:
        root = self.ns.GetDefaultFolder(6).Parent  # 사서함 최상위
        yield from self._walk(root, ())

    def _walk(self, folder, path: tuple[str, ...]) -> Iterator[FolderRef]:
        for sub in folder.Folders:
            sub_path = path + (str(_get(sub, "Name", "?")),)
            try:
                count = sub.Items.Count
            except COM_ERRORS:
                count = -1  # 접근 불가
            yield FolderRef(
                path=sub_path,
                item_count=count,
                is_mail=_get(sub, "DefaultItemType") == OL_MAIL_ITEM,
                kind=self._kinds.get(_get(sub, "EntryID", "")),
                handle=sub,
            )
            yield from self._walk(sub, sub_path)

    def items(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[MailRecord]:
        # Restrict 의 날짜 문자열은 로케일에 따라 해석이 달라서 쓰지 않는다.
        # 최신순 정렬 후 기준일보다 오래된 항목이 나오면 멈춘다.
        items = folder.handle.Items
        items.Sort("[ReceivedTime]", True)
        count = 0
        item = items.GetFirst()
        while item is not None:
            record = self._record(item)
            if since and record.received and record.received < since:
                break
            yield record
            count += 1
            if limit and count >= limit:
                break
            item = items.GetNext()

    def _record(self, item) -> MailRecord:
        try:
            message_id = item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) or None
        except COM_ERRORS:
            message_id = None
        attachments: list[str] = []
        for attachment in _get(item, "Attachments", None) or []:
            attachments.append(str(_get(attachment, "FileName", "")))
        return MailRecord(
            message_class=str(_get(item, "MessageClass", "")),
            subject=str(_get(item, "Subject", "") or ""),
            sender_name=str(_get(item, "SenderName", "") or ""),
            sender_email=str(_get(item, "SenderEmailAddress", "") or ""),
            to=str(_get(item, "To", "") or ""),
            cc=str(_get(item, "CC", "") or ""),
            received=_to_datetime(_get(item, "ReceivedTime")),
            sent=_to_datetime(_get(item, "SentOn")),
            message_id=message_id,
            entry_id=str(_get(item, "EntryID", "")),
            body=str(_get(item, "Body", "") or ""),
            html_body=str(_get(item, "HTMLBody", "") or ""),
            attachments=attachments,
            size=int(_get(item, "Size", 0) or 0),
            handle=item,
        )

    def save_msg(self, item: MailRecord, path: Path) -> None:
        item.handle.SaveAs(str(path), OL_MSG_UNICODE)
