"""Outlook COM 으로 구현한 MailSource. Windows(VDI) 전용.

이 파일은 Mac 에서 테스트할 수 없으므로 최대한 얇게 유지한다.
로직은 collector/probe 쪽에 두고, 여기서는 COM 값을 source 타입으로 옮기기만 한다.
"""

from __future__ import annotations

import platform
import sys
import time
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
    MailHeader,
    MailRecord,
    StoreRef,
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
# OlExchangeConnectionMode
_CONNECTION_MODES = {
    0: "no-exchange",
    100: "offline",
    200: "cached-offline",
    300: "disconnected",
    400: "cached-disconnected",
    500: "cached-headers",
    600: "cached-drizzle",
    700: "cached-full",
    800: "online",
}
# OlExchangeStoreType
_STORE_TYPES = {0: "primary", 1: "mailbox", 2: "public", 3: "not-exchange", 4: "additional"}

OL_MAIL_ITEM = 0  # OlItemType: Folder.DefaultItemType 가 이 값이면 메일 폴더
OL_USER_ITEMS = 0  # OlTableContents
OL_MSG_UNICODE = 9  # OlSaveAsType. olMSG(3)는 ANSI 라 한글이 깨질 수 있다
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"

# Table 열. 순서가 _header() 의 인덱스와 맞아야 한다.
# 기본 속성명(ReceivedTime 등)은 로컬 시각으로 돌아온다.
_TABLE_COLUMNS = (
    "EntryID",
    "MessageClass",
    "Subject",
    "SenderName",
    "ReceivedTime",
    "Size",
    PR_INTERNET_MESSAGE_ID,
    PR_SEARCH_KEY,
)

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


def _to_hex(value) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        return value
    try:
        return bytes(value).hex().upper()
    except (TypeError, ValueError):
        return str(value)


class OutlookSource:
    def __init__(self) -> None:
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        inbox = self.ns.GetDefaultFolder(6)
        self._default_root = inbox.Parent
        self._default_store_id = str(_get(inbox, "StoreID", ""))
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
        mode = _get(self.ns, "ExchangeConnectionMode")
        return {
            "python": sys.version.split()[0],
            "pywin32": pywin32_version,
            "os": platform.platform(),
            "outlook": str(_get(self.app, "Version", "?")),
            "mode": _CONNECTION_MODES.get(mode, str(mode)),
        }

    def stores(self) -> Iterator[StoreRef]:
        for index, store in enumerate(self.ns.Stores, 1):
            yield StoreRef(
                index=index,
                is_default=str(_get(store, "StoreID", "")) == self._default_store_id,
                kind=_STORE_TYPES.get(_get(store, "ExchangeStoreType"), "?"),
                is_data_file=bool(_get(store, "IsDataFileStore", False)),
                handle=store,
            )

    def folders(self, store: StoreRef | None = None) -> Iterator[FolderRef]:
        if store is None or store.is_default:
            root, index = self._default_root, (store.index if store else 1)
        else:
            root, index = store.handle.GetRootFolder(), store.index
        yield from self._walk(root, (), index)

    def _walk(self, folder, path: tuple[str, ...], store_index: int) -> Iterator[FolderRef]:
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
                store_index=store_index,
                handle=sub,
            )
            yield from self._walk(sub, sub_path, store_index)

    def headers(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
        oldest_first: bool = False,
    ) -> Iterator[MailHeader]:
        # Restrict/Table 필터의 날짜 문자열은 로케일 의존이라 쓰지 않는다.
        # 최신순 정렬 후 기준일보다 오래된 행이 나오면 멈춘다.
        store_id = str(_get(folder.handle, "StoreID", ""))
        table = folder.handle.GetTable("", OL_USER_ITEMS)
        columns = table.Columns
        columns.RemoveAll()
        for name in _TABLE_COLUMNS:
            columns.Add(name)
        table.Sort("[ReceivedTime]", not oldest_first)
        count = 0
        while not table.EndOfTable:
            values = table.GetNextRow().GetValues()
            header = MailHeader(
                entry_id=str(values[0] or ""),
                store_id=store_id,
                message_class=str(values[1] or ""),
                subject=str(values[2] or ""),
                sender_name=str(values[3] or ""),
                received=_to_datetime(values[4]),
                size=int(values[5] or 0),
                message_id=str(values[6]) if values[6] else None,
                search_key=_to_hex(values[7]),
            )
            if since and not oldest_first and header.received and header.received < since:
                break
            yield header
            count += 1
            if limit and count >= limit:
                break

    def _item(self, header: MailHeader):
        return self.ns.GetItemFromID(header.entry_id, header.store_id)

    def open(self, header: MailHeader) -> MailRecord:
        item = self._item(header)
        attachments = [str(_get(a, "FileName", "")) for a in (_get(item, "Attachments", None) or [])]
        return MailRecord(
            header=header,
            sender_email=str(_get(item, "SenderEmailAddress", "") or ""),
            to=str(_get(item, "To", "") or ""),
            cc=str(_get(item, "CC", "") or ""),
            sent=_to_datetime(_get(item, "SentOn")),
            body=str(_get(item, "Body", "") or ""),
            html_body=str(_get(item, "HTMLBody", "") or ""),
            attachments=attachments,
            handle=item,
        )

    def measure(self, header: MailHeader) -> dict[str, float]:
        result: dict[str, float] = {}

        def step(name: str, fn) -> None:
            t = time.perf_counter()
            try:
                fn()
            except COM_ERRORS:
                pass
            result[name] = time.perf_counter() - t

        holder: dict[str, object] = {}
        step("open", lambda: holder.__setitem__("item", self._item(header)))
        item = holder.get("item")
        if item is None:
            return result
        step("basic", lambda: (item.Subject, item.ReceivedTime, item.SenderName, item.To, item.CC))
        step("sender_smtp", lambda: item.SenderEmailAddress)
        step("body", lambda: item.Body)
        step("html", lambda: item.HTMLBody)
        step("attachments", lambda: [a.FileName for a in item.Attachments])
        return result

    def save_msg(self, record: MailRecord, path: Path) -> None:
        record.handle.SaveAs(str(path), OL_MSG_UNICODE)
