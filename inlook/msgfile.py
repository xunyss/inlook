""".msg 파일(OLE 복합 문서)에서 첨부를 읽는다. Outlook 없이 동작한다 (olefile).

본문 HTML 의 이미지는 <img src="cid:image001.png@01DB..."> 처럼 첨부를 가리키므로,
미리보기에서 Content-ID 로 첨부 내용을 찾아 보여준다.

MAPI 속성 스트림 이름: __substg1.0_<태그 4자리><형식 4자리>
  3701 0102 첨부 데이터 (바이너리)
  3712 001F Content-ID (유니코드)
  3707 001F 긴 파일명, 3704 001F 짧은 파일명
"""

from __future__ import annotations

import re
import struct
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

import olefile

from .rtfhtml import decompress_rtf, html_from_compressed_rtf

_ATTACH_PREFIX = "__attach_version1.0_"
_DATA = "__substg1.0_37010102"
_CONTENT_ID = "__substg1.0_3712001F"
_LONG_NAME = "__substg1.0_3707001F"
_SHORT_NAME = "__substg1.0_3704001F"


@dataclass(frozen=True)
class Attachment:
    filename: str
    content_id: str  # <> 를 뗀 값. 없으면 ""
    data: bytes


def _text(ole: olefile.OleFileIO, path: list[str]) -> str:
    if not ole.exists("/".join(path)):
        return ""
    return ole.openstream(path).read().decode("utf-16-le", errors="replace").rstrip("\x00")


def _attachments(ole: olefile.OleFileIO) -> list[Attachment]:
    result: list[Attachment] = []
    storages = sorted({e[0] for e in ole.listdir(streams=True, storages=True) if e[0].startswith(_ATTACH_PREFIX)})
    for storage in storages:
        if not ole.exists(f"{storage}/{_DATA}"):
            continue
        data = ole.openstream([storage, _DATA]).read()
        name = _text(ole, [storage, _LONG_NAME]) or _text(ole, [storage, _SHORT_NAME])
        cid = _text(ole, [storage, _CONTENT_ID]).strip().strip("<>")
        result.append(Attachment(name, cid, data))
    return result


def read_attachments(path: Path) -> list[Attachment]:
    """파일 데이터가 있는 첨부만 (메일 안의 메일 같은 내장 개체는 뺀다)."""
    with olefile.OleFileIO(str(path)) as ole:
        return _attachments(ole)


def inline_images(path: Path) -> dict[str, bytes]:
    """'cid:' 뒤에 올 수 있는 값 → 데이터. Content-ID 와, Content-ID 가 없을 때를 대비해 파일명으로도 찾는다."""
    images: dict[str, bytes] = {}
    for att in read_attachments(path):
        if att.content_id:
            images[att.content_id] = att.data
            images.setdefault(att.content_id.split("@", 1)[0], att.data)
        if att.filename:
            images.setdefault(att.filename, att.data)
    return images


# --- 메일 자체 속성 (가져오기용) ---

_PROPS_STREAM = "__properties_version1.0"
_TOP_HEADER = 32  # 최상위 메시지의 고정 길이 속성 스트림 머리말 크기
_FILETIME_EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)

PR_SUBJECT = 0x0037
PR_MESSAGE_CLASS = 0x001A
PR_CLIENT_SUBMIT_TIME = 0x0039  # 보낸 시각
PR_MESSAGE_DELIVERY_TIME = 0x0E06  # 받은 시각
PR_DISPLAY_TO = 0x0E04
PR_DISPLAY_CC = 0x0E03
PR_SENDER_NAME = 0x0C1A
PR_SENDER_EMAIL = 0x0C1F
PR_SENDER_SMTP = 0x5D01
PR_BODY = 0x1000
PR_HTML = 0x1013
PR_RTF_COMPRESSED = 0x1009
PR_INTERNET_MESSAGE_ID = 0x1035
PR_SEARCH_KEY = 0x300B
PR_INTERNET_CPID = 0x3FDE
PR_MESSAGE_SIZE = 0x0E08

_CPID = {65001: "utf-8", 949: "cp949", 51949: "euc-kr", 50225: "iso2022_kr", 1200: "utf-16-le", 20127: "ascii"}
_META_CHARSET = re.compile(rb"charset\s*=\s*[\"']?([A-Za-z0-9_\-]+)", re.IGNORECASE)


@dataclass(frozen=True)
class MsgMessage:
    subject: str
    message_class: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    received: datetime | None  # naive 로컬 시각 (COM 으로 수집한 값과 같은 기준)
    sent: datetime | None
    message_id: str | None
    search_key: str | None  # hex 대문자 (COM Table 값과 같은 형식)
    body: str
    html_body: str
    attachments: list[str]
    size: int


def _stream(ole: olefile.OleFileIO, name: str) -> bytes | None:
    return ole.openstream(name).read() if ole.exists(name) else None


def _string(ole: olefile.OleFileIO, prop: int, ansi: str) -> str:
    data = _stream(ole, f"__substg1.0_{prop:04X}001F")
    if data is not None:
        return data.decode("utf-16-le", errors="replace").rstrip("\x00")
    data = _stream(ole, f"__substg1.0_{prop:04X}001E")  # 예전 ANSI 형식
    if data is not None:
        return data.decode(ansi, errors="replace").rstrip("\x00")
    return ""


def _fixed_props(ole: olefile.OleFileIO) -> dict[int, tuple[int, bytes]]:
    """고정 길이 속성: 속성 ID → (형식, 8바이트 값)."""
    data = _stream(ole, _PROPS_STREAM) or b""
    props: dict[int, tuple[int, bytes]] = {}
    for offset in range(_TOP_HEADER, len(data) - 15, 16):
        tag = struct.unpack_from("<I", data, offset)[0]
        props[tag >> 16] = (tag & 0xFFFF, data[offset + 8:offset + 16])
    return props


def _filetime(props: dict[int, tuple[int, bytes]], prop: int) -> datetime | None:
    entry = props.get(prop)
    if entry is None or entry[0] != 0x0040:
        return None
    ticks = struct.unpack("<Q", entry[1])[0]
    if ticks == 0:
        return None
    utc = _FILETIME_EPOCH + timedelta(microseconds=ticks // 10)
    if utc.year >= 4500:
        return None
    return utc.astimezone().replace(tzinfo=None, microsecond=0)  # 이 PC 의 시간대 기준 로컬 시각


def _int(props: dict[int, tuple[int, bytes]], prop: int) -> int | None:
    entry = props.get(prop)
    if entry is None or entry[0] != 0x0003:
        return None
    return struct.unpack_from("<i", entry[1])[0]


def _decode_html(data: bytes, cpid: int | None) -> str:
    candidates = []
    if cpid in _CPID:
        candidates.append(_CPID[cpid])
    elif cpid:
        candidates.append(f"cp{cpid}")
    m = _META_CHARSET.search(data[:4096])
    if m:
        candidates.append(m.group(1).decode("ascii"))
    candidates += ["utf-8", "cp949"]
    for encoding in candidates:
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def _html(ole: olefile.OleFileIO, cpid: int | None, ansi: str) -> str:
    """PR_HTML, 없으면 압축 RTF 안에 감싼 HTML (회의 요청 등). 둘 다 없으면 ""."""
    raw = _stream(ole, f"__substg1.0_{PR_HTML:04X}0102")
    if raw:
        return _decode_html(raw.rstrip(b"\x00"), cpid)
    html = _string(ole, PR_HTML, ansi)
    if html:
        return html
    rtf = _stream(ole, f"__substg1.0_{PR_RTF_COMPRESSED:04X}0102")
    if rtf:
        try:
            return html_from_compressed_rtf(rtf) or ""
        except ValueError:
            return ""
    return ""


def read_html(path: Path) -> str:
    """.msg 의 HTML 본문. 저장할 때 HTML 을 못 받은 메일(회의 요청 등)의 미리보기에 쓴다."""
    with olefile.OleFileIO(str(path)) as ole:
        cpid = _int(_fixed_props(ole), PR_INTERNET_CPID)
        return _html(ole, cpid, _CPID.get(cpid, "cp949") if cpid else "cp949")


def read_message(path: Path) -> MsgMessage:
    with olefile.OleFileIO(str(path)) as ole:
        props = _fixed_props(ole)
        cpid = _int(props, PR_INTERNET_CPID)
        ansi = _CPID.get(cpid, "cp949") if cpid else "cp949"
        html = _html(ole, cpid, ansi)
        search_key = _stream(ole, f"__substg1.0_{PR_SEARCH_KEY:04X}0102")
        message_id = _string(ole, PR_INTERNET_MESSAGE_ID, ansi).strip()
        attachments = [a.filename for a in _attachments(ole)]
        sent = _filetime(props, PR_CLIENT_SUBMIT_TIME)
        return MsgMessage(
            subject=_string(ole, PR_SUBJECT, ansi),
            message_class=_string(ole, PR_MESSAGE_CLASS, ansi) or "IPM.Note",
            sender_name=_string(ole, PR_SENDER_NAME, ansi),
            sender_email=_string(ole, PR_SENDER_SMTP, ansi) or _string(ole, PR_SENDER_EMAIL, ansi),
            to=_string(ole, PR_DISPLAY_TO, ansi),
            cc=_string(ole, PR_DISPLAY_CC, ansi),
            received=_filetime(props, PR_MESSAGE_DELIVERY_TIME) or sent,
            sent=sent,
            message_id=message_id or None,
            search_key=search_key.hex().upper() if search_key else None,
            body=_string(ole, PR_BODY, ansi),
            html_body=html,
            attachments=attachments,
            size=_int(props, PR_MESSAGE_SIZE) or path.stat().st_size,
        )


def body_format(path: Path) -> str:
    """본문이 어떤 형태로 들어 있는지 (진단용).
    html: PR_HTML / rtf-html: 압축 RTF 안에 감싼 HTML / rtf-only: HTML 을 감싸지 않은 순수 RTF /
    text-only: 텍스트 본문만 / none: 본문 없음"""
    with olefile.OleFileIO(str(path)) as ole:
        if _stream(ole, f"__substg1.0_{PR_HTML:04X}0102") or _stream(ole, f"__substg1.0_{PR_HTML:04X}001F"):
            return "html"
        rtf = _stream(ole, f"__substg1.0_{PR_RTF_COMPRESSED:04X}0102")
        if rtf:
            try:
                raw = decompress_rtf(rtf)
            except ValueError:
                return "rtf-broken"
            return "rtf-html" if b"\\fromhtml" in raw[:1024] else "rtf-only"
        if _stream(ole, f"__substg1.0_{PR_BODY:04X}001F") or _stream(ole, f"__substg1.0_{PR_BODY:04X}001E"):
            return "text-only"
        return "none"
