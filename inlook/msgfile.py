""".msg 파일(OLE 복합 문서)에서 첨부를 읽는다. Outlook 없이 동작한다 (olefile).

본문 HTML 의 이미지는 <img src="cid:image001.png@01DB..."> 처럼 첨부를 가리키므로,
미리보기에서 Content-ID 로 첨부 내용을 찾아 보여준다.

MAPI 속성 스트림 이름: __substg1.0_<태그 4자리><형식 4자리>
  3701 0102 첨부 데이터 (바이너리)
  3712 001F Content-ID (유니코드)
  3707 001F 긴 파일명, 3704 001F 짧은 파일명
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import olefile

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


def read_attachments(path: Path) -> list[Attachment]:
    """파일 데이터가 있는 첨부만 (메일 안의 메일 같은 내장 개체는 뺀다)."""
    result: list[Attachment] = []
    with olefile.OleFileIO(str(path)) as ole:
        storages = sorted({e[0] for e in ole.listdir(streams=True, storages=True) if e[0].startswith(_ATTACH_PREFIX)})
        for storage in storages:
            if not ole.exists(f"{storage}/{_DATA}"):
                continue
            data = ole.openstream([storage, _DATA]).read()
            name = _text(ole, [storage, _LONG_NAME]) or _text(ole, [storage, _SHORT_NAME])
            cid = _text(ole, [storage, _CONTENT_ID]).strip().strip("<>")
            result.append(Attachment(name, cid, data))
    return result


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
