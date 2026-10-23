"""테스트용 최소 OLE 복합 문서(.msg 형식) 작성기. 합성 데이터로 msgfile 을 시험하려고 쓴다.

512바이트 섹터(v3). 4096바이트 미만 스트림은 표준대로 미니 스트림(64바이트 단위)에 넣는다.
형제 노드는 오른쪽 링크로만 잇는다 (olefile 은 균형 트리를 요구하지 않는다).
"""

from __future__ import annotations

import struct
from pathlib import Path

SECTOR, MINI, CUTOFF = 512, 64, 4096
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF

Tree = dict[str, "bytes | Tree"]


def _ceil(n: int, unit: int) -> int:
    return -(-n // unit)


def write_cfb(path: Path, tree: Tree) -> None:
    entries: list[dict] = [{"name": "Root Entry", "type": 5, "kids": [], "data": b""}]

    def add(node: Tree, parent: int) -> None:
        for name, value in node.items():
            index = len(entries)
            is_storage = isinstance(value, dict)
            entries.append({"name": name, "type": 1 if is_storage else 2, "kids": [],
                            "data": b"" if is_storage else value})
            entries[parent]["kids"].append(index)
            if is_storage:
                add(value, index)

    add(tree, 0)

    # 1) 작은 스트림 → 미니 스트림
    mini_stream = b""
    minifat: list[int] = []
    for e in entries[1:]:
        e["start"] = ENDOFCHAIN
        if e["type"] != 2 or not e["data"] or len(e["data"]) >= CUTOFF:
            continue
        n = _ceil(len(e["data"]), MINI)
        e["start"] = len(minifat)
        minifat += [len(minifat) + i + 1 for i in range(n - 1)] + [ENDOFCHAIN]
        mini_stream += e["data"].ljust(n * MINI, b"\0")
    entries[0]["data"] = mini_stream  # 루트 항목의 데이터 = 미니 스트림

    # 2) 일반 섹터에 들어갈 덩어리: 디렉터리, 미니FAT, 루트(미니 스트림), 큰 스트림
    blobs: list[tuple[str, bytes]] = []
    directory_len = len(entries) * 128
    blobs.append(("dir", b"\0" * _ceil(directory_len, SECTOR) * SECTOR))
    if minifat:
        blobs.append(("minifat", struct.pack(f"<{len(minifat)}I", *minifat)))
    if mini_stream:
        blobs.append(("root", mini_stream))
    for i, e in enumerate(entries[1:], 1):
        if e["type"] == 2 and len(e["data"]) >= CUTOFF:
            blobs.append((f"e{i}", e["data"]))

    data_sectors = sum(_ceil(len(b), SECTOR) for _, b in blobs)
    fat_sectors = 1
    while fat_sectors * (SECTOR // 4) < fat_sectors + data_sectors:
        fat_sectors += 1
    fat = [FREESECT] * (fat_sectors * SECTOR // 4)
    for i in range(fat_sectors):
        fat[i] = FATSECT

    starts: dict[str, int] = {}
    counts: dict[str, int] = {}
    body = b""
    sector = fat_sectors
    for key, blob in blobs:
        n = _ceil(len(blob), SECTOR)
        starts[key], counts[key] = sector, n
        for i in range(n):
            fat[sector + i] = sector + i + 1 if i < n - 1 else ENDOFCHAIN
        sector += n
        body += blob.ljust(n * SECTOR, b"\0")

    entries[0]["start"] = starts.get("root", ENDOFCHAIN)
    for i, e in enumerate(entries[1:], 1):
        if f"e{i}" in starts:
            e["start"] = starts[f"e{i}"]

    siblings: dict[int, int] = {}
    for e in entries:
        kids = e["kids"]
        e["child"] = kids[0] if kids else NOSTREAM
        for i, kid in enumerate(kids):
            siblings[kid] = kids[i + 1] if i + 1 < len(kids) else NOSTREAM

    directory = b""
    for index, e in enumerate(entries):
        name = (e["name"] + "\0").encode("utf-16-le")
        directory += (
            name.ljust(64, b"\0")
            + struct.pack("<HBB", len(name), e["type"], 1)
            + struct.pack("<III", NOSTREAM, siblings.get(index, NOSTREAM), e["child"])
            + b"\0" * 36
            + struct.pack("<IQ", e["start"], len(e["data"]))
        )
    dir_start = (starts["dir"] - fat_sectors) * SECTOR
    body = body[:dir_start] + directory.ljust(counts["dir"] * SECTOR, b"\0") + body[dir_start + counts["dir"] * SECTOR:]

    header = (
        bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
        + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
        + struct.pack(
            "<IIIIIIIII", 0, fat_sectors, starts["dir"], 0, CUTOFF,
            starts.get("minifat", ENDOFCHAIN), counts.get("minifat", 0), ENDOFCHAIN, 0,
        )
        + struct.pack("<109I", *(list(range(fat_sectors)) + [FREESECT] * (109 - fat_sectors)))
    )
    assert len(header) == SECTOR
    path.write_bytes(header + struct.pack(f"<{len(fat)}I", *fat) + body)


def fake_msg(path: Path, attachments: list[tuple[str, str, bytes]]) -> None:
    """(파일명, Content-ID, 데이터) 첨부를 가진 .msg."""
    tree: Tree = {"__substg1.0_0037001F": "제목".encode("utf-16-le")}
    for i, (name, cid, data) in enumerate(attachments):
        storage: Tree = {"__substg1.0_37010102": data, "__substg1.0_3707001F": name.encode("utf-16-le")}
        if cid:
            storage["__substg1.0_3712001F"] = cid.encode("utf-16-le")
        tree[f"__attach_version1.0_#{i:08X}"] = storage
    write_cfb(path, tree)


def _filetime(value) -> bytes:
    from datetime import timezone

    utc = value.astimezone(timezone.utc) if value.tzinfo is None else value
    from datetime import datetime as _dt

    ticks = int((utc - _dt(1601, 1, 1, tzinfo=timezone.utc)).total_seconds() * 10_000_000)
    return struct.pack("<Q", ticks)


def fake_mail_msg(
    path: Path,
    *,
    subject: str,
    received=None,
    message_id: str | None = None,
    search_key: bytes | None = None,
    sender: str = "홍길동",
    sender_email: str = "gildong.hong@example.com",
    to: str = "me@example.com",
    body: str = "본문",
    html: str | None = None,
    html_charset: str = "utf-8",
    attachments: list[tuple[str, str, bytes]] | None = None,
    hidden: tuple[int, ...] = (),
    extra: dict[int, str] | None = None,
) -> None:
    """메일 속성까지 갖춘 합성 .msg. received 는 naive 로컬 시각.
    hidden: 숨김 표시할 첨부 번호, extra: 더 넣을 문자열 속성 (속성 ID → 값)."""

    def text(prop: int, value: str) -> tuple[str, bytes]:
        return f"__substg1.0_{prop:04X}001F", value.encode("utf-16-le")

    tree: Tree = dict([
        text(0x0037, subject), text(0x001A, "IPM.Note"), text(0x0C1A, sender), text(0x5D01, sender_email),
        text(0x0E04, to), text(0x1000, body),
    ])
    if message_id:
        tree.update([text(0x1035, message_id)])
    tree.update(text(prop, value) for prop, value in (extra or {}).items())
    if search_key:
        tree["__substg1.0_300B0102"] = search_key
    if html is not None:
        tree["__substg1.0_10130102"] = html.encode(html_charset)
    props = b"\0" * 32
    if received is not None:
        for prop in (0x0E06, 0x0039):
            props += struct.pack("<II", (prop << 16) | 0x0040, 0) + _filetime(received)
    tree["__properties_version1.0"] = props
    for i, (name, cid, data) in enumerate(attachments or []):
        storage: Tree = {"__substg1.0_37010102": data, "__substg1.0_3707001F": name.encode("utf-16-le")}
        if cid:
            storage["__substg1.0_3712001F"] = cid.encode("utf-16-le")
        if i in hidden:  # PR_ATTACHMENT_HIDDEN = true
            storage["__properties_version1.0"] = b"\0" * 8 + struct.pack("<II", (0x7FFE << 16) | 0x000B, 0) + struct.pack("<Q", 1)
        tree[f"__attach_version1.0_#{i:08X}"] = storage
    write_cfb(path, tree)
