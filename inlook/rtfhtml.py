"""압축 RTF(PR_RTF_COMPRESSED) → 그 안에 감싸인 HTML.

Outlook 은 회의 요청 같은 항목이나 일부 .msg 에 HTML 본문을 PR_HTML 로 따로 두지 않고,
압축 RTF 안에 '감싼 HTML'(\\fromhtml1) 로만 넣는다. 이 모듈은
  1) 압축 RTF 풀기 (MS-OXRTFCP, LZFu)
  2) 감싼 HTML 꺼내기 (MS-OXRTFEX)
를 한다. HTML 을 감싼 RTF 가 아니면(순수 RTF 메일) None 을 돌려준다 — 그때는 텍스트 본문을 쓴다.
"""

from __future__ import annotations

import struct

_LZFU = 0x75465A4C  # 'LZFu' 압축
_MELA = 0x414C454D  # 'MELA' 비압축
_PREBUF = (
    b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript "
    b"\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par "
    b"\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"
)
_DICT_SIZE = 4096


def decompress_rtf(data: bytes) -> bytes:
    if len(data) < 16:
        raise ValueError("압축 RTF 머리말이 짧음")
    comp_size, raw_size, kind, _crc = struct.unpack_from("<IIII", data, 0)
    body = data[16:4 + comp_size]
    if kind == _MELA:
        return body[:raw_size]
    if kind != _LZFU:
        raise ValueError("알 수 없는 압축 RTF 형식")
    dictionary = bytearray(_DICT_SIZE)
    dictionary[:len(_PREBUF)] = _PREBUF
    write = len(_PREBUF)
    out = bytearray()
    pos = 0
    while pos < len(body):
        control = body[pos]
        pos += 1
        for bit in range(8):
            if pos >= len(body):
                break
            if not control & (1 << bit):
                byte = body[pos]
                pos += 1
                out.append(byte)
                dictionary[write] = byte
                write = (write + 1) % _DICT_SIZE
                continue
            if pos + 1 >= len(body):
                return bytes(out)
            word = (body[pos] << 8) | body[pos + 1]
            pos += 2
            offset, length = word >> 4, (word & 0xF) + 2
            if offset == write:  # 끝 표시
                return bytes(out)
            for i in range(length):
                byte = dictionary[(offset + i) % _DICT_SIZE]
                out.append(byte)
                dictionary[write] = byte
                write = (write + 1) % _DICT_SIZE
    return bytes(out)


# 내용을 출력하지 않는 RTF 목적지 그룹
_SKIP_DESTINATIONS = {
    "fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "header", "footer",
    "listtable", "listoverridetable", "rsidtbl", "generator", "xmlnstbl", "mmathPr", "themedata",
}
_SYMBOLS = {"par": "\r\n", "line": "\r\n", "tab": "\t", "lquote": "‘", "rquote": "’",
            "ldblquote": "“", "rdblquote": "”", "bullet": "•", "endash": "–",
            "emdash": "—", "emspace": " ", "enspace": " ", "~": " "}


def html_from_rtf(rtf: bytes) -> str | None:
    """\\fromhtml1 로 감싼 HTML 이면 HTML 문자열, 아니면 None."""
    if b"\\fromhtml" not in rtf[:1024]:
        return None
    codepage = "cp1252"
    head = rtf[:1024]
    i = head.find(b"\\ansicpg")
    if i >= 0:
        digits = b""
        for ch in head[i + 8:i + 14]:
            if not chr(ch).isdigit():
                break
            digits += bytes([ch])
        if digits:
            codepage = f"cp{int(digits)}"
    try:
        "".encode(codepage)
    except LookupError:
        codepage = "cp1252"

    out: list[str] = []
    pending = bytearray()  # \'hh 바이트는 모아서 코드페이지로 한 번에 푼다 (한글 2바이트)

    def flush() -> None:
        if pending:
            out.append(pending.decode(codepage, errors="replace"))
            pending.clear()

    # 그룹 상태: [건너뛰기, htmlrtf(RTF 전용 내용), \uc 값]
    stack: list[list] = []
    state = [False, False, 1]
    skip_chars = 0  # \uN 뒤에 따라오는 대체 문자 수
    n = len(rtf)
    pos = 0
    star = False  # 직전에 \* 를 만났는가 (다음 제어어가 목적지 이름)

    def emitting() -> bool:
        return not state[0] and not state[1]

    while pos < n:
        ch = rtf[pos]
        if ch == 0x7B:  # {
            flush()
            stack.append(state[:])
            pos += 1
            star = False
            continue
        if ch == 0x7D:  # }
            flush()
            if stack:
                state = stack.pop()
            pos += 1
            star = False
            continue
        if ch in (0x0D, 0x0A):
            pos += 1
            continue
        if ch == 0x5C:  # \
            if pos + 1 >= n:
                break
            nxt = rtf[pos + 1]
            if nxt == 0x27:  # \'hh
                if skip_chars:
                    skip_chars -= 1
                elif emitting():
                    try:
                        pending.append(int(rtf[pos + 2:pos + 4], 16))
                    except ValueError:
                        pass
                pos += 4
                continue
            if not (0x61 <= nxt <= 0x7A or 0x41 <= nxt <= 0x5A):  # 제어 기호
                pos += 2
                if nxt == 0x2A:  # \*
                    star = True
                    continue
                if skip_chars:
                    skip_chars -= 1
                    continue
                if emitting():
                    if nxt in (0x5C, 0x7B, 0x7D):
                        flush()
                        out.append(chr(nxt))
                    elif chr(nxt) in _SYMBOLS:
                        flush()
                        out.append(_SYMBOLS[chr(nxt)])
                continue
            # 제어어: 글자 + 선택적 숫자 + 선택적 공백 하나
            end = pos + 1
            while end < n and (0x61 <= rtf[end] <= 0x7A or 0x41 <= rtf[end] <= 0x5A):
                end += 1
            word = rtf[pos + 1:end].decode("ascii")
            num_start = end
            if end < n and rtf[end] == 0x2D:
                end += 1
            while end < n and 0x30 <= rtf[end] <= 0x39:
                end += 1
            param = rtf[num_start:end].decode("ascii")
            if end < n and rtf[end] == 0x20:
                end += 1
            pos = end
            if word.startswith("htmltag"):
                state[0] = False  # 감싼 HTML 태그 — 출력한다
                state[1] = False
                star = False
                continue
            if star:
                state[0] = True  # 모르는 \* 목적지는 건너뛴다
                star = False
                continue
            if word in _SKIP_DESTINATIONS:
                state[0] = True
                continue
            if word == "htmlrtf":
                flush()
                state[1] = param != "0"
                continue
            if word == "uc":
                state[2] = int(param or 1)
                continue
            if word == "u":
                if emitting():
                    flush()
                    value = int(param or 0)
                    out.append(chr(value + 65536 if value < 0 else value))
                skip_chars = state[2]
                continue
            if skip_chars:
                skip_chars -= 1
                continue
            if emitting() and word in _SYMBOLS:
                flush()
                out.append(_SYMBOLS[word])
            continue
        # 일반 글자
        if skip_chars:
            skip_chars -= 1
        elif emitting():
            pending.append(ch)
        pos += 1
    flush()
    html = "".join(out).strip()
    return html or None


def html_from_compressed_rtf(data: bytes) -> str | None:
    return html_from_rtf(decompress_rtf(data))
