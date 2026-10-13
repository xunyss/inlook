"""VDI 자가진단.

COM 동작이 설계의 가정대로인지 확인한다. 결과는 사용자가 화면을 캡처하거나 붙여넣어
전달하므로 메일 내용·사람 이름·폴더명은 출력하지 않는다.
저장소는 store#N, 폴더는 folder#NN 익명 ID 로만 표시한다. 날짜·개수는 출력한다.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
from collections import Counter, defaultdict
from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path

from .errors import CODES
from .source import EXCLUDED_KINDS, FolderRef, MailHeader, MailSource, StoreRef

PASS, WARN, FAIL, INFO = "PASS", "WARN", "FAIL", "INFO"
RULE = "=" * 72

MESSAGE_ID_WARN_RATIO = 0.95
WINDOWS_MAX_PATH = 260
SINCE_CHECK_FOLDERS = 5
MEASURE_STEPS = ("open", "basic", "sender_smtp", "body", "html", "attachments")


class Redactor:
    """오류 메시지에 섞여 들어올 수 있는 폴더명·제목·이름을 가린다."""

    def __init__(self) -> None:
        self._secrets: set[str] = set()

    def add(self, *values: str) -> None:
        self._secrets.update(v for v in values if v and len(v.strip()) >= 2)

    def __call__(self, text: str) -> str:
        for secret in sorted(self._secrets, key=len, reverse=True):
            text = text.replace(secret, "***")
        return text


class Reporter:
    def __init__(self) -> None:
        self.lines: list[str] = []
        self.counts: Counter[str] = Counter()

    def line(self, text: str = "") -> None:
        print(text, flush=True)
        self.lines.append(text)

    def result(
        self,
        status: str,
        step: str,
        title: str,
        *,
        code: str | None = None,
        details: list[str] | None = None,
        elapsed: float | None = None,
    ) -> None:
        self.counts[status] += 1
        suffix = f"  ({elapsed:.1f}s)" if elapsed is not None else ""
        tag = f" [{code}]" if code else ""
        self.line(f"[{status}] {step}{tag} {title}{suffix}")
        for detail in details or []:
            self.line(f"       {detail}")
        if code and status in (WARN, FAIL):
            self.line(f"       조치: {CODES[code].action}")


def describe_exc(exc: BaseException) -> str:
    text = f"{type(exc).__name__}: {exc}"
    return text if len(text) <= 300 else text[:300] + "..."


def _fmt(n: int) -> str:
    return f"{n:,}"


def _day(value: datetime | None) -> str:
    return f"{value:%Y-%m-%d}" if value else "----------"


def _span(headers: list[MailHeader]) -> str:
    times = [h.received for h in headers if h.received]
    if not times:
        return "-"
    return f"{_day(min(times))} ~ {_day(max(times))}"


def _ratio(part: int, whole: int) -> str:
    return f"{part}/{whole} ({part / whole:.0%})" if whole else "0/0"


def _is_target(folder: FolderRef) -> bool:
    return folder.is_mail and folder.kind not in EXCLUDED_KINDS and folder.item_count > 0


def run_probe(
    make_source: Callable[[], MailSource],
    *,
    deep: int = 10,
    max_rows: int = 5000,
    archive_root: Path | None = None,
    log_dir: Path | None = None,
) -> int:
    """진단을 실행하고 종료 코드(FAIL 이 있으면 1)를 돌려준다."""
    redact = Redactor()
    out = Reporter()

    def err(exc: BaseException) -> str:
        return redact(describe_exc(exc))

    started = datetime.now()
    out.line(RULE)
    out.line(f" mailer probe v2   {started:%Y-%m-%d %H:%M:%S}   (deep={deep}, max_rows={max_rows})")
    out.line(RULE)

    # P1 연결
    t = time.perf_counter()
    try:
        source = make_source()
        info = source.describe()
    except Exception as exc:  # noqa: BLE001 — 어떤 실패든 화면에 남긴다
        out.result(FAIL, "P1", CODES["E101"].title, code="E101", details=[err(exc)])
        return _finish(out, log_dir, started)
    out.result(
        PASS, "P1", "Outlook 연결",
        details=[" / ".join(f"{k} {v}" for k, v in info.items())],
        elapsed=time.perf_counter() - t,
    )

    # P2 저장소와 폴더
    t = time.perf_counter()
    try:
        stores = list(source.stores())
    except Exception as exc:  # noqa: BLE001
        out.result(WARN, "P2", CODES["E112"].title, code="E112", details=[err(exc)])
        stores = []
    if not any(s.is_default for s in stores):
        stores = [StoreRef(index=1, is_default=True, kind="?", is_data_file=False), *stores]

    folders: list[FolderRef] = []
    for store in stores:
        try:
            folders += list(source.folders(store))
        except Exception as exc:  # noqa: BLE001
            code = "E102" if store.is_default else "E112"
            status = FAIL if store.is_default else WARN
            out.result(status, "P2", CODES[code].title, code=code, details=[f"store#{store.index}: {err(exc)}"])
            if store.is_default:
                return _finish(out, log_dir, started)
    for f in folders:
        redact.add(*f.path)
    ids = {(f.store_index, f.path): f"folder#{i:02d}" for i, f in enumerate(folders, 1)}

    def fid(folder: FolderRef) -> str:
        return ids[(folder.store_index, folder.path)]

    # 목록은 Table API 로 모든 메일 폴더를 읽는다 (빈 폴더 제외)
    headers: dict[str, list[MailHeader]] = {}
    truncated: set[str] = set()
    read_errors: list[str] = []
    mail_folders = [f for f in folders if f.is_mail and f.item_count > 0]
    t_rows = time.perf_counter()
    for i, folder in enumerate(mail_folders, 1):
        _progress(f"  목록 읽는 중 {fid(folder)} ({i}/{len(mail_folders)})")
        try:
            rows = list(source.headers(folder, limit=max_rows))
        except Exception as exc:  # noqa: BLE001
            read_errors.append(f"{fid(folder)}: {err(exc)}")
            continue
        for h in rows:
            redact.add(h.subject, h.sender_name)
        headers[fid(folder)] = rows
        if folder.item_count > len(rows):
            truncated.add(fid(folder))
    _progress("")
    rows_elapsed = time.perf_counter() - t_rows

    details = []
    for store in stores:
        mine = [f for f in folders if f.store_index == store.index]
        rows = [h for f in mine for h in headers.get(fid(f), [])]
        flags = ["기본" if store.is_default else "", "데이터파일" if store.is_data_file else ""]
        details.append(
            f"store#{store.index} {store.kind:12} {' '.join(x for x in flags if x):10} "
            f"폴더 {len(mine)} (메일 {sum(1 for f in mine if f.is_mail)}), "
            f"메일 {_fmt(sum(f.item_count for f in mine if f.is_mail and f.item_count > 0))}, "
            f"수신일 {_span(rows)}"
        )
    out.result(PASS, "P2", f"저장소 {len(stores)}개", details=details, elapsed=time.perf_counter() - t)

    inaccessible = [fid(f) for f in folders if f.item_count < 0]
    if inaccessible:
        out.result(WARN, "P2", CODES["E103"].title, code="E103", details=[", ".join(inaccessible)])

    # P3 폴더 목록 (빈 폴더 생략)
    details = [
        f"{fid(f)}  s{f.store_index}  {(f.kind or '-'):8} d{f.depth}  {'mail' if f.is_mail else 'other':5} "
        f"{_fmt(f.item_count):>7}  {_span(headers.get(fid(f), [])) if f.is_mail else ''}"
        f"{'  (수집대상)' if _is_target(f) else ''}"
        for f in folders
        if f.item_count != 0
    ]
    empty = sum(1 for f in folders if f.item_count == 0)
    details.append(f"(빈 폴더 {empty}개 생략)")
    out.result(INFO, "P3", "폴더별 항목 수와 수신일 범위", details=details)

    # P4 목록 읽기(Table API)와 키 분석
    targets = [f for f in folders if _is_target(f)]
    target_rows = [h for f in targets for h in headers.get(fid(f), [])]
    all_rows = [h for rows in headers.values() for h in rows]
    rate = len(all_rows) / rows_elapsed if rows_elapsed > 0 else 0.0
    details = [f"{_fmt(len(all_rows))}행 읽음, 초당 {rate:.0f}행" if rate else f"{_fmt(len(all_rows))}행 읽음"]
    if truncated:
        details.append(f"max_rows 에서 잘림: {', '.join(sorted(truncated))}")
    out.result(
        PASS if not read_errors else FAIL, "P4", "목록 읽기 (Table API)",
        details=details, elapsed=rows_elapsed,
    )
    if read_errors:
        out.result(FAIL, "P4", CODES["E104"].title, code="E104", details=read_errors)

    n = len(target_rows)
    with_id = sum(1 for h in target_rows if h.message_id)
    details = [
        f"수집 대상 {_fmt(n)}건 기준",
        "MessageClass: " + ", ".join(f"{c or '(empty)'}={k}" for c, k in Counter(h.message_class for h in target_rows).most_common()),
        f"Message-ID 있음 {_ratio(with_id, n)}",
        "  없는 항목 - 종류별: " + _missing_by(target_rows, lambda h: h.message_class),
    ]
    default_index = next(s.index for s in stores if s.is_default)
    group_of = {fid(f): _group(f, default_index) for f in targets}
    details.append(
        "  없는 항목 - 폴더별: "
        + _missing_by(((group_of[k], h) for k, rows in headers.items() if k in group_of for h in rows), None)
    )
    details.append(f"SearchKey 있음 {_ratio(sum(1 for h in target_rows if h.search_key), n)}")
    details.append("키별 중복 그룹 수 (같은 키를 가진 항목 묶음):")
    for name, key in _KEYS.items():
        values = [key(h) for h in target_rows]
        present = [v for v in values if v]
        dupes = sum(1 for c in Counter(present).values() if c > 1)
        details.append(f"  {name:22} 값 있음 {_ratio(len(present), n):18} 중복 그룹 {dupes}")
    out.result(INFO, "P4", "중복 판단 키 분석", details=details)
    if n and with_id / n < MESSAGE_ID_WARN_RATIO:
        out.result(WARN, "P4", CODES["E106"].title, code="E106", details=[f"{with_id / n:.0%} 만 있음"])
    stubs = sum(1 for h in all_rows if h.is_stub)
    if stubs:
        out.result(WARN, "P4", CODES["E107"].title, code="E107", details=[f"{_fmt(stubs)}개"])

    # P5 정렬과 증분 조회
    t = time.perf_counter()
    unordered = [k for k, rows in headers.items() if not _is_descending(rows)]
    mismatches: list[str] = []
    checked = 0
    for folder in targets:
        rows = headers.get(fid(folder), [])
        if len(rows) < 2 or fid(folder) in truncated:
            continue
        times = [h.received for h in rows if h.received]
        if not times:
            continue
        checked += 1
        since = times[len(times) // 2]
        expected = sum(1 for x in times if x >= since)
        try:
            got = sum(1 for h in source.headers(folder, since=since) if h.received)
        except Exception as exc:  # noqa: BLE001
            mismatches.append(f"{fid(folder)}: {err(exc)}")
            continue
        if got != expected:
            mismatches.append(f"{fid(folder)}: since 기대 {expected} / 실제 {got}")
        if checked >= SINCE_CHECK_FOLDERS:
            break
    elapsed = time.perf_counter() - t
    if unordered:
        out.result(FAIL, "P5", CODES["E105"].title, code="E105", details=[", ".join(unordered)], elapsed=elapsed)
    if mismatches:
        out.result(FAIL, "P5", CODES["E108"].title, code="E108", details=mismatches, elapsed=elapsed)
    if not unordered and not mismatches:
        out.result(PASS, "P5", f"정렬·증분 조회 ({checked}개 폴더)", elapsed=elapsed)

    # P6 상세 읽기 단계별 시간
    picks = _spread([headers.get(fid(f), []) for f in targets], deep)
    timings: dict[str, list[float]] = defaultdict(list)
    open_errors: list[str] = []
    for i, header in enumerate(picks, 1):
        _progress(f"  상세 읽기 {i}/{len(picks)}")
        try:
            for step, seconds in source.measure(header).items():
                timings[step].append(seconds)
        except Exception as exc:  # noqa: BLE001
            open_errors.append(err(exc))
    _progress("")
    if picks and timings:
        per_mail = sum(sum(v) / len(v) for v in timings.values())
        details = [f"{len(picks)}건 평균 (ms): " + ", ".join(
            f"{step} {sum(timings[step]) / len(timings[step]) * 1000:.0f}" for step in MEASURE_STEPS if timings[step]
        )]
        details.append(f"메일 1건 상세 읽기 약 {per_mail:.2f}초 → 수집 대상 {_fmt(n)}건이면 약 {n * per_mail / 60:.1f}분 (+ .msg 저장)")
        out.result(PASS if not open_errors else WARN, "P6", "상세 읽기 단계별 시간", details=details)
    else:
        out.result(INFO, "P6", "상세 읽기: 읽어볼 항목이 없음")
    if open_errors:
        out.result(WARN, "P6", CODES["E113"].title, code="E113", details=open_errors[:5])

    # P7 보존 경로 (P8 에서 실제 저장을 이 경로에 해보므로 먼저 만든다)
    _probe_archive_root(out, archive_root, err)

    # P8 .msg 저장
    candidate = next((h for h in picks if h.message_class == "IPM.Note"), next(iter(picks), None))
    if candidate is None:
        out.result(INFO, "P8", ".msg 저장: 저장해볼 항목이 없음")
    else:
        _probe_save(out, source, candidate, archive_root, err)

    return _finish(out, log_dir, started)


def _group(folder: FolderRef, default_index: int) -> str:
    if folder.store_index != default_index:
        return f"store#{folder.store_index}"
    if folder.kind:
        return folder.kind
    return "user-folder"


def _missing_by(rows: Iterable, key: Callable[[MailHeader], str] | None) -> str:
    """Message-ID 없는 항목 수를 그룹별로. rows 가 (group, header) 쌍이면 key 는 None."""
    total: Counter[str] = Counter()
    missing: Counter[str] = Counter()
    for row in rows:
        group, header = (row if key is None else (key(row), row))
        total[group] += 1
        if not header.message_id:
            missing[group] += 1
    parts = [f"{g or '(empty)'} {missing[g]}/{total[g]}" for g, _ in total.most_common() if missing[g]]
    return ", ".join(parts) or "없음"


_KEYS: dict[str, Callable[[MailHeader], str | None]] = {
    "Message-ID": lambda h: h.message_id,
    "SearchKey": lambda h: h.search_key,
    "보낸이+수신시각+제목": lambda h: f"{h.sender_name}|{h.received}|{h.subject}" if h.received else None,
    "MsgID 또는 SearchKey": lambda h: h.message_id or (f"sk:{h.search_key}" if h.search_key else None),
}


def _spread(groups: list[list[MailHeader]], count: int) -> list[MailHeader]:
    """여러 폴더에서 고르게, 일반 메일 위주로 골라낸다."""
    pools = [[h for h in rows if not h.is_stub] for rows in groups]
    picked: list[MailHeader] = []
    depth = 0
    while len(picked) < count and any(len(p) > depth for p in pools):
        for pool in pools:
            if len(pool) > depth and len(picked) < count:
                picked.append(pool[depth])
        depth += 1
    return picked


def _probe_save(
    out: Reporter,
    source: MailSource,
    header: MailHeader,
    archive_root: Path | None,
    err: Callable[[BaseException], str],
) -> None:
    t = time.perf_counter()
    base = archive_root if archive_root and archive_root.is_dir() else None
    work = Path(tempfile.mkdtemp(prefix="mailer-probe-", dir=base))
    try:
        record = source.open(header)
        sizes = []
        for name in ("mailer-probe.msg", "mailer-probe-한글.msg"):
            path = work / name
            source.save_msg(record, path)
            if not path.is_file() or path.stat().st_size == 0:
                raise OSError(f"{name}: 파일이 없거나 비어 있음")
            sizes.append(path.stat().st_size)
        del record
        where = "보존 경로" if base else "임시 폴더"
        out.result(
            PASS, "P8", ".msg 저장 (유니코드, 한글 파일명, 2회)",
            details=[f"{where}에 저장 후 삭제, 크기 {', '.join(_fmt(s) for s in sizes)} bytes"],
            elapsed=time.perf_counter() - t,
        )
    except Exception as exc:  # noqa: BLE001
        out.result(FAIL, "P8", CODES["E109"].title, code="E109", details=[err(exc)])
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _probe_archive_root(out: Reporter, archive_root: Path | None, err: Callable[[BaseException], str]) -> None:
    if archive_root is None:
        out.result(INFO, "P7", "보존 경로 미지정 — probe.bat <경로> 로 지정하면 쓰기 가능 여부를 확인합니다")
        return
    try:
        archive_root.mkdir(parents=True, exist_ok=True)
        test = archive_root / ".mailer-write-test"
        test.write_text("ok", encoding="utf-8")
        test.unlink()
    except Exception as exc:  # noqa: BLE001
        out.result(FAIL, "P7", CODES["E110"].title, code="E110", details=[err(exc)])
        return
    free = shutil.disk_usage(archive_root).free
    root_len = len(str(archive_root.resolve()))
    out.result(
        PASS, "P7", "보존 경로 쓰기",
        details=[f"경로 길이 {root_len}자, 남은 용량 {free / 1024**3:.1f} GB"],
    )
    if _windows_long_paths_enabled() is False:
        out.result(
            WARN, "P7", CODES["E111"].title, code="E111",
            details=[f"파일명에 쓸 수 있는 길이 약 {WINDOWS_MAX_PATH - root_len - 20}자 (파일명은 자동으로 줄입니다)"],
        )


def _windows_long_paths_enabled() -> bool | None:
    if sys.platform != "win32":
        return None
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem") as key:
            value, _ = winreg.QueryValueEx(key, "LongPathsEnabled")
            return bool(value)
    except OSError:
        return False


def _is_descending(rows: list[MailHeader]) -> bool:
    times = [h.received for h in rows if h.received]
    return all(a >= b for a, b in zip(times, times[1:]))


def _progress(text: str) -> None:
    if sys.stdout.isatty():
        print(f"\r{text:<60}", end="\r" if text else "", flush=True)


def _finish(out: Reporter, log_dir: Path | None, started: datetime) -> int:
    c = out.counts
    out.line("-" * 72)
    out.line(f" 결과: PASS {c[PASS]} / WARN {c[WARN]} / FAIL {c[FAIL]} / INFO {c[INFO]}")
    log = None
    if log_dir is not None:
        log_dir.mkdir(parents=True, exist_ok=True)
        log = log_dir / f"probe-{started:%Y%m%d-%H%M%S}.txt"
        out.line(f" 로그: {log.resolve()}")
    out.line(RULE)
    if log is not None:
        log.write_text("\n".join(out.lines) + "\n", encoding="utf-8")
    return 1 if c[FAIL] else 0
