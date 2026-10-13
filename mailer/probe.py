"""VDI 자가진단.

COM 동작이 설계의 가정대로인지 확인한다. 결과는 사용자가 화면을 캡처해 전달하므로
메일 내용·사람 이름·폴더명은 출력하지 않는다. 폴더는 folder#NN 익명 ID 로만 표시한다.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
from collections import Counter
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from .errors import CODES
from .source import EXCLUDED_KINDS, FolderRef, MailRecord, MailSource

PASS, WARN, FAIL, INFO = "PASS", "WARN", "FAIL", "INFO"
RULE = "=" * 64

MESSAGE_ID_WARN_RATIO = 0.95
WINDOWS_MAX_PATH = 260


class Redactor:
    """오류 메시지에 섞여 들어올 수 있는 폴더명·제목을 가린다."""

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
        head = f"[{status}] {step} {title}{suffix}"
        if code:
            head = f"[{status}] {step} [{code}] {title}{suffix}"
        self.line(head)
        for detail in details or []:
            self.line(f"       {detail}")
        if code and status in (WARN, FAIL):
            self.line(f"       조치: {CODES[code].action}")


def describe_exc(exc: BaseException) -> str:
    text = f"{type(exc).__name__}: {exc}"
    return text if len(text) <= 300 else text[:300] + "..."


def _fmt(n: int) -> str:
    return f"{n:,}"


def run_probe(
    make_source: Callable[[], MailSource],
    *,
    sample: int = 100,
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
    out.line(f" mailer probe   {started:%Y-%m-%d %H:%M:%S}   (sample={sample})")
    out.line(RULE)

    # P1 연결
    t = time.perf_counter()
    try:
        source = make_source()
    except Exception as exc:  # noqa: BLE001 — 어떤 실패든 화면에 남긴다
        out.result(FAIL, "P1", CODES["E101"].title, code="E101", details=[err(exc)])
        return _finish(out, log_dir, started)
    info = " / ".join(f"{k} {v}" for k, v in source.describe().items())
    out.result(PASS, "P1", "Outlook 연결", details=[info], elapsed=time.perf_counter() - t)

    # P2 폴더 순회
    t = time.perf_counter()
    try:
        folders = list(source.folders())
    except Exception as exc:  # noqa: BLE001
        out.result(FAIL, "P2", CODES["E102"].title, code="E102", details=[err(exc)])
        return _finish(out, log_dir, started)
    ids = {f.path: f"folder#{i:02d}" for i, f in enumerate(folders, 1)}
    for f in folders:
        redact.add(*f.path)
    mail_folders = [f for f in folders if f.is_mail]
    targets = [f for f in mail_folders if f.kind not in EXCLUDED_KINDS and f.item_count > 0]
    inaccessible = [f for f in folders if f.item_count < 0]
    total_items = sum(f.item_count for f in mail_folders if f.item_count > 0)
    max_name = max((len("\\".join(f.path)) for f in folders), default=0)
    details = [
        f"폴더 {len(folders)}개 (메일 {len(mail_folders)}, 수집 대상 {len(targets)}), "
        f"메일 항목 {_fmt(total_items)}개, 최대 깊이 {max((f.depth for f in folders), default=0)}, "
        f"가장 긴 폴더 경로 {max_name}자",
    ]
    details += [
        f"{ids[f.path]}  {(f.kind or '-'):8} d{f.depth}  {'mail' if f.is_mail else 'other':5} {_fmt(f.item_count):>8}"
        for f in folders
    ]
    out.result(PASS, "P2", "폴더 순회", details=details, elapsed=time.perf_counter() - t)
    if inaccessible:
        out.result(
            WARN, "P2", CODES["E103"].title, code="E103",
            details=[", ".join(ids[f.path] for f in inaccessible)],
        )

    # P3 항목 샘플 읽기
    t = time.perf_counter()
    samples: dict[tuple[str, ...], list[MailRecord]] = {}
    read_errors: list[str] = []
    for i, folder in enumerate(targets, 1):
        _progress(f"  읽는 중 {ids[folder.path]} ({i}/{len(targets)})")
        try:
            records = list(source.items(folder, limit=sample))
        except Exception as exc:  # noqa: BLE001
            read_errors.append(f"{ids[folder.path]}: {err(exc)}")
            continue
        for r in records:
            redact.add(r.subject, r.sender_name, r.sender_email)
        samples[folder.path] = records
    _progress("")
    elapsed = time.perf_counter() - t
    all_records = [r for rs in samples.values() for r in rs]
    n = len(all_records)
    classes = Counter(r.message_class for r in all_records)
    with_id = sum(1 for r in all_records if r.message_id)
    stubs = sum(1 for r in all_records if r.is_stub)
    no_time = sum(1 for r in all_records if r.received is None)
    rate = n / elapsed if elapsed > 0 else 0.0
    details = [
        f"항목 {_fmt(n)}개 읽음, 초당 {rate:.0f}개 → 전체 {_fmt(total_items)}개 예상 {total_items / rate / 60:.1f}분"
        if rate else f"항목 {_fmt(n)}개 읽음",
        f"Message-ID 있음 {_fmt(with_id)}/{_fmt(n)}, 수신시각 없음 {_fmt(no_time)}, "
        f"첨부 있음 {_fmt(sum(1 for r in all_records if r.attachments))}, "
        f"가장 긴 제목 {max((len(r.subject) for r in all_records), default=0)}자",
        "MessageClass: " + ", ".join(f"{c or '(empty)'}={k}" for c, k in classes.most_common()),
    ]
    out.result(PASS if not read_errors else FAIL, "P3", "항목 읽기", details=details, elapsed=elapsed)
    if read_errors:
        out.result(FAIL, "P3", CODES["E104"].title, code="E104", details=read_errors)
    if n and with_id / n < MESSAGE_ID_WARN_RATIO:
        out.result(WARN, "P3", CODES["E106"].title, code="E106", details=[f"{with_id / n:.0%} 만 있음"])
    if stubs:
        out.result(WARN, "P3", CODES["E107"].title, code="E107", details=[f"{_fmt(stubs)}개"])

    unordered = [ids[p] for p, rs in samples.items() if not _is_descending(rs)]
    if unordered:
        out.result(FAIL, "P3", CODES["E105"].title, code="E105", details=[", ".join(unordered)])
    else:
        out.result(PASS, "P3", "수신 시각 내림차순 정렬")

    # P4 증분 조회: 샘플 안에 다 들어오는 작은 폴더로 전체 개수와 since 경계를 확인
    t = time.perf_counter()
    checked, mismatches = 0, []
    for folder in targets:
        records = samples.get(folder.path)
        if not records or folder.item_count > sample or len(records) < 2:
            continue
        checked += 1
        problems = []
        if len(records) != folder.item_count:
            problems.append(f"전체 {folder.item_count} / 읽음 {len(records)}")
        times = [r.received for r in records if r.received]
        if times:
            since = times[len(times) // 2]
            expected = sum(1 for x in times if x >= since)
            got = sum(1 for r in source.items(folder, since=since) if r.received)
            if got != expected:
                problems.append(f"since 기대 {expected} / 실제 {got}")
        if problems:
            mismatches.append(f"{ids[folder.path]}: " + ", ".join(problems))
        if checked >= 5:
            break
    if mismatches:
        out.result(FAIL, "P4", CODES["E108"].title, code="E108", details=mismatches,
                   elapsed=time.perf_counter() - t)
    elif checked:
        out.result(PASS, "P4", f"증분 조회 ({checked}개 폴더 확인)", elapsed=time.perf_counter() - t)
    else:
        out.result(INFO, "P4", "증분 조회: 확인할 만큼 작은 폴더가 없음 (sample 값을 늘려 다시 실행)")

    # P5 보존 경로 (P6 에서 실제 저장을 이 경로에 해보므로 먼저 만든다)
    _probe_archive_root(out, archive_root)

    # P6 .msg 저장
    candidate = next(
        (r for r in all_records if r.message_class == "IPM.Note" and not r.is_stub),
        next(iter(all_records), None),
    )
    if candidate is None:
        out.result(INFO, "P6", ".msg 저장: 저장해볼 항목이 없음")
    else:
        _probe_save(out, source, candidate, archive_root, err)

    return _finish(out, log_dir, started)


def _probe_save(
    out: Reporter,
    source: MailSource,
    item: MailRecord,
    archive_root: Path | None,
    err: Callable[[BaseException], str],
) -> None:
    t = time.perf_counter()
    base = archive_root if archive_root and archive_root.is_dir() else None
    work = Path(tempfile.mkdtemp(prefix="mailer-probe-", dir=base))
    try:
        sizes = []
        for name in ("mailer-probe.msg", "mailer-probe-한글.msg"):
            path = work / name
            source.save_msg(item, path)
            if not path.is_file() or path.stat().st_size == 0:
                raise OSError(f"{name}: 파일이 없거나 비어 있음")
            sizes.append(path.stat().st_size)
        where = "보존 경로" if base else "임시 폴더"
        out.result(
            PASS, "P6", ".msg 저장 (유니코드, 한글 파일명)",
            details=[f"{where}에 저장 후 삭제, 크기 {', '.join(_fmt(s) for s in sizes)} bytes"],
            elapsed=time.perf_counter() - t,
        )
    except Exception as exc:  # noqa: BLE001
        out.result(FAIL, "P6", CODES["E109"].title, code="E109", details=[err(exc)])
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _probe_archive_root(out: Reporter, archive_root: Path | None) -> None:
    if archive_root is None:
        out.result(INFO, "P5", "보존 경로 미지정 — probe.bat <경로> 로 지정하면 쓰기 가능 여부를 확인합니다")
        return
    try:
        archive_root.mkdir(parents=True, exist_ok=True)
        test = archive_root / ".mailer-write-test"
        test.write_text("ok", encoding="utf-8")
        test.unlink()
    except Exception as exc:  # noqa: BLE001
        out.result(FAIL, "P5", CODES["E110"].title, code="E110", details=[describe_exc(exc)])
        return
    free = shutil.disk_usage(archive_root).free
    root_len = len(str(archive_root.resolve()))
    out.result(
        PASS, "P5", "보존 경로 쓰기",
        details=[f"경로 길이 {root_len}자, 남은 용량 {free / 1024**3:.1f} GB"],
    )
    long_paths = _windows_long_paths_enabled()
    if long_paths is False:
        out.result(
            WARN, "P5", CODES["E111"].title, code="E111",
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


def _is_descending(records: list[MailRecord]) -> bool:
    times = [r.received for r in records if r.received]
    return all(a >= b for a, b in zip(times, times[1:]))


def _progress(text: str) -> None:
    if sys.stdout.isatty():
        print(f"\r{text:<60}", end="\r" if text else "", flush=True)


def _finish(out: Reporter, log_dir: Path | None, started: datetime) -> int:
    c = out.counts
    out.line("-" * 64)
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
