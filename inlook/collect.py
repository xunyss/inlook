"""수집(동기화): Outlook 기본 사서함 → .msg + DB.

1. 대상 폴더마다 Table 을 최신순으로 읽다가 기준 시각(마지막 정상 동기화 - 여유)보다
   오래된 메일이 나오면 멈춘다. 첫 동기화는 기준 시각 없이 전부 읽는다.
2. DB 에 없는 메일만 **오래된 것부터** 열어서 .msg 저장 + DB 기록.
   보존 기간이 지나 곧 사라질 메일을 먼저 확보하기 위해서다.
3. 메일 한 건마다 바로 커밋한다. 중간에 멈춰도 저장된 건 남고, 다시 돌리면 나머지를 이어서 한다.
   실패나 중단이 있었던 동기화는 기준 시각으로 쓰지 않는다 (다음에 같은 구간을 다시 본다).

출력은 화면 캡처로 전달되므로 제목·이름·폴더명을 찍지 않는다. 폴더는 folder#NN 으로만.
"""

from __future__ import annotations

import gc
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

from .errors import CODES
from .probe import FAIL, INFO, PASS, RULE, WARN, Redactor, Reporter, describe_exc
from .source import EXCLUDED_KINDS, FolderRef, MailHeader, MailSource
from .msgfile import read_details
from .store import Store, mail_key, msg_relpath

DEFAULT_MARGIN_DAYS = 3  # 사용자 결정 2026-10-14. 시계 차이 + 지운 편지함에서 최근 메일 복원 대비
MAX_CONSECUTIVE_FAILURES = 5
WINDOWS_MAX_PATH = 260
# archive_root 뒤에 붙는 경로 길이: \msg\YYYY\MM\YYYYMMDD-HHMMSS_<12>.msg
_REL_PATH_LEN = len("\\msg\\2026\\10\\20261014-093215_3f9a1c2b7e04.msg")
GC_EVERY = 20


@dataclass
class Candidate:
    key: str
    key_kind: str
    header: MailHeader
    folder_id: str
    labels: list[str] = field(default_factory=list)


def is_target(folder: FolderRef) -> bool:
    return folder.is_mail and folder.kind not in EXCLUDED_KINDS and folder.item_count > 0


# 라벨은 '/' 로 계층을 나눈다. Outlook 폴더 이름 안의 '/' 는 모양이 같은 다른 문자로 바꾼다
# (예: 'AI/IT 센터' 폴더가 'AI' 아래 'IT 센터' 로 쪼개지지 않게).
FOLDER_SLASH = "∕"  # ∕ DIVISION SLASH


def label_of(folder: FolderRef) -> str:
    return "/".join(part.replace("/", FOLDER_SLASH) for part in folder.path)


def _legacy_label_fixes(folders: list[FolderRef]) -> dict[str, str]:
    """예전 방식('/' 그대로)으로 붙은 라벨 → 올바른 라벨. 실제 다른 폴더 경로와 겹치는 것은 건드리지 않는다."""
    real = {label_of(f) for f in folders}
    fixes = {}
    for folder in folders:
        old, new = "/".join(folder.path), label_of(folder)
        if old != new and old not in real:
            fixes[old] = new
    return fixes


def _eta(seconds: float) -> str:
    seconds = int(seconds)
    if seconds >= 3600:
        return f"{seconds // 3600}h{seconds % 3600 // 60:02d}m"
    return f"{seconds // 60}m{seconds % 60:02d}s"


def run_collect(
    make_source: Callable[[], MailSource],
    archive_root: Path,
    *,
    full: bool = False,
    dry_run: bool = False,
    margin_days: int = DEFAULT_MARGIN_DAYS,
    log_dir: Path | None = None,
    echo: Callable[[str], None] | None = None,
    progress: Callable[[int, int], None] | None = None,
    cancel: threading.Event | None = None,
) -> int:
    """종료 코드: 0 정상, 1 일부 실패, 2 시작 불가, 130 중단 (Ctrl+C 또는 cancel).

    echo: 화면 출력 대신 받을 함수 (앱의 로그 창). progress: (저장한 수, 전체) 통지.
    cancel: 앱의 [중단] 버튼. 진행 중인 메일 한 건을 마친 뒤 멈춘다.
    """
    started = datetime.now()
    log = None
    if log_dir is not None:
        log_dir.mkdir(parents=True, exist_ok=True)
        log = (log_dir / f"collect-{started:%Y%m%d-%H%M%S}.txt").open("w", encoding="utf-8")
    try:
        out = Reporter(log, echo)
        collector = _Collector(out, make_source, archive_root, full, dry_run, margin_days, started)
        collector.progress = progress
        collector.cancel = cancel
        return collector.run()
    finally:
        gc.collect()
        if log:
            log.close()


class _Collector:
    def __init__(self, out, make_source, archive_root, full, dry_run, margin_days, started) -> None:
        self.out: Reporter = out
        self.make_source = make_source
        self.root: Path = archive_root
        self.full = full
        self.dry_run = dry_run
        self.margin = timedelta(days=margin_days)
        self.started: datetime = started
        self.redact = Redactor()
        self.store: Store | None = None
        self.sync_id: int | None = None
        self.listed = self.new = self.saved = self.failed = 0
        self.folder_errors = 0
        self.progress: Callable[[int, int], None] | None = None
        self.cancel: threading.Event | None = None

    def check_cancel(self) -> None:
        if self.cancel is not None and self.cancel.is_set():
            raise KeyboardInterrupt

    def err(self, exc: BaseException) -> str:
        return self.redact(describe_exc(exc))

    def fail(self, step: str, code: str, exc: BaseException | None = None) -> int:
        details = [self.err(exc)] if exc else []
        self.out.result(FAIL, step, CODES[code].title, code=code, details=details)
        return 2

    def run(self) -> int:
        out = self.out
        out.line(RULE)
        out.line(f" inlook collect   {self.started:%Y-%m-%d %H:%M:%S}" + ("   (dry-run)" if self.dry_run else ""))
        out.line(RULE)
        try:
            return self._run()
        except KeyboardInterrupt:
            out.line("")
            out.line("[STOP] 중단 — 지금까지 저장한 메일은 DB 에 남아 있습니다. 다시 실행하면 이어서 합니다.")
            self._end("stopped")
            self._summary()
            return 130
        finally:
            if self.store:
                self.store.close()

    def _run(self) -> int:
        out = self.out

        # C1 보존 경로
        root = self.root.resolve()
        if len(str(root)) + _REL_PATH_LEN >= WINDOWS_MAX_PATH:
            return self.fail("C1", "E111")
        try:
            self.store = Store(root)
            probe_file = root / ".write-test"
            probe_file.write_bytes(b"ok")
            probe_file.unlink()
        except Exception as exc:
            return self.fail("C1", "E110", exc)
        last = self.store.last_ok_sync()
        cutoff = None if (self.full or last is None) else last.started - self.margin
        out.result(PASS, "C1", "보존 경로", details=[
            f"DB 저장된 메일: {self.store.count():,}건",
            f"마지막 정상 동기화: {last.started:%Y-%m-%d %H:%M}" if last else "마지막 정상 동기화: 없음 (첫 동기화)",
            f"조회 기준: {cutoff:%Y-%m-%d %H:%M} 이후 수신" if cutoff else "조회 기준: 전체",
        ])

        # C2 Outlook 연결
        out.begin("Outlook 연결")
        try:
            source = self.make_source()
            mode = source.describe().get("mode", "?")
        except Exception as exc:
            return self.fail("C2", "E101", exc)
        out.result(PASS, "C2", f"Outlook 연결 (mode={mode})")

        # C3 폴더 목록
        out.begin("폴더 목록")
        try:
            folders = list(source.folders())
        except Exception as exc:
            return self.fail("C3", "E102", exc)
        targets: list[tuple[str, FolderRef]] = []
        for index, folder in enumerate(folders, 1):
            self.redact.add(*folder.path)
            if is_target(folder):
                targets.append((f"folder#{index:02d}", folder))
        out.result(PASS, "C3", f"수집 대상 폴더 {len(targets)}개 / 전체 {len(folders)}개")
        if not self.dry_run:
            fixed = self.store.fix_labels(_legacy_label_fixes(folders))
            if fixed:
                out.result(INFO, "C3", f"이름에 '/' 가 든 폴더의 라벨 {fixed}개를 바로잡음")

        # C4 목록 조회 (Table, 기준 시각까지만)
        candidates: dict[str, Candidate] = {}
        t0 = time.perf_counter()
        for folder_id, folder in targets:
            self.check_cancel()
            out.begin(f"{folder_id} 목록")
            rows = 0
            try:
                for header in source.headers(folder, since=cutoff):
                    rows += 1
                    key, kind = mail_key(header)
                    cand = candidates.get(key)
                    if cand is None:
                        cand = candidates[key] = Candidate(key, kind, header, folder_id)
                    label = label_of(folder)
                    if label not in cand.labels:
                        cand.labels.append(label)
            except Exception as exc:
                self.folder_errors += 1
                out.result(WARN, "C4", f"{folder_id} 목록 읽기 실패 ({rows}행까지 읽음)",
                           code="E104", details=[self.err(exc)])
            self.listed += rows
        pending = [c for c in candidates.values() if not self.store.known(c.key)]
        pending.sort(key=lambda c: c.header.received or datetime.min)
        self.new = len(pending)
        details = [f"목록 {self.listed:,}행, 서로 다른 메일 {len(candidates):,}건, 새 메일 {self.new:,}건"]
        if pending:
            details.append(
                f"새 메일 수신 범위: {pending[0].header.received or '?'} ~ {pending[-1].header.received or '?'}"
            )
        out.result(PASS if not self.folder_errors else WARN, "C4", "목록 조회", details=details,
                   elapsed=time.perf_counter() - t0)

        if self.dry_run:
            out.result(INFO, "C5", "dry-run: 저장하지 않고 끝냅니다")
            return 0

        # C5 저장 (오래된 것부터)
        self.sync_id = self.store.begin_sync(self.started, cutoff)
        status = self._save_all(source, pending, root)
        self._end(status)
        self._summary()
        return 0 if status == "ok" else 1

    def _save_all(self, source: MailSource, pending: list[Candidate], root: Path) -> str:
        out = self.out
        total = len(pending)
        width = len(str(total))
        consecutive = 0
        t0 = time.perf_counter()
        if self.progress:
            self.progress(0, total)
        for i, cand in enumerate(pending, 1):
            self.check_cancel()
            h = cand.header
            self.redact.add(h.subject, h.sender_name)
            when = h.received or datetime.now()
            rel = msg_relpath(cand.key, when)
            out.part(f"  [{i:>{width}}/{total}] {when:%Y-%m-%d %H:%M} {cand.folder_id} ")
            t = time.perf_counter()
            record = None
            try:
                try:
                    record = source.open(h)
                except Exception as exc:
                    raise _Step("E113", exc) from exc
                try:
                    path = root / rel
                    path.parent.mkdir(parents=True, exist_ok=True)
                    if path.exists():
                        path.unlink()  # 지난번에 저장만 되고 DB 기록 전에 멈춘 경우
                    source.save_msg(record, path)
                except Exception as exc:
                    raise _Step("E109", exc) from exc
                # HTML 은 DB 에 넣지 않는다 (메일당 수백 KB). 미리보기는 .msg 에서 꺼낸다.
                # 저장한 .msg 에서 HTML 을 못 꺼내는 경우에만 COM 이 준 HTML 을 DB 에 남긴다.
                # 첨부 목록은 본문 이미지를 빼고, 보낸 사람은 Exchange 주소 대신 SMTP 주소로.
                try:
                    details = read_details(path)
                except Exception:
                    details = None
                if details is not None:
                    if details.html:
                        record.html_body = ""
                    record.attachments = details.attachments
                    if details.sender_smtp and "@" not in record.sender_email:
                        record.sender_email = details.sender_smtp
                try:
                    self.store.add(cand.key, cand.key_kind, record, rel, cand.labels, self.sync_id)
                except Exception as exc:
                    raise _Step("E114", exc) from exc
            except _Step as step:
                self.failed += 1
                consecutive += 1
                out.line(f"FAIL [{step.code}] {CODES[step.code].title}")
                out.line(f"       {self.err(step.exc)}")
                if consecutive >= MAX_CONSECUTIVE_FAILURES:
                    out.result(FAIL, "C5", f"연속 {consecutive}건 실패 — 수집을 멈춥니다",
                               code=step.code)
                    return "partial"
                continue
            finally:
                if record is not None:
                    record.handle = None  # 온라인 모드 동시 열기 제한
                del record
                if i % GC_EVERY == 0:
                    gc.collect()
            consecutive = 0
            self.saved += 1
            spent = time.perf_counter() - t0
            remain = spent / i * (total - i)
            out.line(f"saved {time.perf_counter() - t:.1f}s   남은 시간 약 {_eta(remain)}")
            if self.progress:
                self.progress(i, total)
        return "ok" if not self.failed and not self.folder_errors else "partial"

    def _end(self, status: str) -> None:
        if self.store and self.sync_id is not None:
            self.store.end_sync(self.sync_id, status, listed=self.listed, new=self.new,
                                saved=self.saved, failed=self.failed)
            self.sync_id = None

    def _summary(self) -> None:
        out = self.out
        out.line(RULE)
        elapsed = datetime.now() - self.started
        out.line(f" 새 메일 {self.new:,}건 중 저장 {self.saved:,}건, 실패 {self.failed:,}건"
                 f"   (소요 {_eta(elapsed.total_seconds())})")
        if self.store:
            out.line(f" DB 전체 메일: {self.store.count():,}건")
        if self.failed or self.folder_errors:
            out.line(" 실패한 메일은 다음 실행 때 다시 시도합니다.")
        out.line(RULE)


class _Step(Exception):
    def __init__(self, code: str, exc: BaseException) -> None:
        super().__init__(code)
        self.code = code
        self.exc = exc
