"""가져오기: 예전에 Outlook 에서 '다른 이름으로 저장' 해 둔 .msg 폴더 → inlook.

Outlook 을 쓰지 않는다. .msg 파일을 직접 읽어(msgfile) 보존 경로로 복사하고 DB 에 넣는다.

라벨 (사용자 결정 2026-10-15):
  가져올 폴더 바로 아래 파일      → '받은 편지함'
  하위 폴더 'X' 안의 파일          → '받은 편지함/X'   (하위의 하위는 '받은 편지함/X/Y')
Outlook 에서 수집한 '받은 편지함/X' 라벨과 이름이 같으면 한 라벨로 합쳐진다.

이미 있는 메일(같은 Message-ID 등)과 inlook 에서 삭제한 메일은 건너뛴다. 백업 안에 같은 메일이
여러 폴더에 있으면 메일 한 건에 라벨 여러 개.

동기화 기록(syncs)은 남기지 않는다 — 가져오기가 증분 동기화의 기준 시각을 바꾸면 안 된다.
화면·로그에는 파일 이름(제목·고객사명)을 찍지 않는다. file#N 으로만.
"""

from __future__ import annotations

import gc
import hashlib
import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .errors import CODES
from .probe import FAIL, INFO, PASS, RULE, WARN, Redactor, Reporter, describe_exc
from .source import MailHeader, MailRecord
from .store import Store, mail_key, msg_relpath

DEFAULT_PARENT_LABEL = "받은 편지함"
FOLDER_SLASH = "∕"  # collect.FOLDER_SLASH 와 같은 규칙


@dataclass
class _Item:
    number: int
    path: Path
    key: str
    key_kind: str
    received: datetime | None
    labels: list[str] = field(default_factory=list)


def label_for(source: Path, file: Path, parent_label: str) -> str:
    parts = [p.replace("/", FOLDER_SLASH) for p in file.parent.relative_to(source).parts]
    return "/".join([parent_label, *parts]) if parent_label else ("/".join(parts) or "가져온 메일")


def _file_key(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return "file:" + h.hexdigest()


def run_import(
    source: Path,
    archive_root: Path,
    *,
    parent_label: str = DEFAULT_PARENT_LABEL,
    dry_run: bool = False,
    log_dir: Path | None = None,
) -> int:
    """종료 코드: 0 정상, 1 일부 실패, 2 시작 불가, 130 중단."""
    from .msgfile import read_message

    started = datetime.now()
    log = None
    if log_dir is not None:
        log_dir.mkdir(parents=True, exist_ok=True)
        log = (log_dir / f"import-{started:%Y%m%d-%H%M%S}.txt").open("w", encoding="utf-8")
    out = Reporter(log)
    redact = Redactor()
    store: Store | None = None
    counts = {"imported": 0, "known": 0, "failed": 0}
    try:
        out.line(RULE)
        out.line(f" inlook import   {started:%Y-%m-%d %H:%M:%S}" + ("   (dry-run)" if dry_run else ""))
        out.line(RULE)

        # I1 경로 확인
        source = source.resolve()
        root = archive_root.resolve()
        if not source.is_dir():
            out.result(FAIL, "I1", "가져올 폴더가 없음", details=["경로를 확인하세요."])
            return 2
        if source == root or root in source.parents:
            out.result(FAIL, "I1", "가져올 폴더가 보존 경로 안에 있음", details=["inlook 이 저장한 폴더는 가져올 수 없습니다."])
            return 2
        try:
            store = Store(root)
        except Exception as exc:
            out.result(FAIL, "I1", CODES["E110"].title, code="E110", details=[describe_exc(exc)])
            return 2
        files = sorted(p for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".msg")
        folders = len({f.parent for f in files})
        out.result(PASS, "I1", f".msg {len(files):,}개 (폴더 {folders}곳)", details=[f"DB 저장된 메일: {store.count():,}건"])

        # I2 읽기: 키·날짜만 먼저 모아 같은 메일을 묶는다
        items: dict[str, _Item] = {}
        t0 = time.perf_counter()
        for number, path in enumerate(files, 1):
            redact.add(path.stem, *path.relative_to(source).parts)
            try:
                msg = read_message(path)
            except Exception as exc:
                counts["failed"] += 1
                out.result(WARN, "I2", f"file#{number} 읽기 실패", code="E115", details=[redact(describe_exc(exc))])
                continue
            header = MailHeader("", "", msg.message_class, "", "", msg.received, 0, msg.message_id, msg.search_key)
            key, kind = mail_key(header)
            if kind == "entry_id":  # Message-ID·SearchKey 둘 다 없음 → 파일 내용으로
                key, kind = _file_key(path), "file_hash"
            item = items.setdefault(key, _Item(number, path, key, kind, msg.received))
            label = label_for(source, path, parent_label)
            if label not in item.labels:
                item.labels.append(label)
        pending = [i for i in items.values() if not store.known(i.key)]
        counts["known"] = len(items) - len(pending)
        pending.sort(key=lambda i: i.received or datetime.min)
        out.result(PASS, "I2", "읽기", elapsed=time.perf_counter() - t0, details=[
            f"서로 다른 메일 {len(items):,}건 = 새 메일 {len(pending):,}건 + 이미 있음 {counts['known']:,}건",
        ])
        if dry_run:
            out.result(INFO, "I3", "dry-run: 저장하지 않고 끝냅니다")
            return 0

        # I3 복사 + DB
        total = len(pending)
        width = len(str(total))
        for i, item in enumerate(pending, 1):
            when = item.received or datetime.fromtimestamp(item.path.stat().st_mtime).replace(microsecond=0)
            out.part(f"  [{i:>{width}}/{total}] {when:%Y-%m-%d %H:%M} file#{item.number} ")
            try:
                msg = read_message(item.path)
                rel = msg_relpath(item.key, when)
                target = root / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(item.path, target)
                header = MailHeader(
                    entry_id="", store_id="", message_class=msg.message_class, subject=msg.subject,
                    sender_name=msg.sender_name, received=when, size=msg.size,
                    message_id=msg.message_id, search_key=msg.search_key,
                )
                record = MailRecord(
                    header=header, sender_email=msg.sender_email, to=msg.to, cc=msg.cc, sent=msg.sent,
                    body=msg.body, html_body=msg.html_body, attachments=msg.attachments,
                )
                store.add(item.key, item.key_kind, record, rel, item.labels, None)
            except Exception as exc:
                counts["failed"] += 1
                out.line(f"FAIL [E115] {CODES['E115'].title}")
                out.line(f"       {redact(describe_exc(exc))}")
                continue
            counts["imported"] += 1
            out.line("ok")
            if i % 50 == 0:
                gc.collect()
        return 0 if not counts["failed"] else 1
    except KeyboardInterrupt:
        out.line("")
        out.line("[STOP] 중단 — 지금까지 가져온 메일은 남아 있습니다. 다시 실행하면 나머지를 가져옵니다.")
        return 130
    finally:
        out.line(RULE)
        out.line(f" 가져옴 {counts['imported']:,}건, 이미 있음 {counts['known']:,}건, 실패 {counts['failed']:,}건")
        if store is not None:
            out.line(f" DB 전체 메일: {store.count():,}건")
            store.close()
        out.line(RULE)
        if log:
            log.close()
