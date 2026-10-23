"""이미 저장한 메일의 첨부 목록·보낸 사람 주소를 원본 .msg 에서 다시 읽어 바로잡는다.

- 첨부: 본문에 들어간 이미지(서명 로고 등)를 목록에서 뺀다 (📎 표시·검색·미리보기 첨부 목록)
- 보낸 사람: 회사 내부(Exchange) 사람의 '/O=...' 주소를 .msg 안의 SMTP 주소로 바꾼다.
  .msg 에 SMTP 주소가 없으면 그대로 둔다 (건수만 알려준다)

.msg 는 읽기만 한다. 메일 내용·이름·주소는 출력하지 않는다.
"""

from __future__ import annotations

import json
import time
from pathlib import Path

from .msgfile import read_details
from .store import DB_NAME, Store


def run_repair(root: Path, *, dry_run: bool = False) -> int:
    db_path = root / DB_NAME
    if not db_path.exists():
        print(f"DB 가 없습니다: {db_path}")
        return 2
    store = Store(root)
    try:
        rows = store.db.execute("SELECT id, msg_path, attachments, sender_email FROM mails").fetchall()
        print(f"메일 {len(rows):,}건의 .msg 를 확인합니다.")
        updates: list[tuple[str, str, int]] = []
        had_clip = clip_after = inline_removed = 0
        sender_fixed = sender_unresolved = 0
        missing = errors = 0
        t0 = time.perf_counter()
        for i, (mail_id, rel, attachments_json, sender_email) in enumerate(rows, 1):
            try:
                old_attachments = json.loads(attachments_json or "[]")
            except ValueError:
                old_attachments = []
            had_clip += bool(old_attachments)
            attachments, sender = old_attachments, sender_email or ""
            path = root / Path(rel)
            if not path.exists():
                missing += 1
            else:
                try:
                    details = read_details(path)
                except Exception:
                    errors += 1
                else:
                    attachments = details.attachments
                    inline_removed += max(0, len(old_attachments) - len(attachments))
                    if sender and "@" not in sender:
                        if details.sender_smtp:
                            sender = details.sender_smtp
                            sender_fixed += 1
                        else:
                            sender_unresolved += 1
            clip_after += bool(attachments)
            if attachments != old_attachments or sender != (sender_email or ""):
                updates.append((json.dumps(attachments, ensure_ascii=False), sender, mail_id))
            if i % 100 == 0 or i == len(rows):
                print(f"  {i:,}/{len(rows):,} 확인  ({time.perf_counter() - t0:.0f}s)", flush=True)

        print(f"첨부 표시(📎) 있는 메일 {had_clip:,}건 → {clip_after:,}건 (본문 이미지 {inline_removed:,}개를 목록에서 뺌)")
        print(f"보낸 사람 Exchange 주소 → SMTP 주소로 바꿈 {sender_fixed:,}건, .msg 에 SMTP 주소 없음 {sender_unresolved:,}건")
        if missing or errors:
            print(f".msg 파일 없음 {missing:,}건, 읽기 오류 {errors:,}건 → 그대로 둠")
        print(f"바뀌는 메일 {len(updates):,}건")
        if dry_run:
            print("dry-run: 바꾸지 않고 끝냅니다")
            return 0
        with store.db:
            store.db.executemany("UPDATE mails SET attachments = ?, sender_email = ? WHERE id = ?", updates)
        print("완료")
    finally:
        store.close()
    return 0
