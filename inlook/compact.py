"""DB 줄이기: HTML 본문을 DB 에서 비우고(원본 .msg 에서 꺼낼 수 있는 메일만) VACUUM 한다.

미리보기는 .msg 에서 HTML 을 꺼내므로 DB 의 html_body 는 필요 없다 (2026-10-15 결정, 용량의 약 90%).
.msg 에서 HTML 을 못 꺼내는 메일은 DB 의 HTML 을 그대로 둔다. 메일 내용은 출력하지 않는다.
"""

from __future__ import annotations

import time
from pathlib import Path

from .msgfile import read_html
from .store import DB_NAME, Store


def run_compact(root: Path, *, dry_run: bool = False) -> int:
    db_path = root / DB_NAME
    if not db_path.exists():
        print(f"DB 가 없습니다: {db_path}")
        return 2
    before = db_path.stat().st_size
    store = Store(root)
    try:
        rows = store.db.execute(
            "SELECT id, msg_path FROM mails WHERE html_body IS NOT NULL AND html_body != ''"
        ).fetchall()
        print(f"HTML 이 DB 에 있는 메일 {len(rows):,}건을 확인합니다. (DB {before / 1e6:.1f} MB)")
        cleared, kept = [], 0
        t0 = time.perf_counter()
        for i, (mail_id, rel) in enumerate(rows, 1):
            path = root / Path(rel)
            try:
                ok = path.exists() and bool(read_html(path))
            except Exception:
                ok = False
            if ok:
                cleared.append(mail_id)
            else:
                kept += 1
            if i % 50 == 0 or i == len(rows):
                print(f"  {i:,}/{len(rows):,} 확인  ({time.perf_counter() - t0:.0f}s)", flush=True)
        print(f".msg 에서 HTML 을 꺼낼 수 있음 {len(cleared):,}건 → DB 에서 비움, 못 꺼냄 {kept:,}건 → 그대로 둠")
        if dry_run:
            print("dry-run: 바꾸지 않고 끝냅니다")
            return 0
        with store.db:
            store.db.executemany("UPDATE mails SET html_body = NULL WHERE id = ?", [(m,) for m in cleared])
        print("빈 공간 정리(VACUUM) 중... 잠시 기다려 주세요", flush=True)
        store.db.execute("VACUUM")
    finally:
        store.close()
    after = db_path.stat().st_size
    print(f"완료: DB {before / 1e6:.1f} MB → {after / 1e6:.1f} MB")
    return 0
