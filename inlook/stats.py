"""DB 크기 진단: 어떤 열이 공간을 차지하는지. 메일 내용·제목은 찍지 않는다 (크기와 개수만)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from .store import DB_NAME

_COLUMNS = ("body", "html_body", "subject", "to_addrs", "cc_addrs", "attachments", "sender_email", "message_id")


def _mb(n: int | None) -> str:
    return f"{(n or 0) / 1_000_000:8.1f} MB"


def run_stats(root: Path) -> int:
    path = root / DB_NAME
    if not path.exists():
        print(f"DB 가 없습니다: {path}")
        return 2
    db = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    count = db.execute("SELECT COUNT(*) FROM mails").fetchone()[0]
    page_size = db.execute("PRAGMA page_size").fetchone()[0]
    pages = db.execute("PRAGMA page_count").fetchone()[0]
    free = db.execute("PRAGMA freelist_count").fetchone()[0]
    print("=" * 60)
    print(f" DB 파일      {_mb(path.stat().st_size)}   메일 {count:,}건")
    print(f" 빈 공간      {_mb(free * page_size)}   (삭제 후 남은 자리, 전체 {pages:,} 페이지)")
    print("-" * 60)
    print(" 열별 합계 (UTF-8 바이트)")
    for col in _COLUMNS:
        total = db.execute(f"SELECT SUM(LENGTH(CAST({col} AS BLOB))) FROM mails").fetchone()[0]
        print(f"   {col:<14}{_mb(total)}")
    print("-" * 60)
    data_uri = db.execute(
        "SELECT COUNT(*), SUM(LENGTH(CAST(html_body AS BLOB))) FROM mails WHERE html_body LIKE '%data:image%'"
    ).fetchone()
    print(f" HTML 안에 이미지가 글자로 박힌 메일(data:image) {data_uri[0]:,}건, 그 HTML 합계 {_mb(data_uri[1]).strip()}")
    sizes = [r[0] for r in db.execute(
        "SELECT LENGTH(CAST(html_body AS BLOB)) AS n FROM mails ORDER BY n DESC"
    ) if r[0]]
    if sizes:
        top = sizes[:10]
        print(f" HTML 큰 순서 10건: {', '.join(f'{n / 1_000_000:.1f}MB' for n in top)}")
        print(f" HTML 중간값: {sizes[len(sizes) // 2] / 1000:.0f} KB")
    for limit in (1_000_000, 100_000):
        n = sum(1 for s in sizes if s >= limit)
        print(f" HTML {limit // 1000:>5} KB 이상: {n:,}건, 합계 {_mb(sum(s for s in sizes if s >= limit)).strip()}")
    print("=" * 60)
    db.close()
    return 0
