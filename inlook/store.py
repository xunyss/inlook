"""보존소: archive_root 아래의 SQLite DB + .msg 파일.

inlook 안에는 폴더 개념이 없다. 메일은 수신 시각 기준 경로에 저장하고,
Outlook 폴더는 처음 저장할 때 라벨로만 남긴다.

    <archive_root>/inlook.db
    <archive_root>/msg/YYYY/MM/YYYYMMDD-HHMMSS_<hash12>.msg
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .source import MailHeader, MailRecord

DB_NAME = "inlook.db"
LEGACY_DB_NAME = "mailer.db"  # 이름을 바꾸기 전(2026-10-15)의 DB. 처음 열 때 inlook.db 로 옮긴다
MSG_DIR = "msg"
SCHEMA_VERSION = 1

_SCHEMA = """
CREATE TABLE IF NOT EXISTS mails (
    id            INTEGER PRIMARY KEY,
    key           TEXT NOT NULL UNIQUE,   -- 중복 판단 키 (mail_key)
    key_kind      TEXT NOT NULL,          -- message_id | search_key | entry_id
    message_id    TEXT,
    search_key    TEXT,
    received      TEXT,                   -- 'YYYY-MM-DD HH:MM:SS' 로컬 시각
    sent          TEXT,
    message_class TEXT,
    subject       TEXT,
    sender_name   TEXT,
    sender_email  TEXT,
    to_addrs      TEXT,
    cc_addrs      TEXT,
    size          INTEGER,
    attachments   TEXT,                   -- 파일명 JSON 배열
    body          TEXT,
    html_body     TEXT,
    msg_path      TEXT NOT NULL,          -- archive_root 기준 상대 경로 ('/' 구분)
    collected_at  TEXT NOT NULL,
    sync_id       INTEGER
);
CREATE INDEX IF NOT EXISTS mails_received ON mails(received);

CREATE TABLE IF NOT EXISTS labels (
    id   INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS mail_labels (
    mail_id  INTEGER NOT NULL REFERENCES mails(id),
    label_id INTEGER NOT NULL REFERENCES labels(id),
    PRIMARY KEY (mail_id, label_id)
);

-- 사용자가 inlook 에서 삭제한 메일의 키. 내용은 남기지 않는다. 다음 동기화가 다시 수집하지 않게 한다
CREATE TABLE IF NOT EXISTS deleted (
    key        TEXT PRIMARY KEY,
    deleted_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS syncs (
    id       INTEGER PRIMARY KEY,
    started  TEXT NOT NULL,
    finished TEXT,
    cutoff   TEXT,                        -- 이 시각보다 오래된 메일은 보지 않았다 (NULL: 전체)
    status   TEXT NOT NULL,               -- running | ok | partial | stopped
    listed   INTEGER DEFAULT 0,
    new      INTEGER DEFAULT 0,
    saved    INTEGER DEFAULT 0,
    failed   INTEGER DEFAULT 0
);
"""


def _ts(value: datetime | None) -> str | None:
    return f"{value:%Y-%m-%d %H:%M:%S}" if value else None


def _parse(value: str | None) -> datetime | None:
    return datetime.strptime(value, "%Y-%m-%d %H:%M:%S") if value else None


def mail_key(header: MailHeader) -> tuple[str, str]:
    """(키, 종류). Message-ID → SearchKey → EntryID 순.

    EntryID 는 폴더를 옮기면 바뀌지만, 키가 없다고 메일을 버리는 것보다는 낫다.
    """
    if header.message_id and header.message_id.strip():
        return header.message_id.strip(), "message_id"
    if header.search_key:
        return "sk:" + header.search_key, "search_key"
    return "eid:" + header.entry_id, "entry_id"


def msg_relpath(key: str, when: datetime) -> str:
    """같은 메일은 언제 수집해도 같은 경로. 제목·폴더명은 넣지 않는다 (260자 제한, 금지 문자)."""
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]
    return f"{MSG_DIR}/{when:%Y}/{when:%m}/{when:%Y%m%d-%H%M%S}_{digest}.msg"


@dataclass(frozen=True)
class LastSync:
    started: datetime
    cutoff: datetime | None


class Store:
    def __init__(self, root: Path) -> None:
        self.root = root
        root.mkdir(parents=True, exist_ok=True)
        legacy = root / LEGACY_DB_NAME
        if legacy.exists() and not (root / DB_NAME).exists():
            for suffix in ("-journal", ""):  # 저널이 남아 있으면 함께 옮겨야 복구된다
                old = root / (LEGACY_DB_NAME + suffix)
                if old.exists():
                    old.rename(root / (DB_NAME + suffix))
        self.db = sqlite3.connect(root / DB_NAME)
        self.db.execute("PRAGMA foreign_keys = ON")
        self.db.executescript(_SCHEMA)
        self.db.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        self.db.commit()

    def close(self) -> None:
        self.db.close()

    # --- 동기화 기록 ---

    def last_ok_sync(self) -> LastSync | None:
        row = self.db.execute(
            "SELECT started, cutoff FROM syncs WHERE status = 'ok' ORDER BY started DESC LIMIT 1"
        ).fetchone()
        return LastSync(_parse(row[0]), _parse(row[1])) if row else None

    def begin_sync(self, started: datetime, cutoff: datetime | None) -> int:
        cur = self.db.execute(
            "INSERT INTO syncs (started, cutoff, status) VALUES (?, ?, 'running')",
            (_ts(started), _ts(cutoff)),
        )
        self.db.commit()
        return int(cur.lastrowid)

    def end_sync(self, sync_id: int, status: str, *, listed: int, new: int, saved: int, failed: int) -> None:
        self.db.execute(
            "UPDATE syncs SET finished = ?, status = ?, listed = ?, new = ?, saved = ?, failed = ? WHERE id = ?",
            (_ts(datetime.now()), status, listed, new, saved, failed, sync_id),
        )
        self.db.commit()

    # --- 메일 ---

    def has(self, key: str) -> bool:
        return self.db.execute("SELECT 1 FROM mails WHERE key = ?", (key,)).fetchone() is not None

    def known(self, key: str) -> bool:
        """저장돼 있거나 사용자가 삭제한 메일. 동기화는 이것만 건너뛴다."""
        return self.has(key) or self.db.execute(
            "SELECT 1 FROM deleted WHERE key = ?", (key,)
        ).fetchone() is not None

    def delete_mails(self, mail_ids: list[int]) -> list[int]:
        """.msg 파일과 DB 기록을 지운다. 파일을 못 지운 메일(다른 프로그램이 열고 있음 등)은
        DB 에 그대로 두고 그 id 를 돌려준다."""
        failed: list[int] = []
        for mail_id in mail_ids:
            row = self.db.execute("SELECT key, msg_path FROM mails WHERE id = ?", (mail_id,)).fetchone()
            if row is None:
                continue
            key, rel = row
            path = self.root / Path(rel)
            try:
                path.unlink(missing_ok=True)
            except OSError:
                failed.append(mail_id)
                continue
            with self.db:
                self.db.execute(
                    "INSERT OR REPLACE INTO deleted (key, deleted_at) VALUES (?, ?)", (key, _ts(datetime.now()))
                )
                self.db.execute("DELETE FROM mail_labels WHERE mail_id = ?", (mail_id,))
                self.db.execute("DELETE FROM mails WHERE id = ?", (mail_id,))
            for folder in (path.parent, path.parent.parent):  # 비어 버린 YYYY/MM 폴더 정리
                try:
                    folder.rmdir()
                except OSError:
                    break
        with self.db:
            self.db.execute("DELETE FROM labels WHERE id NOT IN (SELECT DISTINCT label_id FROM mail_labels)")
        return failed

    def add(
        self,
        key: str,
        key_kind: str,
        record: MailRecord,
        msg_path: str,
        labels: list[str],
        sync_id: int | None,
    ) -> None:
        """메일 한 건과 라벨을 한 트랜잭션으로 넣는다."""
        h = record.header
        with self.db:
            cur = self.db.execute(
                """INSERT INTO mails (key, key_kind, message_id, search_key, received, sent, message_class,
                       subject, sender_name, sender_email, to_addrs, cc_addrs, size, attachments,
                       body, html_body, msg_path, collected_at, sync_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    key, key_kind, h.message_id, h.search_key, _ts(h.received), _ts(record.sent),
                    h.message_class, h.subject, h.sender_name, record.sender_email, record.to, record.cc,
                    h.size, json.dumps(record.attachments, ensure_ascii=False),
                    record.body, record.html_body, msg_path, _ts(datetime.now()), sync_id,
                ),
            )
            mail_id = cur.lastrowid
            for name in labels:
                self.db.execute("INSERT OR IGNORE INTO labels (name) VALUES (?)", (name,))
                (label_id,) = self.db.execute("SELECT id FROM labels WHERE name = ?", (name,)).fetchone()
                self.db.execute(
                    "INSERT OR IGNORE INTO mail_labels (mail_id, label_id) VALUES (?, ?)", (mail_id, label_id)
                )

    def fix_labels(self, renames: dict[str, str]) -> int:
        """라벨 이름을 정확히 일치하는 것만 바꾼다 (하위 라벨은 각자 따로). 바꾼 라벨 수."""
        fixed = 0
        with self.db:
            for old, new in renames.items():
                row = self.db.execute("SELECT id FROM labels WHERE name = ?", (old,)).fetchone()
                if row is None:
                    continue
                self.db.execute("INSERT OR IGNORE INTO labels (name) VALUES (?)", (new,))
                (new_id,) = self.db.execute("SELECT id FROM labels WHERE name = ?", (new,)).fetchone()
                self.db.execute(
                    "INSERT OR IGNORE INTO mail_labels (mail_id, label_id) SELECT mail_id, ? FROM mail_labels WHERE label_id = ?",
                    (new_id, row[0]),
                )
                self.db.execute("DELETE FROM mail_labels WHERE label_id = ?", (row[0],))
                self.db.execute("DELETE FROM labels WHERE id = ?", (row[0],))
                fixed += 1
        return fixed

    def count(self) -> int:
        return self.db.execute("SELECT COUNT(*) FROM mails").fetchone()[0]

    def labels_of(self, key: str) -> list[str]:
        rows = self.db.execute(
            """SELECT l.name FROM labels l
               JOIN mail_labels ml ON ml.label_id = l.id
               JOIN mails m ON m.id = ml.mail_id
               WHERE m.key = ? ORDER BY l.name""",
            (key,),
        )
        return [r[0] for r in rows]
