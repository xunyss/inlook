"""뷰어가 쓰는 DB 조회·라벨 편집. Qt 를 모른다 (Mac 에서 테스트).

라벨은 '/' 로 나뉜 경로처럼 다룬다. 'A' 로 거르면 'A' 와 'A/...' 가 모두 걸린다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .store import Store

ALL = None  # 라벨 필터 없음
UNLABELED = ""  # 라벨이 하나도 없는 메일

_SEARCH_COLUMNS = ("subject", "sender_name", "sender_email", "to_addrs", "cc_addrs", "body", "attachments")


@dataclass(frozen=True)
class MailRow:
    id: int
    received: str
    sender: str
    subject: str
    labels: list[str]
    has_attachments: bool


@dataclass(frozen=True)
class MailDetail:
    id: int
    received: str
    sent: str
    subject: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    attachments: list[str]
    body: str
    html_body: str
    msg_path: Path
    labels: list[str]


def _escape_like(text: str) -> str:
    return text.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _clean_label(name: str) -> str:
    parts = [p.strip() for p in name.replace("\\", "/").split("/")]
    return "/".join(p for p in parts if p)


class Library:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.store = Store(root)
        self.db = self.store.db

    def close(self) -> None:
        self.store.close()

    # --- 조회 ---

    def count(self) -> int:
        return self.store.count()

    def label_counts(self) -> dict[str, int]:
        rows = self.db.execute(
            """SELECT l.name, COUNT(ml.mail_id) FROM labels l
               LEFT JOIN mail_labels ml ON ml.label_id = l.id
               GROUP BY l.id ORDER BY l.name"""
        )
        return {name: n for name, n in rows}

    def tree_counts(self) -> dict[str, int]:
        """라벨 경로마다 하위 라벨까지 합친 메일 수 (중복 없이). 목록에서 그 라벨을 눌렀을 때의 건수와 같다."""
        mails: dict[str, set[int]] = {}
        rows = self.db.execute(
            "SELECT l.name, ml.mail_id FROM mail_labels ml JOIN labels l ON l.id = ml.label_id"
        )
        for name, mail_id in rows:
            parts = name.split("/")
            for depth in range(1, len(parts) + 1):
                mails.setdefault("/".join(parts[:depth]), set()).add(mail_id)
        return {path: len(ids) for path, ids in mails.items()}

    def unlabeled_count(self) -> int:
        return self.db.execute(
            "SELECT COUNT(*) FROM mails m WHERE NOT EXISTS (SELECT 1 FROM mail_labels ml WHERE ml.mail_id = m.id)"
        ).fetchone()[0]

    def last_sync(self) -> tuple[str, str] | None:
        """(시작 시각, 상태) — 정상이 아니어도 가장 최근 것."""
        row = self.db.execute("SELECT started, status FROM syncs ORDER BY started DESC LIMIT 1").fetchone()
        return (row[0], row[1]) if row else None

    def last_ok_sync(self) -> str | None:
        row = self.db.execute(
            "SELECT started FROM syncs WHERE status = 'ok' ORDER BY started DESC LIMIT 1"
        ).fetchone()
        return row[0] if row else None

    def search(self, text: str = "", label: str | None = ALL, limit: int = 5000) -> list[MailRow]:
        where: list[str] = []
        args: list[object] = []
        for term in text.split():
            like = f"%{_escape_like(term)}%"
            where.append("(" + " OR ".join(f"m.{c} LIKE ? ESCAPE '\\'" for c in _SEARCH_COLUMNS) + ")")
            args.extend([like] * len(_SEARCH_COLUMNS))
        if label == UNLABELED:
            where.append("NOT EXISTS (SELECT 1 FROM mail_labels x WHERE x.mail_id = m.id)")
        elif label is not None:
            where.append(
                """EXISTS (SELECT 1 FROM mail_labels x JOIN labels y ON y.id = x.label_id
                          WHERE x.mail_id = m.id AND (y.name = ? OR y.name LIKE ? ESCAPE '\\'))"""
            )
            args.extend([label, _escape_like(label) + "/%"])
        sql = (
            "SELECT m.id, m.received, m.sender_name, m.sender_email, m.subject, m.attachments FROM mails m"
            + (" WHERE " + " AND ".join(where) if where else "")
            + " ORDER BY m.received DESC, m.id DESC LIMIT ?"
        )
        args.append(limit)
        rows = self.db.execute(sql, args).fetchall()
        labels = self._labels_for([r[0] for r in rows])
        return [
            MailRow(
                id=r[0],
                received=r[1] or "",
                sender=r[2] or r[3] or "",
                subject=r[4] or "",
                labels=labels.get(r[0], []),
                has_attachments=bool(r[5] and r[5] != "[]"),
            )
            for r in rows
        ]

    def _labels_for(self, ids: list[int]) -> dict[int, list[str]]:
        result: dict[int, list[str]] = {}
        for start in range(0, len(ids), 500):
            chunk = ids[start:start + 500]
            marks = ",".join("?" * len(chunk))
            rows = self.db.execute(
                f"""SELECT ml.mail_id, l.name FROM mail_labels ml JOIN labels l ON l.id = ml.label_id
                    WHERE ml.mail_id IN ({marks}) ORDER BY l.name""",
                chunk,
            )
            for mail_id, name in rows:
                result.setdefault(mail_id, []).append(name)
        return result

    def get(self, mail_id: int) -> MailDetail | None:
        r = self.db.execute(
            """SELECT id, received, sent, subject, sender_name, sender_email, to_addrs, cc_addrs,
                      attachments, body, html_body, msg_path FROM mails WHERE id = ?""",
            (mail_id,),
        ).fetchone()
        if r is None:
            return None
        try:
            attachments = json.loads(r[8] or "[]")
        except ValueError:
            attachments = []
        return MailDetail(
            id=r[0], received=r[1] or "", sent=r[2] or "", subject=r[3] or "",
            sender_name=r[4] or "", sender_email=r[5] or "", to=r[6] or "", cc=r[7] or "",
            attachments=attachments, body=r[9] or "", html_body=r[10] or "",
            msg_path=self.root / Path(r[11]), labels=self._labels_for([r[0]]).get(r[0], []),
        )

    # --- 라벨 편집 ---

    def add_label(self, mail_ids: list[int], name: str) -> str:
        name = _clean_label(name)
        if not name:
            raise ValueError("빈 라벨 이름")
        with self.db:
            self.db.execute("INSERT OR IGNORE INTO labels (name) VALUES (?)", (name,))
            (label_id,) = self.db.execute("SELECT id FROM labels WHERE name = ?", (name,)).fetchone()
            self.db.executemany(
                "INSERT OR IGNORE INTO mail_labels (mail_id, label_id) VALUES (?, ?)",
                [(m, label_id) for m in mail_ids],
            )
        return name

    def remove_label(self, mail_ids: list[int], name: str) -> None:
        with self.db:
            self.db.executemany(
                "DELETE FROM mail_labels WHERE mail_id = ? AND label_id = (SELECT id FROM labels WHERE name = ?)",
                [(m, name) for m in mail_ids],
            )
            self._drop_unused()

    def rename_label(self, old: str, new: str) -> str:
        """old 와 그 하위 라벨(old/...)을 new 아래로 옮긴다. 이미 있는 라벨이면 합친다."""
        new = _clean_label(new)
        if not new:
            raise ValueError("빈 라벨 이름")
        with self.db:
            rows = self.db.execute(
                "SELECT id, name FROM labels WHERE name = ? OR name LIKE ? ESCAPE '\\'",
                (old, _escape_like(old) + "/%"),
            ).fetchall()
            for label_id, name in rows:
                target = new + name[len(old):]
                if target == name:
                    continue
                self.db.execute("INSERT OR IGNORE INTO labels (name) VALUES (?)", (target,))
                (target_id,) = self.db.execute("SELECT id FROM labels WHERE name = ?", (target,)).fetchone()
                self.db.execute(
                    "INSERT OR IGNORE INTO mail_labels (mail_id, label_id) SELECT mail_id, ? FROM mail_labels WHERE label_id = ?",
                    (target_id, label_id),
                )
                self.db.execute("DELETE FROM mail_labels WHERE label_id = ?", (label_id,))
            self._drop_unused()
        return new

    def delete_label(self, name: str) -> None:
        """라벨만 지운다 (하위 포함). 메일은 지우지 않는다."""
        with self.db:
            self.db.execute(
                """DELETE FROM mail_labels WHERE label_id IN
                   (SELECT id FROM labels WHERE name = ? OR name LIKE ? ESCAPE '\\')""",
                (name, _escape_like(name) + "/%"),
            )
            self._drop_unused()

    def _drop_unused(self) -> None:
        self.db.execute("DELETE FROM labels WHERE id NOT IN (SELECT DISTINCT label_id FROM mail_labels)")
