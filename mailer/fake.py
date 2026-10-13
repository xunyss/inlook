"""Mac 개발용 가짜 MailSource. 전부 합성 데이터다.

실제 환경에서 문제가 될 만한 경우를 일부러 넣어둔다:
같은 분·같은 제목, 아주 긴 제목, 파일명 금지 문자, 빈 제목, Message-ID 없음,
회의 요청, 배달 실패 보고서, 아카이브 바로가기(stub), 깊은 하위 폴더, 일정 폴더.
"""

from __future__ import annotations

import sys
from collections.abc import Iterator
from datetime import datetime, timedelta
from pathlib import Path

from .source import (
    KIND_DELETED,
    KIND_INBOX,
    KIND_JUNK,
    KIND_SENT,
    FolderRef,
    MailRecord,
)

_PEOPLE = [
    ("홍길동", "gildong.hong@example.com"),
    ("김철수", "chulsoo.kim@example.com"),
    ("이영희", "younghee.lee@example.com"),
    ("Jane Doe", "jane.doe@example.com"),
]

_HTML = (
    "<html><body><p>안녕하세요, {name}입니다.</p>"
    "<table border=1><tr><th>항목</th><th>값</th></tr>"
    "<tr><td>번호</td><td>{n}</td></tr></table>"
    "<p>감사합니다.</p></body></html>"
)


class FakeMailSource:
    def __init__(self, now: datetime | None = None, *, fail_save: bool = False) -> None:
        self.now = (now or datetime.now()).replace(second=0, microsecond=0)
        self.fail_save = fail_save
        self._folders: list[FolderRef] = []
        self._items: dict[tuple[str, ...], list[MailRecord]] = {}
        self._seq = 0
        self._build()

    # --- MailSource ---

    def describe(self) -> dict[str, str]:
        return {"python": sys.version.split()[0], "source": "fake"}

    def folders(self) -> Iterator[FolderRef]:
        yield from self._folders

    def items(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[MailRecord]:
        records = sorted(
            self._items.get(folder.path, []),
            key=lambda r: r.received or datetime.min,
            reverse=True,
        )
        count = 0
        for record in records:
            if since and record.received and record.received < since:
                break
            yield record
            count += 1
            if limit and count >= limit:
                break

    def save_msg(self, item: MailRecord, path: Path) -> None:
        if self.fail_save:
            raise OSError("fake save failure")
        path.write_bytes(b"FAKE-MSG\n" + item.subject.encode("utf-8"))

    # --- 합성 데이터 ---

    def _folder(self, path: tuple[str, ...], kind: str | None = None, is_mail: bool = True) -> None:
        self._folders.append(FolderRef(path=path, item_count=0, is_mail=is_mail, kind=kind))
        self._items[path] = []

    def _mail(
        self,
        path: tuple[str, ...],
        days_ago: float,
        subject: str,
        *,
        message_class: str = "IPM.Note",
        message_id: bool = True,
        sender: int | None = None,
        attachments: list[str] | None = None,
        received: datetime | None = None,
    ) -> MailRecord:
        self._seq += 1
        n = self._seq
        name, email = _PEOPLE[(sender if sender is not None else n) % len(_PEOPLE)]
        when = received or self.now - timedelta(days=days_ago)
        is_stub = message_class.startswith("IPM.Note.EnterpriseVault")
        body = "아카이브된 항목입니다." if is_stub else f"안녕하세요, {name}입니다.\n번호: {n}\n감사합니다."
        record = MailRecord(
            message_class=message_class,
            subject=subject,
            sender_name=name,
            sender_email=email,
            to="me@example.com",
            cc="",
            received=when,
            sent=when - timedelta(minutes=1),
            message_id=f"<fake-{n:05d}@example.com>" if message_id else None,
            entry_id=f"FAKEENTRY{n:08d}",
            body=body,
            html_body="" if is_stub else _HTML.format(name=name, n=n),
            attachments=attachments or [],
            size=len(body) * 3,
        )
        self._items[path].append(record)
        return record

    def _build(self) -> None:
        inbox = ("받은 편지함",)
        proj = inbox + ("프로젝트A",)
        minutes = proj + ("회의록",)
        news = inbox + ("뉴스레터",)
        sent = ("보낸 편지함",)
        keep = ("보관",)

        self._folder(inbox, KIND_INBOX)
        self._folder(proj)
        self._folder(minutes)
        self._folder(news)
        self._folder(sent, KIND_SENT)
        self._folder(("지운 편지함",), KIND_DELETED)
        self._folder(("정크 메일",), KIND_JUNK)
        self._folder(("일정",), is_mail=False)
        self._folder(keep)

        # 평범한 메일: 0~120일 전에 고르게
        for i in range(60):
            self._mail(inbox, i * 2, f"주간 업무 공유 #{i}", attachments=["보고서.docx"] if i % 5 == 0 else None)
        for i in range(15):
            self._mail(proj, i * 7 + 1, f"[프로젝트A] 진행 현황 {i}")
        for i in range(5):
            self._mail(minutes, i * 10 + 3, f"[회의록] {i}차 회의", attachments=["회의록.pdf", "첨부 자료 (최종).xlsx"])
        for i in range(30):
            self._mail(news, i, f"News Letter {i}")
        for i in range(20):
            self._mail(sent, i * 3, f"RE: 문의 드립니다 {i}", sender=0)

        # 까다로운 경우
        same = self.now - timedelta(days=5, hours=3)
        self._mail(inbox, 0, "RE: 확인 부탁드립니다", received=same)
        self._mail(inbox, 0, "RE: 확인 부탁드립니다", received=same)
        self._mail(inbox, 6, "아주 긴 제목 " + "가나다라마바사 " * 30)
        self._mail(inbox, 7, 'a/b:c*d?e"f<g>h|i 금지 문자 제목')
        self._mail(inbox, 8, "")
        self._mail(inbox, 9, "Message-ID 없는 메일", message_id=False)
        self._mail(inbox, 10, "회의: 분기 리뷰", message_class="IPM.Schedule.Meeting.Request")
        self._mail(inbox, 11, "배달되지 않음: 테스트", message_class="REPORT.IPM.Note.NDR", message_id=False)
        for i in range(8):
            self._mail(inbox, 40 + i, f"오래된 메일 {i}", message_class="IPM.Note.EnterpriseVault.Shortcut")
        self._mail(keep, 200, "1년 가까이 된 보관 메일")

        self._folders = [
            FolderRef(path=f.path, item_count=len(self._items[f.path]), is_mail=f.is_mail, kind=f.kind)
            for f in self._folders
        ]
