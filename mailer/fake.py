"""Mac 개발용 가짜 MailSource. 전부 합성 데이터다.

실제 환경에서 문제가 될 만한 경우를 일부러 넣어둔다:
같은 분·같은 제목, 아주 긴 제목, 파일명 금지 문자, 빈 제목, Message-ID 없음
(보낸 메일 등), 회의 요청, 배달 실패 보고서, 아카이브 바로가기(stub),
깊은 하위 폴더, 일정 폴더, 두 번째 저장소(아카이브)에 있는 오래된 메일.
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
    MailHeader,
    MailRecord,
    StoreRef,
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
        self._stores = [
            StoreRef(index=1, is_default=True, kind="primary", is_data_file=False),
            StoreRef(index=2, is_default=False, kind="not-exchange", is_data_file=True),
        ]
        self._folders: dict[int, list[FolderRef]] = {1: [], 2: []}
        self._items: dict[tuple[int, tuple[str, ...]], list[MailRecord]] = {}
        self._by_id: dict[str, MailRecord] = {}
        self._seq = 0
        self._build()

    # --- MailSource ---

    def describe(self) -> dict[str, str]:
        return {"python": sys.version.split()[0], "source": "fake", "mode": "online"}

    def stores(self) -> Iterator[StoreRef]:
        yield from self._stores

    def folders(self, store: StoreRef | None = None) -> Iterator[FolderRef]:
        yield from self._folders[store.index if store else 1]

    def headers(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[MailHeader]:
        records = sorted(
            self._items.get((folder.store_index, folder.path), []),
            key=lambda r: r.header.received or datetime.min,
            reverse=True,
        )
        count = 0
        for record in records:
            header = record.header
            if since and header.received and header.received < since:
                break
            yield header
            count += 1
            if limit and count >= limit:
                break

    def open(self, header: MailHeader) -> MailRecord:
        return self._by_id[header.entry_id]

    def measure(self, header: MailHeader) -> dict[str, float]:
        self.open(header)
        return {"open": 0.05, "basic": 0.01, "sender_smtp": 0.01, "body": 0.1, "html": 0.2, "attachments": 0.02}

    def save_msg(self, record: MailRecord, path: Path) -> None:
        if self.fail_save:
            raise OSError("fake save failure")
        path.write_bytes(b"FAKE-MSG\n" + record.header.subject.encode("utf-8"))

    # --- 합성 데이터 ---

    def _folder(self, store: int, path: tuple[str, ...], kind: str | None = None, is_mail: bool = True) -> None:
        self._folders[store].append(FolderRef(path=path, item_count=0, is_mail=is_mail, kind=kind, store_index=store))
        self._items[(store, path)] = []

    def _mail(
        self,
        path: tuple[str, ...],
        days_ago: float,
        subject: str,
        *,
        store: int = 1,
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
        stub = message_class.startswith("IPM.Note.EnterpriseVault")
        body = "아카이브된 항목입니다." if stub else f"안녕하세요, {name}입니다.\n번호: {n}\n감사합니다."
        header = MailHeader(
            entry_id=f"FAKEENTRY{n:08d}",
            store_id=f"FAKESTORE{store}",
            message_class=message_class,
            subject=subject,
            sender_name=name,
            received=when,
            size=len(body) * 3,
            message_id=f"<fake-{n:05d}@example.com>" if message_id else None,
            search_key=f"{n:032X}",
        )
        record = MailRecord(
            header=header,
            sender_email=email,
            to="me@example.com",
            cc="",
            sent=when - timedelta(minutes=1),
            body=body,
            html_body="" if stub else _HTML.format(name=name, n=n),
            attachments=attachments or [],
        )
        self._items[(store, path)].append(record)
        self._by_id[header.entry_id] = record
        return record

    def _build(self) -> None:
        inbox = ("받은 편지함",)
        proj = inbox + ("프로젝트A",)
        minutes = proj + ("회의록",)
        news = inbox + ("뉴스레터",)
        sent = ("보낸 편지함",)
        keep = ("보관",)

        self._folder(1, inbox, KIND_INBOX)
        self._folder(1, proj)
        self._folder(1, minutes)
        self._folder(1, news)
        self._folder(1, sent, KIND_SENT)
        self._folder(1, ("지운 편지함",), KIND_DELETED)
        self._folder(1, ("정크 메일",), KIND_JUNK)
        self._folder(1, ("일정",), is_mail=False)
        self._folder(1, keep)
        self._folder(2, inbox)
        self._folder(2, proj)

        # 평범한 메일: 기본 사서함에는 최근 30일 안쪽만 남아 있다
        for i in range(30):
            self._mail(inbox, i, f"주간 업무 공유 #{i}", attachments=["보고서.docx"] if i % 5 == 0 else None)
        for i in range(10):
            self._mail(proj, i * 3 + 1, f"[프로젝트A] 진행 현황 {i}", message_id=i % 3 != 0)
        for i in range(5):
            self._mail(minutes, i * 6 + 2, f"[회의록] {i}차 회의", attachments=["회의록.pdf", "첨부 자료 (최종).xlsx"])
        for i in range(30):
            self._mail(news, i, f"News Letter {i}")
        for i in range(20):
            # 사내 Exchange 로 보낸 메일은 Message-ID 가 없을 수 있다
            self._mail(sent, i * 1.5, f"RE: 문의 드립니다 {i}", sender=0, message_id=i % 2 == 0)

        # 까다로운 경우
        same = self.now - timedelta(days=5, hours=3)
        self._mail(inbox, 0, "RE: 확인 부탁드립니다", received=same)
        self._mail(inbox, 0, "RE: 확인 부탁드립니다", received=same)
        self._mail(inbox, 6, "아주 긴 제목 " + "가나다라마바사 " * 30)
        self._mail(inbox, 7, 'a/b:c*d?e"f<g>h|i 금지 문자 제목')
        self._mail(inbox, 8, "")
        self._mail(inbox, 10, "회의: 분기 리뷰", message_class="IPM.Schedule.Meeting.Request", message_id=False)
        self._mail(inbox, 11, "배달되지 않음: 테스트", message_class="REPORT.IPM.Note.NDR", message_id=False)
        for i in range(3):
            self._mail(inbox, 25 + i, f"바로가기 {i}", message_class="IPM.Note.EnterpriseVault.Shortcut")
        self._mail(keep, 200, "1년 가까이 된 보관 메일")

        # 두 번째 저장소: 1~3개월 전 메일
        for i in range(40):
            self._mail(inbox, 31 + i * 1.5, f"지난달 메일 {i}", store=2, message_id=i % 4 != 0)
        for i in range(10):
            self._mail(proj, 35 + i * 5, f"[프로젝트A] 지난 현황 {i}", store=2)

        for store, folders in self._folders.items():
            self._folders[store] = [
                FolderRef(
                    path=f.path,
                    item_count=len(self._items[(store, f.path)]),
                    is_mail=f.is_mail,
                    kind=f.kind,
                    store_index=store,
                )
                for f in folders
            ]
