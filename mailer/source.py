"""메일 소스 인터페이스.

수집기·probe·뷰어는 이 모듈의 타입만 알고, Outlook COM 은 모른다.
VDI 에서는 OutlookSource, Mac 에서는 FakeMailSource 를 끼운다.

VDI 의 Outlook 은 온라인 모드라 속성을 읽을 때마다 서버를 다녀온다. 그래서
- 목록은 headers() (Outlook Table API) 로 필요한 열만 한 번에 읽고
- 본문·.msg 저장이 필요한 항목만 open() 으로 연다.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Protocol

# Outlook 기본 폴더 종류. 회사·사람을 드러내지 않으므로 probe 출력에 써도 된다.
KIND_INBOX = "inbox"
KIND_SENT = "sent"
KIND_DRAFTS = "drafts"
KIND_OUTBOX = "outbox"
KIND_DELETED = "deleted"
KIND_JUNK = "junk"

# 보관 대상이 아닌 기본 폴더
EXCLUDED_KINDS = frozenset({KIND_DRAFTS, KIND_OUTBOX, KIND_DELETED, KIND_JUNK})

# 아카이브(Enterprise Vault) 로 옮겨진 뒤 남는 바로가기 항목의 MessageClass 접두어
STUB_CLASS_PREFIXES = ("IPM.Note.EnterpriseVault.Shortcut",)


def is_stub(message_class: str) -> bool:
    return message_class.startswith(STUB_CLASS_PREFIXES)


@dataclass(frozen=True)
class StoreRef:
    """Outlook 에 연결된 저장소(사서함, 아카이브, 데이터 파일 등). 이름은 갖지 않는다."""

    index: int
    is_default: bool
    kind: str  # primary / mailbox / public / additional / not-exchange / ?
    is_data_file: bool
    handle: object = field(default=None, compare=False, repr=False)


@dataclass(frozen=True)
class FolderRef:
    """폴더 하나. `handle` 은 소스 구현이 쓰는 내부 값이다."""

    path: tuple[str, ...]  # 루트부터의 이름. 화면/로그에 그대로 찍지 않는다
    item_count: int  # -1: 접근 불가
    is_mail: bool  # 메일 폴더인지 (일정·연락처 폴더 제외용)
    kind: str | None = None  # KIND_* 또는 None (사용자 폴더)
    store_index: int = 1
    handle: object = field(default=None, compare=False, repr=False)

    @property
    def depth(self) -> int:
        return len(self.path)


@dataclass(frozen=True)
class MailHeader:
    """목록용 가벼운 정보. Table API 한 번으로 여러 건을 읽는다."""

    entry_id: str
    store_id: str
    message_class: str
    subject: str
    sender_name: str
    received: datetime | None  # naive local time
    size: int
    message_id: str | None  # Internet Message-ID. 사내 메일 등은 없을 수 있다
    search_key: str | None  # PR_SEARCH_KEY (hex). 복사본끼리 같은 값

    @property
    def is_stub(self) -> bool:
        return is_stub(self.message_class)


@dataclass
class MailRecord:
    """open() 으로 연 메일 한 통."""

    header: MailHeader
    sender_email: str
    to: str
    cc: str
    sent: datetime | None
    body: str
    html_body: str
    attachments: list[str]
    handle: object = field(default=None, compare=False, repr=False)


class MailSource(Protocol):
    def describe(self) -> dict[str, str]:
        """환경 정보 (버전, 연결 모드 등). 개인정보를 넣지 않는다."""
        ...

    def stores(self) -> Iterator[StoreRef]: ...

    def folders(self, store: StoreRef | None = None) -> Iterator[FolderRef]:
        """저장소의 모든 폴더를 깊이 우선으로. store 가 None 이면 기본 사서함."""
        ...

    def headers(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
        oldest_first: bool = False,
    ) -> Iterator[MailHeader]:
        """수신 시각 내림차순(oldest_first 면 오름차순). `since` 보다 오래된 항목이 나오면 멈춘다."""
        ...

    def open(self, header: MailHeader) -> MailRecord:
        """본문·첨부까지 읽는다. 다 쓰면 바로 버린다 (온라인 모드 동시 열기 제한)."""
        ...

    def measure(self, header: MailHeader) -> dict[str, float]:
        """항목을 새로 열어 단계별 소요 시간(초)을 잰다. probe 전용."""
        ...

    def save_msg(self, record: MailRecord, path: Path) -> None:
        """원본을 .msg (유니코드) 로 저장."""
        ...
