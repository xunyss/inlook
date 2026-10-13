"""메일 소스 인터페이스.

수집기·probe·뷰어는 이 모듈의 타입만 알고, Outlook COM 은 모른다.
VDI 에서는 OutlookSource, Mac 에서는 FakeMailSource 를 끼운다.
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


@dataclass(frozen=True)
class FolderRef:
    """폴더 하나. `handle` 은 소스 구현이 쓰는 내부 값이다."""

    path: tuple[str, ...]  # 루트부터의 이름. 화면/로그에 그대로 찍지 않는다
    item_count: int
    is_mail: bool  # 메일 폴더인지 (일정·연락처 폴더 제외용)
    kind: str | None = None  # KIND_* 또는 None (사용자 폴더)
    handle: object = field(default=None, compare=False, repr=False)

    @property
    def depth(self) -> int:
        return len(self.path)


@dataclass
class MailRecord:
    """메일 한 통의 메타데이터와 본문."""

    message_class: str
    subject: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    received: datetime | None  # naive local time
    sent: datetime | None
    message_id: str | None  # Internet Message-ID. 없을 수 있다
    entry_id: str
    body: str
    html_body: str
    attachments: list[str]
    size: int
    handle: object = field(default=None, compare=False, repr=False)

    @property
    def is_stub(self) -> bool:
        return self.message_class.startswith(STUB_CLASS_PREFIXES)


class MailSource(Protocol):
    def describe(self) -> dict[str, str]:
        """환경 정보 (버전 등). 개인정보를 넣지 않는다."""
        ...

    def folders(self) -> Iterator[FolderRef]:
        """기본 사서함의 모든 폴더를 깊이 우선으로."""
        ...

    def items(
        self,
        folder: FolderRef,
        since: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[MailRecord]:
        """수신 시각 내림차순. `since` 보다 오래된 항목이 나오면 멈춘다."""
        ...

    def save_msg(self, item: MailRecord, path: Path) -> None:
        """원본을 .msg (유니코드) 로 저장."""
        ...
