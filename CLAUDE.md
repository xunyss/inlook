# mailer — Outlook 메일 보존 도구

> 이 저장소는 **public** 이다. 회사명·메일 도메인·사람 이름·폴더명(고객사명)·
> 실제 메일 내용을 이 파일과 코드, 커밋 메시지, 테스트 데이터에 넣지 않는다.

## 목적
회사 Exchange 메일이 보존 정책으로 사라지기 전에 사본(.msg)을 남기고, 검색해서 다시 읽을 수 있게 한다.

## 실행 환경 제약 (2026-10-13 확인)
- 계정: Microsoft Exchange. 회사 VDI(Windows) 안의 Outlook 으로만 접근 가능 (전용망)
- 보존 정책: 메일함에서 약 1개월 후 아카이브로 이동, 그 후 약 2개월 뒤 삭제. 아카이브는 실사용하지 않음
- VDI에서 되는 것: Python(버전 자유 설치) + pywin32, Outlook VBA, `.msg` SaveAs, git clone
- VDI에서 안 되는 것: PowerShell, Claude Code

## 방향
- 송수신 클라이언트가 아니라 "Outlook 옆의 보존기 + 검색 뷰어"
- 데이터는 VDI 안에만 저장한다. VDI 밖으로 반출하는 설계는 하지 않는다

## 결정 사항 (2026-10-13)
- 주력: **Python(>=3.12) + pywin32 로 Outlook COM 자동화**. VBA는 필요시 "도착 즉시 저장" 보조
- 의존성 최소화: pywin32(수집), PySide6(뷰어) 외에는 표준 라이브러리(sqlite3 등)
- pywin32는 Windows 전용 → `sys_platform == 'win32'` 마커로 Mac 개발 환경과 공존
- `[tool.uv] package = false` — 패키지로 설치하지 않고 저장소 루트에서 `python -m mailer` 로 실행
- 뷰어는 **자체 데스크톱 UI (PySide6)**. 브라우저·로컬 웹서버 방식은 쓰지 않는다
- 중복 판단 키는 Internet Message-ID (EntryID는 폴더 이동/아카이브 시 바뀜)
- `.msg` 저장 형식은 olMSGUnicode(9). olMSG(3)는 ANSI라 한글이 깨질 수 있음
- 증분 조회는 `Items.Sort("[ReceivedTime]", True)` 후 기준일보다 오래되면 중단.
  `Restrict` 날짜 문자열은 로케일 의존이라 쓰지 않는다

## 개발·테스트 전략 (Mac에서 개발, VDI에서 실행)
- Outlook COM 접근은 `MailSource` 인터페이스(`mailer/source.py`) 뒤로 격리한다
  - `OutlookSource`(VDI 전용, 얇게) / `FakeMailSource`(Mac용 합성 데이터)
- 테스트 데이터는 합성 데이터만 쓴다
- COM 동작은 VDI에서 `probe.bat`(자가진단)으로 검증하고 결과 화면만 받아본다
- 수집기는 `--dry-run` 지원 예정

## 배포 흐름
- Mac(Claude Code) → git push → VDI에서 git clone/pull. **VDI는 받기만 한다 (push 안 함)**
- VDI 결과는 사용자가 화면 캡처/텍스트로 전달한다. 그래서:
  - VDI 쪽 실행은 `.bat` 더블클릭 한 번으로 (`.bat` 은 ASCII 전용, CRLF)
  - 오류는 `[E코드] 원인 / 조치` 형식으로, 화면 한 장만 보고 원인을 알 수 있게
  - probe 출력에는 메일 내용·사람 이름·폴더명을 넣지 않는다. 개수·통과여부·익명 ID(`folder#07`)만
- 저장소에 실제 메일·DB·로컬 설정·로그를 커밋하지 않는다 (.gitignore)

## Mac 개발 명령
```bash
uv sync                           # .venv 생성 + 의존성 (pytest 포함)
uv run python -m mailer probe --fake
uv run pytest
```
