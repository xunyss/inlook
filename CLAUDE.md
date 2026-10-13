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
- **Outlook 은 온라인 모드** (캐시 모드 아님). 속성을 처음 읽을 때마다 Exchange 서버 왕복
- probe v1 결과: 기본 사서함 메일 693건, 메일 폴더 31개(깊이 2), 저장소 2개,
  항목별 COM 읽기 초당 약 2건, Message-ID 있음 69%, stub 0건, LongPathsEnabled 꺼짐
- probe v4 결과 (2026-10-13, 기본 사서함만, 표본): 정상 종료, Outlook 문제 없음
  - 수집 대상 폴더 10개(받은 편지함+하위 8개, 보낸 편지함) 490건, 지운 편지함 213건
  - 폴더별 가장 오래된 메일이 대부분 3~5주 전 → 실제 보존 기간 약 1개월 (예외적으로 오래된 메일 일부 존재)
  - Table API 로 읽으면 Message-ID·SearchKey 모두 100% (v1 의 69% 는 항목별 PropertyAccessor 방식 탓으로 추정)
  - Table 목록 초당 약 52행, 메일 1건 열어 읽기 약 0.47초, 1.8MB 메일 열기+.msg 2회 저장 4.5초

## 방향
- 송수신 클라이언트가 아니라 "Outlook 옆의 보존기 + 검색 뷰어"
- 데이터는 VDI 안에만 저장한다. VDI 밖으로 반출하는 설계는 하지 않는다

## 결정 사항 (2026-10-13)
- 주력: **Python(>=3.12) + pywin32 로 Outlook COM 자동화**. VBA는 필요시 "도착 즉시 저장" 보조
- 의존성 최소화: pywin32(수집), PySide6(뷰어) 외에는 표준 라이브러리(sqlite3 등)
- pywin32는 Windows 전용 → `sys_platform == 'win32'` 마커로 Mac 개발 환경과 공존
- `[tool.uv] package = false` — 패키지로 설치하지 않고 저장소 루트에서 `python -m mailer` 로 실행
- 뷰어는 **자체 데스크톱 UI (PySide6)**. 브라우저·로컬 웹서버 방식은 쓰지 않는다
- 중복 판단 키: **Internet Message-ID**, 없으면 SearchKey (probe v4 에서 둘 다 100%, 중복 0).
  EntryID는 폴더 이동/아카이브 시 바뀌므로 키로 쓰지 않는다
- `.msg` 저장 형식은 olMSGUnicode(9). olMSG(3)는 ANSI라 한글이 깨질 수 있음
- 목록은 **Table API**(`Folder.GetTable`)로 필요한 열만 읽는다 (`headers()`).
  본문·첨부·.msg 저장은 새 메일만 `open()` 해서 한다 (온라인 모드라 건당 서버 왕복)
- 증분 조회는 Table 을 `[ReceivedTime]` 내림차순 정렬 후 기준일보다 오래되면 중단.
  `Restrict`/Table 필터의 날짜 문자열은 로케일 의존이라 쓰지 않는다
- 온라인 모드의 동시 열기 제한(Exchange 기본 약 250개) 때문에 연 메일 객체는 처리 후 바로 버린다
- **기본 Exchange 사서함만 다룬다. 다른 저장소는 절대 건드리지 않는다** (사용자 지시, 2026-10-13)
  - Outlook 에 붙은 두 번째 저장소는 문서관리(DM) 플러그인이다. 아카이브가 아니다
  - probe v2 가 그 저장소의 폴더를 조회하다가, v3 는 저장소 목록(종류 조회)만 돌다가 DM 오류와 함께 Outlook 이 죽었다
  - 그래서 `ns.Stores`, `GetRootFolder`, 저장소 속성 조회를 쓰지 않는다. `GetDefaultFolder` 에서만 출발한다
    (`tests/test_probe.py` 가 outlook.py 에 금지 호출이 없는지 검사한다)
  - 아카이브 메일은 Outlook 에서 보이지 않는다 → 수집 대상은 기본 사서함뿐. 1개월 안에 놓치지 않는 것이 핵심
- probe 는 표본만 읽는다 (폴더당 30행 + 가장 오래된 1행). v2 에서 전체 목록을 읽게 했더니
  오래 걸려 사용자가 probe 창을 강제로 닫았고, 그 직후 Outlook 이 죽었다
- **COM 클라이언트는 강제 종료(창 닫기)하지 않게 한다.** Ctrl+C 를 받으면 진행 중 호출이 끝난 뒤
  객체를 정리하고 정상 종료한다. 수집기는 오래 돌 수 있으므로 진행 표시·중단·재개를 기본으로 갖춘다

## 수집(동기화) 방식 (2026-10-14 결정)
- **동기화는 mailer 앱(PySide6 UI)에서 사용자가 수동으로 실행**한다. 자동 실행·스케줄러 없음
  - COM 은 UI 스레드가 아닌 작업 스레드에서 호출한다 (pythoncom.CoInitialize 필요)
- **전체 대조는 하지 않는다.** 증분만:
  - DB 에 마지막 동기화 일시, 메일별 Message-ID·수신 일시·수집 당시 폴더를 저장
  - 폴더마다 Table 을 수신 일시 내림차순으로 읽다가 (마지막 동기화 일시 - 여유) 보다 오래되면 멈춤
  - Message-ID 가 DB 에 없는 것만 열어서 .msg 저장
- 따라서 mailer 의 폴더는 **수집 당시 폴더의 스냅샷**이다. 수집 후 Outlook 에서 폴더를 옮겨도 갱신되지 않는다
  (사용자 확인함). 오래된 메일이 나중에 대상 폴더로 들어온 경우(예: 지운 편지함에서 복원)도 수집되지 않는다
- 동기화 간격이 보존 기간(약 1개월)을 넘으면 그 사이 메일은 잃는다

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
