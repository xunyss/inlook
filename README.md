# inlook

Outlook(Exchange) 메일을 보존 정책으로 사라지기 전에 `.msg` 로 모아 두고, 검색해서 다시 읽는 데스크톱 도구.

Outlook 이 앞을 내다보는(out-look) 도구라면, inlook 은 지난 메일을 돌아보는 도구다.

- **보존기 + 검색 뷰어**: 메일을 보내거나 받는 클라이언트가 아니다. Outlook 옆에서 사본을 남긴다
- **수집**: Python + pywin32 로 Outlook COM 자동화 (Windows 전용). 기본 Exchange 사서함만 읽는다
- **뷰어**: PySide6 데스크톱 앱. 브라우저나 웹 서버를 쓰지 않는다
- **데이터는 그 PC 안에만** 저장된다 (SQLite DB + `.msg` 파일)

## 주요 기능

| 기능 | 설명 |
|---|---|
| 동기화 | Outlook 의 새 메일을 `.msg` 원본 + DB 로 저장. 증분, 오래된 메일부터, 중단·재개 가능 |
| 검색 | 제목·보낸 사람·받는 사람·본문·첨부 파일명. 띄어쓰기로 여러 단어(AND) |
| 라벨 | 처음 수집할 때 Outlook 폴더 경로가 라벨이 된다. 이후 inlook 안에서 바꾸고 정리 |
| 미리보기 | HTML 본문, 본문 속 이미지, 어두운 테마 색 보정(☀ 로 원본 색), `.msg` 를 Outlook 으로 열기 |
| 삭제 | `.msg` 와 DB 기록을 함께 삭제. 다음 동기화에서 다시 가져오지 않는다 |
| 가져오기 | 예전에 '다른 이름으로 저장' 해 둔 `.msg` 폴더를 Outlook 없이 가져오기 |

## 설치 (Windows)

Python 3.12 이상과 Outlook 이 있는 Windows PC. [uv](https://docs.astral.sh/uv/) 가 있으면 uv 로, 없으면 venv + pip 으로 설치한다.

```
git clone https://github.com/xunyss/inlook.git
cd inlook
setup.bat
```

`setup.bat` 은 처음 한 번, 그리고 `git pull` 로 의존성이 바뀌었을 때 다시 실행한다.
보안 프로그램 때문에 설치 중 "액세스가 거부되었습니다" 가 나면 자동으로 다시 시도한다 (최대 5번).

## 사용법

### 앱

```
app.bat                 # 기본 보존 경로 D:\mailbox
app.bat E:\other        # 다른 보존 경로
```

- 화면: **라벨 트리 | 메일 목록 | 미리보기**
- 라벨을 누르면 그 라벨이 붙은 메일만 보인다 (하위 라벨 메일은 제외). 숫자도 같은 기준이고, 하위 포함 건수는 툴팁
- 메일 우클릭
  - **라벨 변경 ▸** — 라벨 트리에서 고르면 바로 바뀐다 (기존 라벨은 떼어짐). 맨 아래 **새 라벨…**
  - **.msg 열기** — Outlook 으로 원본 열기 (더블클릭도 같음)
  - **삭제** (Del) — 확인 후 `.msg` 와 DB 기록 삭제. 되돌릴 수 없다
- 라벨 우클릭: **이름 바꾸기**, **라벨 삭제** (메일은 남음). 하위 라벨까지 함께 적용
- 오른쪽 위 **[동기화]** — `collect.bat` 과 같은 일. 동기화 창은 끝날 때까지 닫히지 않는다 (멈출 때는 [중단])
- 마지막 동기화가 21일 넘게 지나면 오른쪽 위 표시가 빨갛게 바뀐다

### Outlook 리본에 실행 버튼 달기 (선택)

Outlook VBA 매크로로 `app.bat` 을 실행하는 버튼을 홈 탭에 단다. 처음 한 번만 하면 된다.

1. Outlook 에서 `Alt+F11` → 왼쪽 `Project1` 우클릭 → **삽입 → 모듈** → 아래 코드를 붙여 넣고 저장.
   경로는 inlook 을 clone 한 폴더로 바꾼다

   ```vba
   Public Sub inlook()
       Shell "cmd /c ""D:\inlook\app.bat""", vbHide
   End Sub
   ```

2. **파일 → 옵션 → 리본 사용자 지정**
   - 오른쪽 목록에서 **홈(메일)** 을 고르고 **새 그룹**. 맨 아래에 두면 홈 탭 오른쪽 끝에 붙는다
   - 왼쪽 위 **명령 선택** 드롭다운을 **매크로** 로 바꾸고 `Project1.inlook` 을 새 그룹에 **추가 >>**
   - **이름 바꾸기** 로 표시 이름(inlook)과 아이콘을 고른다. 그룹 이름도 같은 방법으로 바꾼다

- **매크로** 목록이 비어 있으면: 모듈을 저장하지 않았거나 `Public Sub` 가 아닌 경우,
  또는 매크로가 꺼진 경우 (**보안 센터 → 보안 센터 설정 → 매크로 설정**)
- 매크로 설정이 '알림 표시' 면 Outlook 을 켤 때마다 매크로 사용 여부를 묻는다
- 버튼을 눌러도 앱이 안 뜨면 `app.bat` 을 직접 실행해 오류를 확인한다 (버튼은 실행 창을 숨긴다)
- 앱은 중복 실행을 막지 않는다. 버튼을 두 번 누르면 창이 두 개 뜨니, 동기화는 한쪽에서만 한다

### 동기화 (명령줄)

```
collect.bat             # 새 메일을 D:\mailbox 에 저장
collect.bat --dry-run   # 저장하지 않고 새 메일이 몇 건인지만
collect.bat --full      # 마지막 동기화와 상관없이 전체 목록을 다시 확인
collect.bat E:\other    # 다른 보존 경로
```

- 실행 창을 **닫지 말 것** (Outlook 이 죽을 수 있다). 멈출 때는 **Ctrl+C** — 저장된 메일은 남고, 다시 실행하면 이어서 한다

### 예전 .msg 폴더 가져오기

```
import.bat D:\메일백업 --dry-run   # 몇 건이 새 메일인지 확인만
import.bat D:\메일백업             # 가져오기 (앱은 닫고 실행)
```

- 폴더 바로 아래 파일은 `받은 편지함`, 하위 폴더 `X` 안의 파일은 `받은 편지함/X` 라벨이 붙는다
  (`--parent-label ""` 이면 폴더 이름만)
- 파일 이름이 아니라 `.msg` 안의 **받은 시각**(없으면 보낸 시각)을 쓴다
- 이미 있는 메일(같은 Message-ID)과 inlook 에서 삭제한 메일은 건너뛴다. 원본 폴더는 읽기만 한다

### DB 크기 확인·줄이기

```
stats.bat               # 어떤 정보가 DB 공간을 차지하는지 (크기·건수만)
compact.bat --dry-run   # 줄일 수 있는 양 확인
compact.bat             # HTML 본문을 DB 에서 비우고 정리 (앱은 닫고 실행)
```

- HTML 본문의 원본은 `.msg` 안에 있어서 미리보기는 `.msg` 에서 꺼낸다. DB 에는 검색용 텍스트만 둔다
- `.msg` 에서 HTML 을 꺼낼 수 없는 메일은 DB 의 HTML 을 그대로 둔다

### 자가진단

```
probe.bat D:\mailbox    # Outlook 연결, 목록 읽기 방식, 저장 가능 여부를 표본으로 확인
```

## 동작 방식

### 동기화

- 기본 사서함의 메일 폴더를 돈다. **지운 편지함·임시 보관함·보낼 편지함은 제외**, 보낸 편지함·정크 메일은 포함
- 폴더마다 받은 시각 최신순으로 읽다가 **(마지막 정상 동기화 − 3일)** 보다 오래된 메일이 나오면 멈춘다.
  첫 동기화는 전체를 읽는다
- DB 에 없는 메일만 열어서 저장한다. 같은 메일 판단은 Internet Message-ID (없으면 SearchKey)
- 메일 한 건마다 바로 기록한다. 실패나 중단이 있었던 동기화는 기준 시각으로 쓰지 않아, 다음 실행이 빠진 것을 채운다
- **동기화 간격이 메일함 보존 기간보다 길면 그 사이 메일은 잃는다.** 며칠에 한 번은 동기화할 것

### 라벨

- inlook 에는 폴더가 없고 **라벨**만 있다. 메일은 받은 시각 기준으로 저장된다
- Outlook 폴더 경로는 **처음 수집할 때만** 라벨로 붙는다. 이후 Outlook 에서 폴더를 옮겨도 inlook 라벨은 그대로다.
  그래서 Outlook 에서 폴더 정리를 먼저 하고 동기화하는 편이 좋다
- 폴더 이름 안의 `/` 는 라벨에서 `∕` 로 바뀐다 (`/` 는 라벨 계층 구분자)

### 저장 구조

```
D:\mailbox\
  inlook.db                                   메일 정보·검색용 본문 텍스트·라벨·동기화 기록 (SQLite)
  msg\YYYY\MM\YYYYMMDD-HHMMSS_<hash>.msg      원본 (olMSGUnicode)
```

- 파일 이름 = 받은 시각 + 메일 키의 해시. 같은 메일은 언제 저장해도 같은 이름이다
- 제목·폴더명을 파일 이름에 넣지 않아 Windows 260자 경로 제한이나 금지 문자 문제가 없다
- HTML 본문·본문 속 이미지·첨부는 `.msg` 안에만 있고 DB 에는 넣지 않는다 (미리보기 때 `.msg` 에서 꺼낸다)

## 문제 해결

| 증상 | 조치 |
|---|---|
| 앱이 뜨지 않음 | cmd 에서 `.venv\Scripts\python.exe -m inlook app` 으로 실행해 오류 확인. 오류는 `logs\app-error.txt` 에도 남는다 |
| `[E코드]` 오류 | 화면에 원인과 조치가 함께 나온다. 로그는 `logs\` 폴더 |
| `git pull` 후 실행 안 됨 | `setup.bat` 다시 실행 |
| 메일이 0건으로 보임 | 보존 경로(기본 `D:\mailbox`)가 맞는지, 그 안에 `inlook.db` 가 있는지 확인 |

동기화·가져오기·진단 출력에는 메일 제목·사람 이름·폴더명이 나오지 않는다 (`folder#NN`, `file#N` 으로만 표시).
화면을 그대로 공유해도 된다.

## 개발 (Mac 등, Outlook 없이)

```bash
uv sync --python 3.12
uv run pytest
uv run python -m inlook app ./tmp-archive --fake   # [동기화] 가 합성 데이터를 넣는다
uv run python -m inlook probe --fake
```

- Outlook COM 은 `MailSource` 인터페이스 뒤에 격리돼 있다. Windows 에서는 `OutlookSource`, 그 밖에서는 합성 데이터 `FakeMailSource`
- 테스트 데이터는 모두 합성이다. 테스트용 `.msg` 는 `tests/cfb.py` 가 만든다

| 파일 | 역할 |
|---|---|
| `inlook/source.py` | 메일 소스 인터페이스와 타입 |
| `inlook/outlook.py` | Outlook COM 구현 (Windows 전용, 얇게) |
| `inlook/fake.py` | 합성 데이터 소스 |
| `inlook/collect.py` | 동기화 엔진 |
| `inlook/importer.py` | `.msg` 폴더 가져오기 |
| `inlook/store.py` | DB 스키마·저장·삭제 |
| `inlook/library.py` | 뷰어용 조회·라벨 편집 |
| `inlook/msgfile.py` | `.msg` 파일 읽기 (olefile) |
| `inlook/rtfhtml.py` | 압축 RTF 안의 HTML 꺼내기 (회의 요청 등) |
| `inlook/htmlcolor.py` | 어두운 테마 본문 색 보정 |
| `inlook/stats.py`, `compact.py` | DB 크기 진단·줄이기 |
| `inlook/ui.py` | PySide6 앱 |
| `inlook/probe.py` | 자가진단 |
