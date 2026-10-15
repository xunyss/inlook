# mailer

Outlook(Exchange) 메일을 보존 정책으로 사라지기 전에 `.msg` 로 저장하고, 검색해서 다시 읽기 위한 도구.

- 수집: Python + pywin32 로 Outlook COM 자동화 (Windows 전용)
- 뷰어: PySide6 데스크톱 UI (`app.bat`)

## Windows 에서 사용

```
git clone https://github.com/xunyss/mailer.git
cd mailer
setup.bat              # 최초 1회, 그리고 의존성이 바뀌었을 때
probe.bat H:\mail      # 자가진단. 인자는 보존할 폴더 (생략 가능). 옵션: --rows 30 --deep 5
collect.bat            # 동기화. 새 메일을 D:\mailbox 에 .msg + DB 로 저장
collect.bat --dry-run  # 저장하지 않고 새 메일이 몇 건인지만 확인
collect.bat E:\other   # 다른 보존 경로. --full: 마지막 동기화와 상관없이 전체 목록 재확인
```

- 저장 위치: `D:\mailbox\mailer.db`, `D:\mailbox\msg\YYYY\MM\YYYYMMDD-HHMMSS_<hash>.msg`
- 오래된 메일부터 저장한다. 중간에 Ctrl+C 로 멈춰도 저장된 건 남고, 다시 실행하면 이어서 한다.
- 실행 창을 닫지 말 것 (Outlook 이 죽을 수 있다). 멈출 때는 Ctrl+C.

```
app.bat                # 뷰어. 검색, 라벨 변경/이름 바꾸기, 메일 삭제, 미리보기, .msg 열기, [동기화] 버튼
```

- 동기화 버튼은 `collect.bat` 과 같은 일을 한다. 동기화 창은 끝날 때까지 닫히지 않는다 (멈출 때는 [중단]).
- 마지막 동기화가 21일 넘게 지났으면 오른쪽 위 표시가 빨갛게 바뀐다.
- 앱이 안 뜨면 cmd 에서 `.venv\Scripts\python.exe -m mailer app` 으로 실행해 오류를 확인한다.

`probe` 출력에는 메일 내용·이름·폴더명이 나오지 않는다. 폴더는 `folder#NN` 으로만 표시된다.

## 개발 (Mac 등, Outlook 없이)

```bash
uv sync --python 3.12
uv run python -m mailer probe --fake
uv run python -m mailer app ./tmp-archive --fake   # 동기화 버튼이 합성 데이터를 넣는다
uv run pytest
```
