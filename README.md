# mailer

Outlook(Exchange) 메일을 보존 정책으로 사라지기 전에 `.msg` 로 저장하고, 검색해서 다시 읽기 위한 도구.

- 수집: Python + pywin32 로 Outlook COM 자동화 (Windows 전용)
- 뷰어: PySide6 데스크톱 UI (예정)

## Windows 에서 사용

```
git clone https://github.com/xunyss/mailer.git
cd mailer
setup.bat              # 최초 1회, 그리고 의존성이 바뀌었을 때
probe.bat H:\mail      # 자가진단. 인자는 보존할 폴더 (생략 가능). 옵션: --rows 30 --deep 5 --stores
```

`probe` 출력에는 메일 내용·이름·폴더명이 나오지 않는다. 폴더는 `folder#NN` 으로만 표시된다.

## 개발 (Mac 등, Outlook 없이)

```bash
uv sync --python 3.12
uv run python -m mailer probe --fake
uv run pytest
```
