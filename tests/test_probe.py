from datetime import datetime

from mailer.fake import FakeMailSource
from mailer.probe import run_probe

NOW = datetime(2026, 10, 13, 12, 0)


def test_probe_passes_with_fake(tmp_path, capsys):
    code = run_probe(
        lambda: FakeMailSource(NOW), rows=100,
        archive_root=tmp_path / "archive", log_dir=tmp_path / "logs",
    )
    out = capsys.readouterr().out
    assert code == 0
    assert "FAIL 0" in out
    assert "[E106]" in out  # Message-ID 없는 항목 경고
    assert "[E107]" in out  # stub 경고
    assert list((tmp_path / "logs").glob("probe-*.txt"))
    assert not list((tmp_path / "archive").iterdir())  # 테스트 저장 파일은 지운다


def test_probe_output_has_no_names(tmp_path, capsys):
    source = FakeMailSource(NOW)
    run_probe(lambda: source, log_dir=tmp_path)
    out = capsys.readouterr().out
    secrets: set[str] = set()
    for f in source.folders():
        secrets.update(f.path)
        for h in source.headers(f):
            record = source.open(h)
            secrets.update({h.subject, h.sender_name, record.sender_email})
    leaked = [s for s in secrets if len(s) >= 2 and s in out]
    assert not leaked


def test_probe_reports_connection_failure(tmp_path, capsys):
    def broken():
        raise RuntimeError("no outlook")

    assert run_probe(broken, log_dir=tmp_path) == 1
    assert "[E101]" in capsys.readouterr().out


def test_probe_reports_save_failure(tmp_path, capsys):
    assert run_probe(lambda: FakeMailSource(NOW, fail_save=True), log_dir=tmp_path) == 1
    assert "[E109]" in capsys.readouterr().out


def test_source_has_no_store_access():
    from mailer import source

    assert not hasattr(source, "StoreRef")
    assert not hasattr(FakeMailSource, "stores")


def test_probe_reads_only_sample_rows(tmp_path):
    class Counting(FakeMailSource):
        calls: list[int | None] = []

        def headers(self, folder, since=None, limit=None, oldest_first=False):
            self.calls.append(limit if since is None else -1)
            return super().headers(folder, since=since, limit=limit, oldest_first=oldest_first)

    source = Counting(NOW)
    run_probe(lambda: source, rows=7, log_dir=tmp_path)
    full_scans = [c for c in source.calls if c is None]
    assert not full_scans
    assert max(c for c in source.calls) <= 7


def test_log_written_line_by_line(tmp_path):
    def broken():
        raise RuntimeError("no outlook")

    run_probe(broken, log_dir=tmp_path)
    text = next(tmp_path.glob("probe-*.txt")).read_text(encoding="utf-8")
    assert "· Outlook 연결 ..." in text and "[E101]" in text


def test_ctrl_c_stops_cleanly(tmp_path, capsys):
    class Interrupted(FakeMailSource):
        def measure(self, header):
            raise KeyboardInterrupt

    assert run_probe(lambda: Interrupted(NOW), log_dir=tmp_path) == 130
    out = capsys.readouterr().out
    assert "[STOP]" in out and "결과:" in out


def test_outlook_source_never_touches_other_stores():
    # outlook.py 는 Mac 에서 import 할 수 없으므로 코드 문자열로 확인한다
    from pathlib import Path

    code = Path(__file__).parent.parent / "mailer" / "outlook.py"
    lines = [line for line in code.read_text(encoding="utf-8").splitlines() if not line.lstrip().startswith("#")]
    body = "\n".join(lines).split('"""', 2)[-1]  # 모듈 docstring 제외
    for banned in (".Stores", "GetRootFolder", "ExchangeStoreType", "IsDataFileStore"):
        assert banned not in body, banned
