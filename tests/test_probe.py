from datetime import datetime

from mailer.fake import FakeMailSource
from mailer.probe import run_probe

NOW = datetime(2026, 10, 13, 12, 0)


def test_probe_passes_with_fake(tmp_path, capsys):
    code = run_probe(lambda: FakeMailSource(NOW), archive_root=tmp_path / "archive", log_dir=tmp_path / "logs")
    out = capsys.readouterr().out
    assert code == 0
    assert "FAIL 0" in out
    assert "store#2" in out
    assert "[E106]" in out  # Message-ID 없는 항목 경고
    assert "[E107]" in out  # stub 경고
    assert list((tmp_path / "logs").glob("probe-*.txt"))
    assert not list((tmp_path / "archive").iterdir())  # 테스트 저장 파일은 지운다


def test_probe_output_has_no_names(tmp_path, capsys):
    source = FakeMailSource(NOW)
    run_probe(lambda: source, log_dir=tmp_path)
    out = capsys.readouterr().out
    secrets: set[str] = set()
    for store in source.stores():
        for f in source.folders(store):
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
