from datetime import datetime

from mailer.fake import FakeMailSource
from mailer.probe import run_probe

NOW = datetime(2026, 10, 13, 12, 0)


def test_probe_passes_with_fake(tmp_path, capsys):
    code = run_probe(
        lambda: FakeMailSource(NOW), rows=100, other_stores=True,
        archive_root=tmp_path / "archive", log_dir=tmp_path / "logs",
    )
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
    run_probe(lambda: source, other_stores=True, log_dir=tmp_path)
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


def test_probe_skips_other_stores_by_default(tmp_path, capsys):
    run_probe(lambda: FakeMailSource(NOW), log_dir=tmp_path)
    out = capsys.readouterr().out
    assert "store#2" in out and "폴더 조회 생략" in out
    assert " s2 " not in out


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
