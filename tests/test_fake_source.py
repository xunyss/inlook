from datetime import datetime, timedelta

from mailer.fake import FakeMailSource

NOW = datetime(2026, 10, 13, 12, 0)


def _folder(source, name, store=None):
    return next(f for f in source.folders(store) if f.path[-1] == name)


def test_headers_newest_first_and_limit():
    source = FakeMailSource(NOW)
    rows = list(source.headers(_folder(source, "받은 편지함"), limit=10))
    assert len(rows) == 10
    times = [h.received for h in rows]
    assert times == sorted(times, reverse=True)


def test_headers_since_stops_at_boundary():
    source = FakeMailSource(NOW)
    inbox = _folder(source, "받은 편지함")
    since = NOW - timedelta(days=10)
    rows = list(source.headers(inbox, since=since))
    assert rows and all(h.received >= since for h in rows)
    assert len(rows) == sum(1 for h in source.headers(inbox) if h.received >= since)


def test_item_count_matches_headers():
    source = FakeMailSource(NOW)
    for store in source.stores():
        for folder in source.folders(store):
            assert folder.item_count == len(list(source.headers(folder)))


def test_open_returns_full_record():
    source = FakeMailSource(NOW)
    header = next(source.headers(_folder(source, "회의록")))
    record = source.open(header)
    assert record.header == header
    assert record.html_body and record.attachments


def test_second_store_has_older_mail():
    source = FakeMailSource(NOW)
    archive = next(s for s in source.stores() if not s.is_default)
    rows = [h for f in source.folders(archive) for h in source.headers(f)]
    assert rows and all(h.received < NOW - timedelta(days=30) for h in rows)


def test_edge_cases_present():
    source = FakeMailSource(NOW)
    rows = [h for s in source.stores() for f in source.folders(s) for h in source.headers(f)]
    assert any(h.is_stub for h in rows)
    assert any(h.message_id is None for h in rows)
    assert any(h.subject == "" for h in rows)
    assert any(len(h.subject) > 200 for h in rows)
    assert not any(f.is_mail for f in source.folders() if f.path == ("일정",))
