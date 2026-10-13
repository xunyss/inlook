from datetime import datetime, timedelta

from mailer.fake import FakeMailSource

NOW = datetime(2026, 10, 13, 12, 0)


def _folder(source, name):
    return next(f for f in source.folders() if f.path[-1] == name)


def test_items_newest_first_and_limit():
    source = FakeMailSource(NOW)
    records = list(source.items(_folder(source, "받은 편지함"), limit=10))
    assert len(records) == 10
    times = [r.received for r in records]
    assert times == sorted(times, reverse=True)


def test_items_since_stops_at_boundary():
    source = FakeMailSource(NOW)
    inbox = _folder(source, "받은 편지함")
    since = NOW - timedelta(days=30)
    records = list(source.items(inbox, since=since))
    assert records and all(r.received >= since for r in records)
    everything = list(source.items(inbox))
    assert len(records) == sum(1 for r in everything if r.received >= since)


def test_item_count_matches_items():
    source = FakeMailSource(NOW)
    for folder in source.folders():
        assert folder.item_count == len(list(source.items(folder)))


def test_edge_cases_present():
    source = FakeMailSource(NOW)
    records = [r for f in source.folders() for r in source.items(f)]
    assert any(r.is_stub for r in records)
    assert any(r.message_id is None for r in records)
    assert any(r.subject == "" for r in records)
    assert any(len(r.subject) > 200 for r in records)
    assert not any(f.is_mail for f in source.folders() if f.path == ("일정",))
