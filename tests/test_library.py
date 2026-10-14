from __future__ import annotations

from datetime import datetime

import pytest

from mailer.collect import run_collect
from mailer.fake import FakeMailSource
from mailer.library import UNLABELED, Library

NOW = datetime(2026, 10, 14, 9, 30)


@pytest.fixture
def lib(tmp_path):
    run_collect(lambda: FakeMailSource(NOW), tmp_path)
    library = Library(tmp_path)
    yield library
    library.close()


def test_search_newest_first_and_terms_are_anded(lib):
    rows = lib.search()
    assert len(rows) == lib.count()
    assert [r.received for r in rows] == sorted((r.received for r in rows), reverse=True)

    assert {r.subject for r in lib.search("회의록")} >= {"[회의록] 0차 회의"}
    both = lib.search("프로젝트A 진행")
    assert both and all("프로젝트A" in r.subject and "진행" in r.subject for r in both)


def test_search_matches_body_sender_and_attachment_names(lib):
    assert lib.search("홍길동")
    assert lib.search("chulsoo.kim")
    assert all(r.has_attachments for r in lib.search("보고서.docx"))
    assert lib.search("번호: 7")


def test_like_wildcards_are_literal(lib):
    assert lib.search("%") == []
    assert lib.search("_") == []


def test_label_filter_includes_children(lib):
    parent = lib.search(label="받은 편지함/프로젝트A")
    child = lib.search(label="받은 편지함/프로젝트A/회의록")
    assert len(parent) == 15 and len(child) == 5
    assert lib.search(label=UNLABELED) == []


def test_add_remove_label_and_unlabeled(lib):
    ids = [r.id for r in lib.search("News Letter")][:3]
    assert lib.add_label(ids, " 읽을거리 / 나중에 ") == "읽을거리/나중에"
    assert {r.id for r in lib.search(label="읽을거리")} == set(ids)

    for name in lib.get(ids[0]).labels:
        lib.remove_label([ids[0]], name)
    assert [r.id for r in lib.search(label=UNLABELED)] == [ids[0]]
    assert lib.unlabeled_count() == 1


def test_rename_label_moves_children_and_merges(lib):
    lib.rename_label("받은 편지함/프로젝트A", "고객/A")
    counts = lib.label_counts()
    assert "받은 편지함/프로젝트A" not in counts
    assert counts["고객/A"] == 10 and counts["고객/A/회의록"] == 5
    assert len(lib.search(label="고객")) == 15


def test_delete_label_keeps_mail(lib):
    before = lib.count()
    lib.delete_label("받은 편지함/뉴스레터")
    assert lib.count() == before
    assert "받은 편지함/뉴스레터" not in lib.label_counts()
    assert len(lib.search(label=UNLABELED)) == 30


def test_get_returns_detail_with_msg_path(lib):
    row = lib.search("배달되지 않음")[0]
    mail = lib.get(row.id)
    assert mail.msg_path.is_file()
    assert mail.labels == ["받은 편지함"]


def test_tree_counts_match_list_size(lib):
    totals = lib.tree_counts()
    for path in ("받은 편지함", "받은 편지함/프로젝트A", "보낸 편지함"):
        assert totals[path] == len(lib.search(label=path))
    assert totals["받은 편지함"] + totals["보낸 편지함"] + totals["보관"] == lib.count()
