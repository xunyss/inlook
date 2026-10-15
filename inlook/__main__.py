"""python -m inlook <명령>"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .collect import DEFAULT_MARGIN_DAYS
from .errors import CODES

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
DEFAULT_ARCHIVE_ROOT = Path("D:/mailbox")


def _outlook_source():
    try:
        from .outlook import OutlookSource
    except ImportError as exc:
        code = CODES["E100"]
        raise RuntimeError(f"[{code.code}] {code.title} ({exc}) — {code.action}") from exc
    return OutlookSource()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="inlook")
    sub = parser.add_subparsers(dest="command", required=True)

    probe = sub.add_parser("probe", help="VDI 자가진단")
    probe.add_argument("archive_root", nargs="?", type=Path, help="보존 경로 (쓰기 가능 여부 확인)")
    probe.add_argument("--fake", action="store_true", help="Outlook 대신 합성 데이터 사용 (Mac 개발용)")
    probe.add_argument("--rows", type=int, default=30, help="폴더당 읽어볼 목록 행 수 (기본 30)")
    probe.add_argument("--deep", type=int, default=5, help="열어서 시간을 잴 메일 수 (기본 5)")

    collect = sub.add_parser("collect", help="동기화: 새 메일을 .msg + DB 로 저장")
    collect.add_argument("archive_root", nargs="?", type=Path, default=DEFAULT_ARCHIVE_ROOT,
                         help=f"보존 경로 (기본 {DEFAULT_ARCHIVE_ROOT})")
    collect.add_argument("--fake", action="store_true", help="Outlook 대신 합성 데이터 사용 (Mac 개발용)")
    collect.add_argument("--dry-run", action="store_true", help="목록만 읽고 저장하지 않음")
    collect.add_argument("--full", action="store_true", help="마지막 동기화 시각과 상관없이 전체 목록을 다시 확인")
    collect.add_argument("--margin-days", type=int, default=DEFAULT_MARGIN_DAYS,
                         help=f"마지막 동기화보다 며칠 더 거슬러 볼지 (기본 {DEFAULT_MARGIN_DAYS})")

    app = sub.add_parser("app", help="뷰어 (검색·라벨·동기화)")
    app.add_argument("archive_root", nargs="?", type=Path, default=DEFAULT_ARCHIVE_ROOT,
                     help=f"보존 경로 (기본 {DEFAULT_ARCHIVE_ROOT})")
    app.add_argument("--fake", action="store_true", help="동기화에 합성 데이터 사용 (Mac 개발용)")

    args = parser.parse_args(argv)

    if args.fake:
        from .fake import FakeMailSource

        make_source = FakeMailSource
    else:
        make_source = _outlook_source

    if args.command == "app":
        from .ui import run_app

        return run_app(args.archive_root, make_source, LOG_DIR)
    if args.command == "collect":
        from .collect import run_collect

        return run_collect(
            make_source,
            args.archive_root,
            full=args.full,
            dry_run=args.dry_run,
            margin_days=args.margin_days,
            log_dir=LOG_DIR,
        )
    if args.command == "probe":
        from .probe import run_probe

        return run_probe(
            make_source,
            rows=args.rows,
            deep=args.deep,
            archive_root=args.archive_root,
            log_dir=LOG_DIR,
        )
    return 2


if __name__ == "__main__":
    sys.exit(main())
