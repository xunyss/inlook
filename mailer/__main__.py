"""python -m mailer <명령>"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import CODES

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"


def _outlook_source():
    try:
        from .outlook import OutlookSource
    except ImportError as exc:
        code = CODES["E100"]
        raise RuntimeError(f"[{code.code}] {code.title} ({exc}) — {code.action}") from exc
    return OutlookSource()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="mailer")
    sub = parser.add_subparsers(dest="command", required=True)

    probe = sub.add_parser("probe", help="VDI 자가진단")
    probe.add_argument("archive_root", nargs="?", type=Path, help="보존 경로 (쓰기 가능 여부 확인)")
    probe.add_argument("--fake", action="store_true", help="Outlook 대신 합성 데이터 사용 (Mac 개발용)")
    probe.add_argument("--deep", type=int, default=10, help="상세 읽기 시간을 잴 메일 수 (기본 10)")
    probe.add_argument("--max-rows", type=int, default=5000, help="폴더당 목록을 읽을 최대 행 수 (기본 5000)")

    args = parser.parse_args(argv)

    if args.command == "probe":
        from .probe import run_probe

        if args.fake:
            from .fake import FakeMailSource

            make_source = FakeMailSource
        else:
            make_source = _outlook_source
        return run_probe(
            make_source, deep=args.deep, max_rows=args.max_rows, archive_root=args.archive_root, log_dir=LOG_DIR
        )
    return 2


if __name__ == "__main__":
    sys.exit(main())
