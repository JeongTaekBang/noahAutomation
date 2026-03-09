#!/usr/bin/env python
"""
Order Book 월별 마감 (스냅샷)
==============================

월별 수주잔고를 스냅샷으로 저장하여 Start를 고정하고,
이후 소급 변경분은 Variance로 자동 감지합니다.

사용법:
    python close_period.py 2026-01                    # 1월 마감
    python close_period.py 2026-02 --note "정기 마감"   # 노트 포함
    python close_period.py --undo 2026-02              # 마감 취소
    python close_period.py --list                      # 마감 현황
    python close_period.py --status                    # 현재 상태
"""

from __future__ import annotations

import argparse
import sys

from po_generator.config import DB_FILE
from po_generator.logging_config import setup_logging
from po_generator.snapshot import SnapshotEngine


def _fmt_amt(val: float) -> str:
    """금액 포맷 (백만 단위, 소수점 1자리)"""
    if abs(val) < 1000:
        return f"{val:,.0f}"
    return f"{val / 1_000_000:,.1f}M"


def print_list(snapshots: list[dict]) -> None:
    """마감 현황 출력"""
    print("\nOrder Book 마감 현황")
    print("=" * 96)

    if not snapshots:
        print("마감된 Period가 없습니다.")
        return

    # 활성 스냅샷만 필터
    active = [s for s in snapshots if s['is_active']]

    print(f"{'Period':<10} {'건수':>5} {'Start':>12} {'Input':>12} {'Output':>12} {'Variance':>12} {'Ending':>12} {'마감일시':<16}")
    print("-" * 96)

    prev_ending = None
    for s in snapshots:
        start = s['total_start'] or 0
        inp = s['total_input'] or 0
        out = s['total_output'] or 0
        var = s['total_variance'] or 0
        ending = s['total_ending'] or 0

        # 정합성 체크: 전월 Ending == 당월 Start
        check = ""
        if prev_ending is not None and s['is_active']:
            diff = abs(prev_ending - start)
            if diff > 1:
                check = f" (!차이 {_fmt_amt(diff)})"

        status_mark = "" if s['is_active'] else " [취소]"
        var_str = _fmt_amt(var) if abs(var) > 0.5 else "-"
        closed = (s['closed_at'] or '')[:16].replace('T', ' ')

        print(
            f"{s['period']:<10} {s['row_count']:>5} "
            f"{_fmt_amt(start):>12} {_fmt_amt(inp):>12} {_fmt_amt(out):>12} "
            f"{var_str:>12} {_fmt_amt(ending):>12} "
            f"{closed:<16}{status_mark}"
        )

        if check:
            print(f"{'':>10} ** Start != 전월 Ending{check}")

        if s['is_active']:
            prev_ending = ending

    print("-" * 96)

    # 합계 (활성만)
    if active:
        total_input = sum(s['total_input'] or 0 for s in active)
        total_output = sum(s['total_output'] or 0 for s in active)
        total_var = sum(s['total_variance'] or 0 for s in active)
        last_ending = active[-1]['total_ending'] or 0
        print(
            f"{'합계':<10} {'':>5} "
            f"{'':>12} {_fmt_amt(total_input):>12} {_fmt_amt(total_output):>12} "
            f"{_fmt_amt(total_var) if abs(total_var) > 0.5 else '-':>12} {_fmt_amt(last_ending):>12}"
        )

    # 비고 표시 (있는 것만)
    notes = [(s['period'], s['note']) for s in snapshots if s.get('note')]
    if notes:
        print()
        for period, note in notes:
            print(f"  {period}: {note}")


def print_status(status: dict) -> None:
    """현재 상태 출력"""
    print("\nOrder Book 스냅샷 상태")
    print("=" * 40)
    last = status['last_closed_period'] or '(없음)'
    nxt = status['next_period'] or '(첫 마감 필요)'
    print(f"마지막 마감 Period: {last}")
    print(f"다음 마감 가능:     {nxt}")
    print(f"활성 스냅샷 수:     {status['active_snapshots']}")
    print(f"총 스냅샷 행:       {status['total_snapshot_rows']}")
    print(f"DB 파일:            {DB_FILE}")


def create_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog='close_period',
        description='Order Book 월별 마감 - 스냅샷 기반 Variance 추적',
        epilog='예시: python close_period.py 2026-01 --note "1월 정기 마감"',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'period',
        nargs='?',
        metavar='YYYY-MM',
        help='마감할 Period (예: 2026-01)',
    )

    parser.add_argument(
        '--note',
        default='',
        help='마감 비고',
    )

    parser.add_argument(
        '--undo',
        metavar='YYYY-MM',
        help='마감 취소할 Period (최신 마감만 가능)',
    )

    parser.add_argument(
        '--list',
        action='store_true',
        help='마감 현황 조회',
    )

    parser.add_argument(
        '--status',
        action='store_true',
        help='현재 상태 요약',
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )

    return parser


def main() -> int:
    parser = create_argument_parser()
    args = parser.parse_args()

    setup_logging(verbose=args.verbose)

    if not DB_FILE.exists():
        print(f"[오류] DB 파일이 없습니다: {DB_FILE}")
        print("sync_db.py를 먼저 실행하세요.")
        return 1

    engine = SnapshotEngine()

    # --list: 마감 현황
    if args.list:
        snapshots = engine.list_snapshots()
        print_list(snapshots)
        return 0

    # --status: 현재 상태
    if args.status:
        status = engine.get_status()
        print_status(status)
        return 0

    # --undo: 마감 취소
    if args.undo:
        result = engine.undo_snapshot(args.undo)
        if result.success:
            print(f"\n[완료] {result.message}")
            return 0
        else:
            print(f"\n[오류] {result.message}")
            return 1

    # 마감 실행
    if not args.period:
        parser.print_help()
        return 1

    print(f"\nOrder Book '{args.period}' 마감 시작...")
    result = engine.take_snapshot(args.period, note=args.note)

    if result.success:
        print(f"\n[완료] {result.message}")
        if result.variance_rows > 0:
            print(f"  → Variance가 감지되었습니다. order_book_snapshot.sql로 확인하세요.")
        return 0
    else:
        print(f"\n[오류] {result.message}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
