#!/usr/bin/env python
"""
NOAH Excel → SQLite 동기화
===========================

NOAH_SO_PO_DN.xlsx의 수동 입력 시트(SO, PO, DN, PMT)를
SQLite DB에 업로드하여 데이터를 안전하게 백업합니다.

사용법:
    python sync_db.py                           # 전체 동기화
    python sync_db.py -v                        # 상세 로그
    python sync_db.py --sheets SO_국내 PO_국내  # 특정 시트만
    python sync_db.py --dry-run                 # 시뮬레이션
    python sync_db.py --info                    # DB 현황 조회
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.config import NOAH_SO_PO_DN_FILE, DB_FILE, DATA_DIR
from po_generator.db_schema import SYNC_SHEETS, get_sync_metadata, get_table_row_count
from po_generator.db_sync import SyncEngine, SyncSummary
from po_generator.logging_config import setup_logging


def print_summary(summary: SyncSummary, dry_run: bool = False) -> None:
    """동기화 결과를 테이블 형태로 출력"""
    mode = " (DRY-RUN)" if dry_run else ""
    print(f"\nNOAH Excel → SQLite 동기화{mode}")
    print("=" * 56)
    print(f"소스: {summary.source_file}")
    print(f"DB:   {summary.db_file}")

    # 테이블 헤더
    print()
    print(f"{'시트':<14} {'행수':>6} {'신규':>8} {'수정':>8} {'삭제':>8} {'에러':>6}")
    print("-" * 64)

    for r in summary.results:
        err_mark = f"  *{r.errors}" if r.errors > 0 else f"  {r.errors}"
        print(f"{r.sheet_name:<14} {r.total_rows:>6} {r.inserted:>8} {r.updated:>8} {r.pruned:>8} {err_mark:>6}")

    print("-" * 64)
    print(
        f"{'합계':<14} {summary.total_rows:>6} "
        f"{summary.total_inserted:>8} {summary.total_updated:>8} "
        f"{summary.total_pruned:>8} "
        f"{'  *' + str(summary.total_errors) if summary.total_errors > 0 else '  ' + str(summary.total_errors):>6}"
    )

    print(f"\n소요시간: {summary.elapsed_seconds:.1f}초")

    # 에러 상세
    for r in summary.results:
        if r.error_messages:
            print(f"\n[에러] {r.sheet_name}:")
            for msg in r.error_messages[:5]:
                print(f"  - {msg}")
            if len(r.error_messages) > 5:
                print(f"  ... 외 {len(r.error_messages) - 5}건")


def _format_pk(pk: tuple) -> str:
    return ' | '.join(str(v) for v in pk)


def _format_val(val) -> str:
    if val is None:
        return '(빈값)'
    s = str(val)
    return s[:40] + '...' if len(s) > 40 else s


def print_changes(summary: SyncSummary) -> None:
    """신규/수정/삭제된 레코드 상세 출력"""
    has_changes = any(r.inserted_details or r.updated_details or r.pruned_pks for r in summary.results)
    if not has_changes:
        print("\n변경 사항 없음")
        return

    for r in summary.results:
        if not r.inserted_details and not r.updated_details and not r.pruned_pks:
            continue

        print(f"\n--- {r.sheet_name} ---")

        if r.inserted_details:
            print(f"  [신규] {len(r.inserted_details)}건:")
            for detail in r.inserted_details:
                print(f"    + {_format_pk(detail['pk'])}")
                for col, val in detail['values'].items():
                    print(f"        {col}: {_format_val(val)}")

        if r.updated_details:
            print(f"  [수정] {len(r.updated_details)}건:")
            for detail in r.updated_details[:20]:
                print(f"    ~ {_format_pk(detail['pk'])}")
                for col, (old, new) in detail['changes'].items():
                    print(f"        {col}: {_format_val(old)} → {_format_val(new)}")
            if len(r.updated_details) > 20:
                print(f"    ... 외 {len(r.updated_details) - 20}건")

        if r.pruned_pks:
            print(f"  [삭제] {len(r.pruned_pks)}건:")
            for pk in r.pruned_pks[:20]:
                print(f"    - {_format_pk(pk)}")
            if len(r.pruned_pks) > 20:
                print(f"    ... 외 {len(r.pruned_pks) - 20}건")


SYNC_LOG_FILE: Path = DATA_DIR / "sync_log.csv"
CSV_HEADER = "동기화시각,시트,유형,PK,컬럼,이전값,변경값\n"


def _csv_escape(val) -> str:
    """CSV 셀 값 이스케이프 (쉼표, 줄바꿈, 따옴표 처리)"""
    s = str(val) if val is not None else ''
    if ',' in s or '"' in s or '\n' in s:
        return '"' + s.replace('"', '""') + '"'
    return s


def write_sync_log(summary: SyncSummary) -> None:
    """동기화 변경 내역을 CSV 로그 파일에 기록"""
    has_changes = any(r.inserted_details or r.updated_details or r.pruned_pks for r in summary.results)
    if not has_changes:
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    is_new_file = not SYNC_LOG_FILE.exists() or SYNC_LOG_FILE.stat().st_size == 0

    with open(SYNC_LOG_FILE, "a", encoding="utf-8-sig") as f:
        # BOM + 헤더 (새 파일일 때만)
        if is_new_file:
            f.write(CSV_HEADER)

        for r in summary.results:
            # 신규 (필드별 1행 — 비어있지 않은 값만)
            for detail in r.inserted_details:
                pk_str = _csv_escape(_format_pk(detail['pk']))
                for col, val in detail['values'].items():
                    col_esc = _csv_escape(col)
                    val_esc = _csv_escape(val)
                    f.write(f"{now},{r.sheet_name},신규,{pk_str},{col_esc},,{val_esc}\n")

            # 수정 (필드별 1행)
            for detail in r.updated_details:
                pk_str = _csv_escape(_format_pk(detail['pk']))
                for col, (old, new) in detail['changes'].items():
                    old_esc = _csv_escape(old)
                    new_esc = _csv_escape(new)
                    col_esc = _csv_escape(col)
                    f.write(f"{now},{r.sheet_name},수정,{pk_str},{col_esc},{old_esc},{new_esc}\n")

            # 삭제 (PK만 기록)
            for pk in r.pruned_pks:
                pk_str = _csv_escape(_format_pk(pk))
                f.write(f"{now},{r.sheet_name},삭제,{pk_str},,,\n")

    print(f"\n동기화 로그 저장: {SYNC_LOG_FILE.name}")


def show_info() -> int:
    """DB 현황 조회"""
    if not DB_FILE.exists():
        print(f"DB 파일이 없습니다: {DB_FILE}")
        print("sync_db.py를 먼저 실행하세요.")
        return 1

    print(f"\nNOAH SQLite DB 현황")
    print("=" * 60)
    print(f"DB: {DB_FILE}")
    print(f"크기: {DB_FILE.stat().st_size / 1024:.1f} KB")

    conn = sqlite3.connect(str(DB_FILE))
    try:
        meta = get_sync_metadata(conn)

        print(f"\n{'테이블':<16} {'행수':>8} {'마지막 동기화':>22}")
        print("-" * 60)

        total = 0
        for config in SYNC_SHEETS:
            row_count = get_table_row_count(conn, config.table_name)
            total += row_count
            m = meta.get(config.table_name, {})
            last_sync = m.get('last_sync', '-')
            if last_sync and last_sync != '-':
                # ISO → 읽기 쉬운 형식
                last_sync = last_sync[:19].replace('T', ' ')
            print(f"{config.table_name:<16} {row_count:>8} {last_sync:>22}")

        print("-" * 60)
        print(f"{'합계':<16} {total:>8}")

    finally:
        conn.close()

    return 0


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    parser = argparse.ArgumentParser(
        prog='sync_db',
        description='NOAH Excel → SQLite 동기화 — 데이터 백업 및 관리',
        epilog='예시: python sync_db.py --sheets SO_국내 PO_국내 -v',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        '--sheets',
        nargs='+',
        metavar='SHEET',
        help='동기화할 시트명 (기본: 전체)',
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='실제 DB 변경 없이 시뮬레이션만 수행',
    )

    parser.add_argument(
        '--info',
        action='store_true',
        help='DB 현황 조회 (동기화 수행 안 함)',
    )

    parser.add_argument(
        '--changes',
        action='store_true',
        help='동기화 후 신규/수정된 레코드 상세 표시',
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )

    return parser


def main() -> int:
    """메인 함수"""
    parser = create_argument_parser()
    args = parser.parse_args()

    setup_logging(verbose=args.verbose)

    # DB 현황 조회
    if args.info:
        return show_info()

    # Excel 파일 존재 확인
    if not NOAH_SO_PO_DN_FILE.exists():
        print(f"[오류] Excel 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")
        return 1

    # 동기화 실행
    engine = SyncEngine()
    try:
        summary = engine.sync_all(
            dry_run=args.dry_run,
            sheet_filter=args.sheets,
        )
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1
    except Exception as e:
        print(f"[오류] 동기화 실패: {e}")
        return 1

    print_summary(summary, dry_run=args.dry_run)

    if args.changes:
        print_changes(summary)

    # dry-run이 아니면 로그 파일에 변경 내역 기록
    if not args.dry_run:
        write_sync_log(summary)

    return 1 if summary.total_errors > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
