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
import json
import sqlite3
import sys
import unicodedata
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.config import NOAH_SO_PO_DN_FILE, DB_FILE
from po_generator.db_schema import (
    SYNC_SHEETS, get_sync_metadata, get_table_row_count,
    ensure_sync_log_tables, create_sync_run, finalize_sync_run,
)
from po_generator.db_sync import SyncEngine, SyncSummary, _normalize_pk
from po_generator.logging_config import setup_logging


def _disp_width(s: str) -> int:
    """문자열의 터미널 표시폭 — 전각(W/F) 문자는 2칸으로 계산."""
    return sum(2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1 for ch in str(s))


def _ljust_w(s, width: int) -> str:
    """표시폭 기준 좌측 정렬(우측 공백 패딩)."""
    s = str(s)
    return s + ' ' * max(0, width - _disp_width(s))


def _rjust_w(s, width: int) -> str:
    """표시폭 기준 우측 정렬(좌측 공백 패딩)."""
    s = str(s)
    return ' ' * max(0, width - _disp_width(s)) + s


def _elapsed_str(started: str | None, ended: str | None) -> str:
    """started/ended('YYYY-MM-DD HH:MM:SS') 차이를 '12s' 형태로. 계산 불가 시 '-'."""
    if not started or not ended:
        return "-"
    try:
        fmt = "%Y-%m-%d %H:%M:%S"
        d = (datetime.strptime(ended, fmt) - datetime.strptime(started, fmt)).total_seconds()
        return f"{int(round(d))}s" if d >= 0 else "-"
    except Exception:
        return "-"


def print_summary(summary: SyncSummary, dry_run: bool = False) -> None:
    """동기화 결과를 테이블 형태로 출력 (전각/한글 표시폭 보정)"""
    mode = " (DRY-RUN)" if dry_run else ""

    # 컬럼: (헤더, 표시폭, 정렬 l=좌/r=우)
    cols = [("시트", 16, "l"), ("행수", 7, "r"), ("신규", 8, "r"),
            ("수정", 8, "r"), ("삭제", 8, "r"), ("에러", 7, "r")]
    table_w = sum(w for _, w, _ in cols) + (len(cols) - 1)

    def _row(values) -> str:
        parts = []
        for (_, w, align), v in zip(cols, values):
            parts.append(_ljust_w(v, w) if align == "l" else _rjust_w(v, w))
        return " ".join(parts)

    print(f"\nNOAH Excel → SQLite 동기화{mode}")
    print("=" * table_w)
    print(f"소스: {summary.source_file}")
    print(f"DB:   {summary.db_file}")

    print()
    print(_row([h for h, _, _ in cols]))
    print("-" * table_w)

    for r in summary.results:
        err_mark = f"*{r.errors}" if r.errors > 0 else f"{r.errors}"
        print(_row([r.sheet_name, r.total_rows, r.inserted, r.updated, r.pruned, err_mark]))

    print("-" * table_w)
    total_err = f"*{summary.total_errors}" if summary.total_errors > 0 else f"{summary.total_errors}"
    print(_row(["합계", summary.total_rows, summary.total_inserted,
                summary.total_updated, summary.total_pruned, total_err]))

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


def print_changes(summary: SyncSummary, rolled_back: bool = False) -> None:
    """신규/수정/삭제된 레코드 상세 출력.

    rolled_back=True면 에러로 트랜잭션이 ROLLBACK되어 아래 변경이 실제 DB에
    반영되지 않았음을 경고 헤더로 명시한다(콘솔 출력만 보고 적용된 것으로 오인 방지).
    """
    has_changes = any(r.inserted_details or r.updated_details or r.pruned_pks for r in summary.results)
    if not has_changes:
        print("\n변경 사항 없음")
        return

    if rolled_back:
        print("\n" + "!" * 60)
        print("[주의] 에러로 동기화가 ROLLBACK됨 — 아래 변경은 DB에 적용되지 않았습니다.")
        print("       데이터 수정 후 재실행하세요. (아래는 참고용 상세)")
        print("!" * 60)

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


def _to_text(val) -> str | None:
    """로그 값 → JSON-호환 텍스트. None/빈 문자열은 None으로 통일."""
    if val is None:
        return None
    s = str(val)
    return s if s else None


def _jdump(obj) -> str:
    """JSON 직렬화 — 한글 비-escape, 키 순서 유지."""
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'))


# 시트명 → PK 컬럼 (재키잉 인식 시 '키 컬럼'을 비교에서 제외하기 위함)
# SYNC_SHEETS 밖 결과(FX 등)는 여기 없어 pk_columns=()로 떨어지고, _reconcile_rekeys가
# 빈 키에서 즉시 통과한다 — 재키잉 페어링은 위치성 키를 가진 세로형 시트 전용(의도된 제외).
_PK_BY_SHEET: dict[str, tuple[str, ...]] = {c.sheet_name: c.pk_columns for c in SYNC_SHEETS}


def _content_matches(snap: dict, vals: dict, key_set: set[str]) -> bool:
    """삭제 스냅샷(snap)과 신규 값(vals)이 '같은 논리 행'인지 — 비키 컬럼 내용 일치 판정.

    - 키 컬럼(key_set)은 비교에서 제외 (재키잉으로 바뀌는 부분이라 당연히 다름).
    - snap에 있는 모든 비키 컬럼이 vals에서 동일해야 함 (snap ⊆ vals).
      vals가 더 많은 컬럼을 가질 수 있음(빈 키 채우며 함께 입력된 값) — 그건 허용.
    - 최소 1개 이상의 비키 컬럼이 실제로 매칭돼야 함 (빈 행끼리 오매칭 방지).
    """
    matched = 0
    for col, sv in snap.items():
        if col in key_set:
            continue
        if _to_text(sv) != _to_text(vals.get(col)):
            return False
        matched += 1
    return matched >= 1


def _reconcile_rekeys(pk_columns: tuple[str, ...],
                      inserted_details: list[dict],
                      pruned_snapshots: list[dict]) -> tuple[list[dict], list[dict], list[dict]]:
    """동일 논리 행의 (삭제 스냅샷 ↔ 신규 값)을 1:1로 묶어 '키변경(수정)'으로 합친다.

    위치성 키 컬럼(Line item/_row_seq 등)을 편집하면 같은 행이 삭제+신규로 기록되는데,
    실제로는 수정이다. 같은 sync 안에서 문서ID(pk[0])가 같고 비키 내용이 일치하는
    삭제·신규 쌍을 찾아 단일 '수정' 이벤트로 변환한다.

    Returns: (rekey_events, 잔여_inserted, 잔여_pruned)
        rekey_events: [{'pk': new_pk_tuple, 'changes': {col: {old, new}}}]
    """
    if not pk_columns or not inserted_details or not pruned_snapshots:
        return [], list(inserted_details), list(pruned_snapshots)

    key_set = set(pk_columns)
    inserts = [
        {'pk': _normalize_pk(tuple(d['pk'])), 'values': d.get('values', {}), 'used': False}
        for d in inserted_details
    ]
    rekeys: list[dict] = []
    rem_pruned: list[dict] = []

    for s in pruned_snapshots:
        spk = _normalize_pk(tuple(s['pk']))
        snap = s.get('snapshot', {})
        cands = [
            i for i, ins in enumerate(inserts)
            if not ins['used']
            and ins['pk'][0] == spk[0]                       # 같은 문서ID (재키잉 시 불변)
            and _content_matches(snap, ins['values'], key_set)
        ]
        if len(cands) == 1:                                  # 모호하면(0/2+) 묶지 않음 → 삭제 유지
            ins = inserts[cands[0]]
            ins['used'] = True
            # 옛 행/새 행의 '전체 표현' 비교로 변경 컬럼 산출.
            # 키 컬럼 값은 PK 튜플(권위 있는 키)에서, 비키 값은 snapshot/values에서.
            old_map = {**dict(zip(pk_columns, spk)), **snap}
            new_map = {**dict(zip(pk_columns, ins['pk'])), **ins['values']}
            changes = {}
            for col in list(old_map.keys()) + [c for c in new_map if c not in old_map]:
                old, new = _to_text(old_map.get(col)), _to_text(new_map.get(col))
                if old != new:
                    changes[col] = {'old': old, 'new': new}
            rekeys.append({'pk': ins['pk'], 'changes': changes})
        else:
            rem_pruned.append(s)

    rem_inserted = [d for d, ins in zip(inserted_details, inserts) if not ins['used']]
    return rekeys, rem_inserted, rem_pruned


def write_sync_log_to_db(summary: SyncSummary, note: str | None = None,
                         db_path: Path = DB_FILE) -> None:
    """동기화 변경 내역을 _sync_log v2 스키마에 기록.

    record(레코드)당 1행으로 압축 저장:
    - 신규: changes_json = {col: value, ...}, row_snapshot_json = NULL
    - 수정: changes_json = {col: {old, new}, ...}, row_snapshot_json = NULL
    - 삭제: changes_json = NULL, row_snapshot_json = {col: value, ...}

    변경이 0건이어도 _sync_runs 실행 이력은 1행 남긴다(누가/언제 동기화했는지 감사).
    PK는 _normalize_pk로 통일 — 동일 레코드의 신규/수정/삭제 이벤트가 같은 키로 남도록
    ('1.0' vs '1' 비대칭 방지). sync 트랜잭션과 분리된 별도 트랜잭션으로 기록하며,
    호출부(main)가 예외를 격리하므로 로그 실패가 이미 commit된 sync 결과를 해치지 않는다.
    """
    rows: list[tuple] = []
    # placeholder sync_id — INSERT 시점에 채움
    for r in summary.results:
        # PK 정의 변경으로 테이블을 재적재한 경우: 전 행이 '신규'로 잡히지만 데이터가
        # 들어온 게 아니라 키 체계가 바뀐 것이다. 수천 건의 가짜 '신규'로 이력을
        # 덮는 대신 '재적재' 1건만 남긴다 (진짜 변경은 다음 sync부터 정상 기록).
        if r.pk_migrated:
            rows.append((r.sheet_name, '재적재', _jdump([r.table_name]), r.table_name,
                         _jdump({'reason': 'PK 정의 변경 → 테이블 재생성',
                                 'rows': r.total_rows}), None))
            continue

        # 재키잉 인식: 위치성 키(Line item 등) 편집으로 발생한 삭제+신규 쌍을
        # '키변경(수정)' 단일 이벤트로 합쳐 "데이터는 있는데 삭제로 뜨는" 오해 제거.
        # 묶이지 않은 신규/삭제만 그대로 신규/삭제로 기록한다.
        pk_cols = _PK_BY_SHEET.get(r.sheet_name, ())
        rekeys, inserted_details, pruned_snapshots = _reconcile_rekeys(
            pk_cols, r.inserted_details, r.pruned_snapshots,
        )

        # 신규 (재키잉으로 묶이지 않은 것만)
        for detail in inserted_details:
            pk_tuple = _normalize_pk(tuple(detail['pk']))
            pk_json = _jdump(list(pk_tuple))
            pk_disp = _format_pk(pk_tuple)
            changes = {col: _to_text(val) for col, val in detail['values'].items()
                       if _to_text(val) is not None}
            rows.append((r.sheet_name, '신규', pk_json, pk_disp,
                         _jdump(changes) if changes else None, None))

        # 수정
        for detail in r.updated_details:
            pk_tuple = _normalize_pk(tuple(detail['pk']))
            pk_json = _jdump(list(pk_tuple))
            pk_disp = _format_pk(pk_tuple)
            changes = {col: {'old': _to_text(old), 'new': _to_text(new)}
                       for col, (old, new) in detail['changes'].items()}
            rows.append((r.sheet_name, '수정', pk_json, pk_disp,
                         _jdump(changes) if changes else None, None))

        # 수정 (재키잉) — 같은 논리 행의 키가 바뀐 것. 살아남은 새 PK 기준으로 기록.
        for rk in rekeys:
            pk_tuple = _normalize_pk(tuple(rk['pk']))
            pk_json = _jdump(list(pk_tuple))
            pk_disp = _format_pk(pk_tuple)
            changes = rk['changes']
            rows.append((r.sheet_name, '수정', pk_json, pk_disp,
                         _jdump(changes) if changes else None, None))

        # 삭제 — pruned_snapshots는 물리 행 단위(각자 pk+snapshot)다. pruned_pks를
        # 재키잉하면 동일 정규화 pk가 여러 물리행을 가질 때 스냅샷이 유실/중복되므로
        # pruned_snapshots를 직접 순회해 행:스냅샷을 1:1로 보존한다.
        for s in pruned_snapshots:
            pk_tuple = _normalize_pk(tuple(s['pk']))
            pk_json = _jdump(list(pk_tuple))
            pk_disp = _format_pk(pk_tuple)
            snap = s.get('snapshot', {})
            snap_clean = {k: _to_text(v) for k, v in snap.items() if _to_text(v) is not None}
            rows.append((r.sheet_name, '삭제', pk_json, pk_disp,
                         None, _jdump(snap_clean) if snap_clean else None))

    conn = sqlite3.connect(str(db_path))
    try:
        conn.execute('PRAGMA foreign_keys=ON')  # _sync_log.sync_id → _sync_runs FK 강제
        ensure_sync_log_tables(conn)
        sync_id = create_sync_run(conn, dry_run=False, note=note,
                                  started_at=summary.started_at or None)
        if rows:
            rows_with_id = [(sync_id, *r) for r in rows]
            conn.executemany(
                "INSERT INTO _sync_log "
                "(sync_id, sheet_name, change_type, pk_json, pk_display, "
                " changes_json, row_snapshot_json) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                rows_with_id,
            )
        finalize_sync_run(conn, sync_id, len(rows))
        conn.commit()
    finally:
        conn.close()

    if rows:
        print(f"\n동기화 로그 저장: _sync_log {len(rows):,}행 (sync_id={sync_id})")
    else:
        print(f"\n동기화 로그: 변경 없음 — 실행 이력만 기록 (sync_id={sync_id})")
    print("  변경 이력 조회: python sync_db.py --log   또는  대시보드 '동기화 로그' 페이지")


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


def show_log(limit: int = 20) -> int:
    """최근 동기화 세션 이력 조회 (_sync_runs 기준, 최신순)."""
    if not DB_FILE.exists():
        print(f"DB 파일이 없습니다: {DB_FILE}")
        print("sync_db.py를 먼저 실행하세요.")
        return 1

    conn = sqlite3.connect(str(DB_FILE))
    try:
        try:
            runs = conn.execute(
                "SELECT sync_id, started_at, ended_at, actor, host, total_changes, note "
                "FROM _sync_runs ORDER BY sync_id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        except sqlite3.OperationalError:
            print("동기화 세션 이력이 없습니다 (_sync_runs 테이블 없음 — 먼저 동기화 실행).")
            return 0

        if not runs:
            print("동기화 세션 이력이 없습니다.")
            return 0

        cols = [("sync_id", 8, "r"), ("시작", 21, "l"), ("소요", 6, "r"),
                ("변경", 6, "r"), ("사용자", 14, "l"), ("호스트", 16, "l"), ("메모", 20, "l")]
        table_w = sum(w for _, w, _ in cols) + (len(cols) - 1)

        def _row(values) -> str:
            parts = []
            for (_, w, align), v in zip(cols, values):
                parts.append(_ljust_w(v, w) if align == "l" else _rjust_w(v, w))
            return " ".join(parts)

        print(f"\n최근 동기화 세션 (최대 {limit}개, 최신순)")
        print("=" * table_w)
        print(_row([h for h, _, _ in cols]))
        print("-" * table_w)
        for sync_id, started, ended, actor, host, total, note in runs:
            print(_row([
                sync_id, started or "-", _elapsed_str(started, ended),
                total if total is not None else 0,
                actor or "-", host or "-", (note or "")[:20],
            ]))
        print("-" * table_w)
        print("\n변경 상세는 대시보드 '동기화 로그' 페이지에서 조회하세요.")
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
        '--log',
        nargs='?',
        const=20,
        type=int,
        metavar='N',
        help='최근 N개 동기화 세션 이력 조회 (기본 20, 동기화 수행 안 함)',
    )

    parser.add_argument(
        '--note',
        metavar='TEXT',
        help='이 동기화 세션에 남길 메모 (_sync_runs.note에 기록)',
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

    # 동기화 세션 이력 조회 (동기화 수행 안 함)
    if args.log is not None:
        return show_log(args.log)

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

    if args.changes:
        # 에러로 ROLLBACK된 경우 '미적용' 경고를 함께 출력 (적용된 것으로 오인 방지)
        print_changes(summary, rolled_back=summary.total_errors > 0)

    # 실제 commit된 동기화에 한해서만 _sync_log 기록 — db_sync.sync_all은
    # total_errors>0이면 전체 트랜잭션을 ROLLBACK 하므로, 그 경우 변경기록을 남기면
    # 적용되지 않은 유령 변경이 감사로그/대시보드 변경이력에 노출된다.
    # 로그 기록 실패가 이미 commit된 sync 결과/요약 출력을 가리지 않도록 예외 격리.
    if not args.dry_run and summary.total_errors == 0:
        try:
            write_sync_log_to_db(summary, note=args.note)
        except Exception as e:
            print(f"[경고] 동기화 로그 저장 실패 (데이터는 정상 반영됨): {e}")

    print_summary(summary, dry_run=args.dry_run)

    return 1 if summary.total_errors > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
