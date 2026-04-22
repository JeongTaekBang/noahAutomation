#!/usr/bin/env python
"""
_sync_log v1 → v2 마이그레이션 (1회성)
=======================================

기존 v1 스키마(필드 단위 행, sync_time 문자열로 그룹핑)를
v2 스키마(record 단위 행, sync_id FK + JSON)로 변환.

변경 사항:
- _sync_runs 테이블 신설 (세션 메타: actor, host, dry_run, total_changes)
- _sync_log 재구성:
  - 신규: 컬럼별 N행 → record당 1행, changes_json = {col: val,...}
  - 수정: 컬럼별 N행 → record당 1행, changes_json = {col: {old, new},...}
  - 삭제: PK만 → 그대로 (snapshot은 v1에 없음, NULL)
- pk_json 추가 (JSON 배열) + pk_display 유지 (검색 호환)

기존 _sync_log는 _sync_log_legacy로 백업.

사용법:
    python migrate_sync_log_v2.py             # 마이그레이션 실행
    python migrate_sync_log_v2.py --dry-run   # 변환 통계만 출력
    python migrate_sync_log_v2.py --drop-legacy   # 마이그레이션 + legacy 테이블 제거
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

from po_generator.config import DB_FILE
from po_generator.db_schema import ensure_sync_log_tables


def _jdump(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'))


def _is_v2_schema(conn: sqlite3.Connection) -> bool:
    """현재 _sync_log가 v2인지 판정 (changes_json 컬럼 존재 여부)."""
    cols = {r[1] for r in conn.execute("PRAGMA table_info(_sync_log)")}
    return 'changes_json' in cols and 'sync_id' in cols


def _legacy_exists(conn: sqlite3.Connection) -> bool:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='_sync_log'"
    ).fetchone()
    if not row:
        return False
    cols = {r[1] for r in conn.execute("PRAGMA table_info(_sync_log)")}
    # v1 식별: pk + column_name + old_value + new_value 조합
    return {'pk', 'column_name', 'old_value', 'new_value'}.issubset(cols)


def main() -> int:
    ap = argparse.ArgumentParser(description="_sync_log v1 → v2 마이그레이션")
    ap.add_argument("--dry-run", action="store_true", help="변환만 시뮬레이션 (DB 변경 없음)")
    ap.add_argument("--drop-legacy", action="store_true", help="마이그레이션 후 _sync_log_legacy 삭제")
    args = ap.parse_args()

    if not DB_FILE.exists():
        print(f"DB 파일이 없습니다: {DB_FILE}")
        return 1

    conn = sqlite3.connect(str(DB_FILE))
    try:
        # 0. 현재 상태 확인
        if _is_v2_schema(conn):
            print("이미 v2 스키마입니다. 마이그레이션 불필요.")
            return 0

        if not _legacy_exists(conn):
            print("v1 _sync_log 테이블이 없습니다. 신규 v2 스키마만 생성.")
            if args.dry_run:
                print("[DRY-RUN] v2 스키마 생성 생략")
                return 0
            ensure_sync_log_tables(conn)
            conn.commit()
            print("v2 스키마 생성 완료.")
            return 0

        # 1. v1 데이터 로드
        legacy_rows = conn.execute(
            "SELECT sync_time, sheet_name, change_type, pk, "
            "       column_name, old_value, new_value "
            "FROM _sync_log ORDER BY id"
        ).fetchall()
        print(f"v1 _sync_log: {len(legacy_rows):,}행 로드")

        # 2. 그룹핑 — (sync_time, sheet, change_type, pk) 단위로 컬럼들 묶기
        groups: dict[tuple, list] = defaultdict(list)
        for st, sheet, ctype, pk, col, oldv, newv in legacy_rows:
            groups[(st, sheet, ctype, pk)].append((col, oldv, newv))
        print(f"record 단위로 그룹핑: {len(groups):,}행으로 압축 (기존 대비 {len(legacy_rows)/max(len(groups),1):.1f}x)")

        # 3. _sync_runs 후보 — distinct sync_time
        distinct_times = sorted({k[0] for k in groups.keys()})
        print(f"_sync_runs 생성 예정: {len(distinct_times):,}개 세션")

        # 4. v2 row 빌드 (sync_id는 INSERT 후에 매핑)
        records_per_run: dict[str, int] = defaultdict(int)
        v2_rows: list[tuple] = []  # (sync_time, sheet, ctype, pk_json, pk_display, changes_json, snap_json)
        for (st, sheet, ctype, pk), cols in groups.items():
            pk_parts = pk.split(' | ') if pk else []
            pk_json = _jdump(pk_parts)
            if ctype == '신규':
                changes = {c: nv for c, _ov, nv in cols if c}
                changes_json = _jdump(changes) if changes else None
                snap_json = None
            elif ctype == '수정':
                changes = {c: {'old': ov, 'new': nv} for c, ov, nv in cols if c}
                changes_json = _jdump(changes) if changes else None
                snap_json = None
            else:  # 삭제
                changes_json = None
                snap_json = None  # legacy 데이터에는 스냅샷 없음
            v2_rows.append((st, sheet, ctype, pk_json, pk, changes_json, snap_json))
            records_per_run[st] += 1

        print(f"v2 _sync_log INSERT 예정: {len(v2_rows):,}행")

        if args.dry_run:
            print("\n[DRY-RUN] 종료")
            print("샘플 변환 (처음 3개):")
            for r in v2_rows[:3]:
                st, sheet, ctype, pk_json, pk_disp, ch, sn = r
                print(f"  [{st}] {sheet} {ctype} pk={pk_disp}")
                if ch:
                    print(f"    changes: {ch[:120]}{'...' if len(ch) > 120 else ''}")
            return 0

        # 5. 실제 변경 — 트랜잭션
        conn.execute('BEGIN')
        # 5a. 기존 _sync_log → _sync_log_legacy 백업
        conn.execute("ALTER TABLE _sync_log RENAME TO _sync_log_legacy")
        print("_sync_log → _sync_log_legacy 백업 완료")

        # 5b. v2 스키마 생성
        ensure_sync_log_tables(conn)

        # 5c. _sync_runs 채우기 → sync_id 매핑 확보
        sync_id_map: dict[str, int] = {}
        for st in distinct_times:
            cur = conn.execute(
                "INSERT INTO _sync_runs "
                "(started_at, ended_at, actor, host, dry_run, total_changes, note) "
                "VALUES (?, ?, NULL, NULL, 0, ?, 'migrated from v1')",
                (st, st, records_per_run[st]),
            )
            sync_id_map[st] = cur.lastrowid
        print(f"_sync_runs INSERT 완료: {len(sync_id_map):,}개")

        # 5d. v2 _sync_log INSERT
        rows_with_id = [
            (sync_id_map[r[0]], r[1], r[2], r[3], r[4], r[5], r[6])
            for r in v2_rows
        ]
        conn.executemany(
            "INSERT INTO _sync_log "
            "(sync_id, sheet_name, change_type, pk_json, pk_display, "
            " changes_json, row_snapshot_json) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            rows_with_id,
        )
        print(f"_sync_log v2 INSERT 완료: {len(rows_with_id):,}행")

        # 5e. legacy 정리 (옵션)
        if args.drop_legacy:
            conn.execute("DROP TABLE _sync_log_legacy")
            print("_sync_log_legacy 삭제됨")
        else:
            print("_sync_log_legacy 보존 — 검증 후 수동 삭제 권장")
            print("  DROP TABLE _sync_log_legacy;  -- SQL")
            print("  python migrate_sync_log_v2.py --drop-legacy  -- 다시 실행")

        conn.commit()
        print("\n마이그레이션 완료.")

        # 검증 출력
        v2_count = conn.execute("SELECT COUNT(*) FROM _sync_log").fetchone()[0]
        runs_count = conn.execute("SELECT COUNT(*) FROM _sync_runs").fetchone()[0]
        print(f"\n[검증]")
        print(f"  _sync_runs : {runs_count:,}개 세션")
        print(f"  _sync_log  : {v2_count:,}행 (v1 {len(legacy_rows):,} → v2 {v2_count:,}, "
              f"{(1 - v2_count/max(len(legacy_rows),1))*100:.1f}% 압축)")
    except Exception as e:
        conn.rollback()
        print(f"[오류] 롤백됨: {e}")
        return 1
    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
