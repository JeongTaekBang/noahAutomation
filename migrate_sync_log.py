#!/usr/bin/env python
"""
sync_log.csv → _sync_log_legacy_v1 테이블 마이그레이션 (1회성, 구 v1 형식)
==========================================================================

기존 CSV 로그(v1 7컬럼)를 SQLite `_sync_log_legacy_v1` 테이블로 이관.

주의: 현재 운영 `_sync_log`는 v2 스키마(record 단위 + JSON)다. v1 CSV는 컬럼 구성이
달라 v2 테이블에 직접 넣을 수 없으므로, 충돌·오염을 막기 위해 전용 legacy 테이블에
적재한다. 이후 v2로 합치려면 migrate_sync_log_v2.py 로 변환하라.
실행 후 CSV 파일은 삭제해도 되지만, 안전을 위해 기본은 유지.

사용법:
    python migrate_sync_log.py             # 마이그레이션 + 건수 출력
    python migrate_sync_log.py --delete    # 마이그레이션 성공 시 CSV 삭제
    python migrate_sync_log.py --dry-run   # 파싱만, DB 쓰기 없음
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import sys
from pathlib import Path

from po_generator.config import DATA_DIR, DB_FILE

CSV_FILE: Path = DATA_DIR / "sync_log.csv"
LEGACY_TABLE = "_sync_log_legacy_v1"


def _ensure_legacy_table(conn: sqlite3.Connection) -> None:
    """v1 CSV 전용 legacy 테이블 생성 (운영 v2 _sync_log와 분리)."""
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {LEGACY_TABLE} (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            sync_time   TEXT,
            sheet_name  TEXT,
            change_type TEXT,
            pk          TEXT,
            column_name TEXT,
            old_value   TEXT,
            new_value   TEXT
        )
        """
    )


def _parse_csv(path: Path) -> list[tuple]:
    """CSV 파싱 → INSERT rows"""
    rows: list[tuple] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        if header is None:
            return rows
        for raw in reader:
            if not raw:
                continue
            # 7컬럼 고정: 동기화시각,시트,유형,PK,컬럼,이전값,변경값
            if len(raw) < 7:
                raw = raw + [''] * (7 - len(raw))
            sync_time, sheet, change_type, pk, col, old, new = raw[:7]
            if not sync_time or not sheet or not change_type or not pk:
                continue
            rows.append((
                sync_time, sheet, change_type, pk,
                col or None, old or None, new or None,
            ))
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="sync_log.csv → _sync_log 테이블 마이그레이션")
    ap.add_argument("--dry-run", action="store_true", help="파싱만 수행, DB 쓰기 없음")
    ap.add_argument("--delete", action="store_true", help="마이그레이션 성공 후 CSV 파일 삭제")
    args = ap.parse_args()

    if not CSV_FILE.exists():
        print(f"CSV 파일이 없습니다: {CSV_FILE}")
        return 1
    if not DB_FILE.exists():
        print(f"DB 파일이 없습니다: {DB_FILE}")
        print("먼저 sync_db.py를 실행해 DB를 생성하세요.")
        return 1

    csv_size = CSV_FILE.stat().st_size
    print(f"CSV: {CSV_FILE}")
    print(f"크기: {csv_size / 1024 / 1024:.2f} MB")

    rows = _parse_csv(CSV_FILE)
    print(f"파싱 완료: {len(rows):,}행")

    if not rows:
        print("마이그레이션할 데이터 없음")
        return 0

    if args.dry_run:
        print("\n[DRY-RUN] DB 쓰기 없이 종료")
        print("예시 (처음 3행):")
        for r in rows[:3]:
            print(f"  {r}")
        return 0

    conn = sqlite3.connect(str(DB_FILE))
    try:
        _ensure_legacy_table(conn)
        existing = conn.execute(f"SELECT COUNT(*) FROM {LEGACY_TABLE}").fetchone()[0]
        if existing > 0:
            print(f"\n경고: {LEGACY_TABLE}에 이미 {existing:,}행 존재")
            # 비대화형(CI/리다이렉트/파이프) 환경에서 input()이 EOFError로 죽거나
            # hang되는 것을 방지 — stdin이 tty가 아니면 중복 삽입을 막기 위해 취소.
            if not sys.stdin.isatty():
                print("비대화형 실행 감지 — 중복 삽입 방지를 위해 취소. "
                      "(의도적이면 대화형 터미널에서 재실행)")
                return 0
            try:
                ans = input("계속하면 중복 가능. 진행? [y/N]: ").strip().lower()
            except EOFError:
                print("입력 없음 — 취소됨")
                return 0
            if ans != 'y':
                print("취소됨")
                return 0

        conn.executemany(
            f"INSERT INTO {LEGACY_TABLE} "
            "(sync_time, sheet_name, change_type, pk, column_name, old_value, new_value) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            rows,
        )
        conn.commit()
        final = conn.execute(f"SELECT COUNT(*) FROM {LEGACY_TABLE}").fetchone()[0]
        print(f"\n완료: {LEGACY_TABLE} 총 {final:,}행 (이번 삽입 {len(rows):,}행)")
        print("v2 _sync_log로 합치려면 migrate_sync_log_v2.py 로 변환하세요.")
    finally:
        conn.close()

    if args.delete:
        CSV_FILE.unlink()
        print(f"CSV 삭제됨: {CSV_FILE.name}")
    else:
        print(f"CSV 파일 유지: {CSV_FILE.name} (필요 없으면 수동 삭제 또는 --delete 옵션)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
