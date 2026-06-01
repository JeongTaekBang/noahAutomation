#!/usr/bin/env python
"""ob_snapshot 금액 컬럼 원 단위 ROUND 마이그레이션 (1회성, 멱등)
================================================================

배경:
    Order Book 스냅샷의 금액(KRW)이 REAL로 저장되며, 이른 마감 시점에
    동결된 소수점 값이 이후 정수로 정리된 원본과 < 1원씩 어긋나
    '유령 ending 잔량'을 남겼다. 이 잔량이 월별로 누적되어
    `close_period.py --list`의 'Start != 전월 Ending (!차이 1)' 경고를 유발.

해결:
    KRW는 정수 통화이므로 스냅샷 금액 5개 컬럼을 원 단위로 ROUND.
    - 유령 sub-won 잔량 소멸 (ending 정수화)
    - 연속성 회복: 드롭된 잔량은 전부 < 0.5원 → round=0 → Σ Start = Σ 전월 Ending
    - Variance(소급변경) 이력은 보존 (값만 정수화, 관계 불변)

    rolling/snapshot SQL(order_book.sql, order_book_snapshot.sql,
    snapshot.py)에는 이미 소스 금액 ROUND가 적용되어 향후 마감은 정수 유지.

멱등성: 정수를 ROUND해도 그대로이므로 반복 실행해도 안전.

사용법:
    python migrate_snapshot_round.py            # 적용
    python migrate_snapshot_round.py --dry-run  # 변경 건수만 확인
"""
from __future__ import annotations

import argparse
import sqlite3
import sys

from po_generator.config import DB_FILE

AMOUNT_COLS = [
    "start_amount", "input_amount", "output_amount",
    "variance_amount", "ending_amount",
]


def _continuity(conn: sqlite3.Connection) -> None:
    rows = conn.execute("""
        SELECT m.period,
               COALESCE(SUM(s.start_amount),0)  AS ts,
               COALESCE(SUM(s.ending_amount),0) AS te,
               COALESCE(SUM(s.variance_amount),0) AS tv,
               SUM(CASE WHEN s.ending_amount IS NOT NULL
                        AND ABS(s.ending_amount-ROUND(s.ending_amount))>1e-6
                        THEN 1 ELSE 0 END) AS nonint
        FROM ob_snapshot_meta m
        LEFT JOIN ob_snapshot s ON m.period=s.snapshot_period
        WHERE m.is_active=1
        GROUP BY m.period ORDER BY m.period
    """).fetchall()
    prev = None
    for period, ts, te, tv, nonint in rows:
        diff = "" if prev is None else f"  Start-전월Ending={ts-prev:+.4f}"
        print(f"  {period}  start={ts:>16,.2f}  ending={te:>16,.2f}  "
              f"var={tv:>14,.2f}  비정수ending={nonint:>3}{diff}")
        prev = te


def main() -> int:
    ap = argparse.ArgumentParser(description="ob_snapshot 금액 원 단위 ROUND")
    ap.add_argument("--dry-run", action="store_true", help="변경 건수만 확인")
    args = ap.parse_args()

    if not DB_FILE.exists():
        print(f"[오류] DB 파일이 없습니다: {DB_FILE}")
        return 1

    conn = sqlite3.connect(str(DB_FILE))

    # 변경 대상 건수 (어느 한 컬럼이라도 비정수)
    cond = " OR ".join(
        f"({c} IS NOT NULL AND ABS({c}-ROUND({c}))>1e-9)" for c in AMOUNT_COLS
    )
    n = conn.execute(f"SELECT COUNT(*) FROM ob_snapshot WHERE {cond}").fetchone()[0]
    total = conn.execute("SELECT COUNT(*) FROM ob_snapshot").fetchone()[0]

    print(f"DB: {DB_FILE}")
    print(f"총 스냅샷 행: {total},  비정수 금액 포함 행: {n}\n")
    print("=== BEFORE ===")
    _continuity(conn)

    if args.dry_run:
        print(f"\n[dry-run] {n}개 행이 ROUND 대상입니다. 변경하지 않았습니다.")
        conn.close()
        return 0

    set_clause = ", ".join(f"{c}=ROUND({c})" for c in AMOUNT_COLS)
    conn.execute(f"UPDATE ob_snapshot SET {set_clause}")
    conn.commit()

    print("\n=== AFTER ===")
    _continuity(conn)
    conn.close()
    print(f"\n[완료] {n}개 행의 금액을 원 단위로 ROUND했습니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
