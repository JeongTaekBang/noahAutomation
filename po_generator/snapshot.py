"""
Order Book 스냅샷 엔진
=====================

월별 마감(스냅샷)을 통해 Start를 고정하고,
소급 변경분을 Variance로 자동 감지합니다.
"""

from __future__ import annotations

import re
import sqlite3
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from po_generator.config import DB_FILE
from po_generator.db_schema import create_snapshot_tables

logger = logging.getLogger(__name__)

# order_book.sql의 이벤트 기반 CTE (os_grouped까지)
# 특정 period의 결과를 누적 SUM 패턴으로 추출
_ORDER_BOOK_BASE_SQL = """
WITH
so_combined AS (
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL) AS [Item qty],
        CAST([Sales amount] AS REAL) AS [Sales amount KRW],
        Period, [AX Period], [Model code], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
    UNION ALL
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER),
        CAST([Item qty] AS REAL),
        CAST([Sales amount KRW] AS REAL),
        Period, [AX Period], [Model code], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '해외'
    FROM so_export
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
),
dn_combined AS (
    SELECT SO_ID, CAST([Line item] AS INTEGER) AS [Line item],
        CAST(Qty AS REAL) AS Qty, CAST([Total Sales] AS REAL) AS 출고금액,
        SUBSTR([출고일], 1, 7) AS 출고월
    FROM dn_domestic
    WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
    UNION ALL
    SELECT SO_ID, CAST([Line item] AS INTEGER),
        CAST(Qty AS REAL), CAST([Total Sales KRW] AS REAL),
        SUBSTR([선적일], 1, 7)
    FROM dn_export
    WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
),
dn_by_month AS (
    SELECT SO_ID, [Line item], 출고월,
        SUM(Qty) AS Output_qty, SUM(출고금액) AS Output_amount
    FROM dn_combined WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item], 출고월
),
events_line_item AS (
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        s.Period AS event_period,
        s.[Item qty] AS Value_Input_qty, s.[Sales amount KRW] AS Value_Input_amount,
        0 AS Value_Output_qty, 0 AS Value_Output_amount
    FROM so_combined s
    UNION ALL
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        dm.출고월, 0, 0, dm.Output_qty, dm.Output_amount
    FROM dn_by_month dm
    INNER JOIN so_combined s ON dm.SO_ID = s.SO_ID AND dm.[Line item] = s.[Line item]
),
os_grouped AS (
    SELECT SO_ID, [OS name], [Expected delivery date], event_period AS Period,
        MIN([Customer name]) AS [Customer name], MIN([Customer PO]) AS [Customer PO],
        MIN([Item name]) AS [Item name], MIN(구분) AS 구분, MIN(등록Period) AS 등록Period,
        MIN(Sector) AS Sector,
        MIN([Business registration number]) AS [Business registration number],
        MIN([Industry code]) AS [Industry code],
        GROUP_CONCAT(DISTINCT [AX Period]) AS [AX Period],
        GROUP_CONCAT(DISTINCT [Model code]) AS [Model code],
        SUM(Value_Input_qty) AS Value_Input_qty, SUM(Value_Input_amount) AS Value_Input_amount,
        SUM(Value_Output_qty) AS Value_Output_qty, SUM(Value_Output_amount) AS Value_Output_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date], event_period
)
"""


def _validate_period_format(period: str) -> bool:
    """yyyy-MM 형식 검증"""
    return bool(re.match(r'^\d{4}-(0[1-9]|1[0-2])$', period))


def _next_period(period: str) -> str:
    """다음 월 반환 (yyyy-MM → yyyy-MM)"""
    year = int(period[:4])
    month = int(period[5:7])
    if month == 12:
        return f"{year + 1:04d}-01"
    return f"{year:04d}-{month + 1:02d}"


def _prev_period(period: str) -> str:
    """이전 월 반환"""
    year = int(period[:4])
    month = int(period[5:7])
    if month == 1:
        return f"{year - 1:04d}-12"
    return f"{year:04d}-{month - 1:02d}"


@dataclass
class SnapshotResult:
    """스냅샷 결과"""
    period: str
    success: bool
    message: str
    row_count: int = 0
    variance_rows: int = 0  # Variance != 0인 행 수


class SnapshotEngine:
    """Order Book 스냅샷 엔진 — 월별 마감/Variance 추적"""

    def __init__(self, db_path: Path | None = None):
        self.db_path = db_path or DB_FILE

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA journal_mode=WAL')
        return conn

    def take_snapshot(self, period: str, note: str = '') -> SnapshotResult:
        """월 마감 스냅샷 생성

        1. period 형식 검증
        2. 순차 마감 검증 (이전 period가 마감되었는지)
        3. 롤링 order_book으로 해당 period 결과 추출
        4. Variance 계산 (이전 스냅샷 대비 소급 변경분)
        5. ob_snapshot + ob_snapshot_meta 저장
        """
        if not _validate_period_format(period):
            return SnapshotResult(period, False, f"잘못된 형식: '{period}' (yyyy-MM 필요)")

        conn = self._connect()
        try:
            create_snapshot_tables(conn)

            # 이미 마감된 period인지 확인
            if self.is_period_closed(period, conn):
                return SnapshotResult(period, False, f"'{period}'는 이미 마감되었습니다")

            # 순차 마감 검증
            last_closed = self.get_last_closed_period(conn)
            if last_closed is not None:
                expected_next = _next_period(last_closed)
                if period != expected_next:
                    return SnapshotResult(
                        period, False,
                        f"순차 마감 필요: 마지막 마감='{last_closed}', "
                        f"다음 마감 가능='{expected_next}'"
                    )

            now_iso = datetime.now().isoformat()

            # 이벤트 누적 계산으로 해당 period 시점 결과 추출
            rolling_sql = _ORDER_BOOK_BASE_SQL + """,
                target(p) AS (SELECT ?)
                SELECT
                    og.SO_ID, og.[OS name], og.[Expected delivery date],
                    MIN(og.[Customer name]) AS [Customer name],
                    MIN(og.[Item name]) AS [Item name],
                    MIN(og.구분) AS 구분,
                    MIN(og.등록Period) AS 등록Period,
                    MIN(og.Sector) AS Sector,
                    GROUP_CONCAT(DISTINCT og.[AX Period]) AS [AX Period],
                    GROUP_CONCAT(DISTINCT og.[Model code]) AS [Model code],
                    SUM(CASE WHEN og.Period = t.p THEN og.Value_Input_qty ELSE 0 END) AS Value_Input_qty,
                    SUM(CASE WHEN og.Period = t.p THEN og.Value_Input_amount ELSE 0 END) AS Value_Input_amount,
                    SUM(CASE WHEN og.Period = t.p THEN og.Value_Output_qty ELSE 0 END) AS Value_Output_qty,
                    SUM(CASE WHEN og.Period = t.p THEN og.Value_Output_amount ELSE 0 END) AS Value_Output_amount,
                    SUM(CASE WHEN og.Period < t.p THEN og.Value_Input_qty - og.Value_Output_qty ELSE 0 END) AS Value_Start_qty,
                    SUM(CASE WHEN og.Period < t.p THEN og.Value_Input_amount - og.Value_Output_amount ELSE 0 END) AS Value_Start_amount,
                    SUM(og.Value_Input_qty - og.Value_Output_qty) AS Value_Ending_qty,
                    SUM(og.Value_Input_amount - og.Value_Output_amount) AS Value_Ending_amount
                FROM os_grouped og, target t
                WHERE og.Period <= t.p
                GROUP BY og.SO_ID, og.[OS name], og.[Expected delivery date]
                HAVING ABS(SUM(og.Value_Input_qty - og.Value_Output_qty)) > 0.001
                    OR ABS(SUM(og.Value_Input_amount - og.Value_Output_amount)) > 0.5
                    OR SUM(CASE WHEN og.Period = t.p THEN og.Value_Input_qty + og.Value_Output_qty ELSE 0 END) > 0
            """
            rows = conn.execute(rolling_sql, (period,)).fetchall()

            if not rows:
                return SnapshotResult(period, False, f"'{period}'에 해당하는 Order Book 데이터가 없습니다")

            # Variance 계산
            variance_map: dict[tuple, tuple[float, float]] = {}  # (SO_ID, OS name, EDD) → (var_qty, var_amt)
            variance_count = 0

            if last_closed is not None:
                # 현재 raw 데이터로 이전 마감 period를 재계산 (누적 SUM)
                recalc_sql = _ORDER_BOOK_BASE_SQL + """,
                    target(p) AS (SELECT ?)
                    SELECT
                        og.SO_ID, og.[OS name], og.[Expected delivery date],
                        SUM(og.Value_Input_qty - og.Value_Output_qty) AS Value_Ending_qty,
                        SUM(og.Value_Input_amount - og.Value_Output_amount) AS Value_Ending_amount
                    FROM os_grouped og, target t
                    WHERE og.Period <= t.p
                    GROUP BY og.SO_ID, og.[OS name], og.[Expected delivery date]
                """
                recalc_rows = conn.execute(recalc_sql, (last_closed,)).fetchall()
                recalc_map = {
                    (r['SO_ID'], r['OS name'], r['Expected delivery date'] or ''):
                    (r['Value_Ending_qty'] or 0, r['Value_Ending_amount'] or 0)
                    for r in recalc_rows
                }

                # 이전 스냅샷의 Ending 조회
                snap_rows = conn.execute("""
                    SELECT SO_ID, [OS name], [Expected delivery date],
                           ending_qty, ending_amount
                    FROM ob_snapshot
                    WHERE snapshot_period = ?
                """, (last_closed,)).fetchall()
                snap_map = {
                    (r['SO_ID'], r['OS name'], r['Expected delivery date'] or ''):
                    (r['ending_qty'] or 0, r['ending_amount'] or 0)
                    for r in snap_rows
                }

                # Variance = recalc_ending - snap_ending (모든 키 합집합)
                all_keys = set(recalc_map.keys()) | set(snap_map.keys())
                for key in all_keys:
                    recalc_qty, recalc_amt = recalc_map.get(key, (0, 0))
                    snap_qty, snap_amt = snap_map.get(key, (0, 0))
                    var_qty = recalc_qty - snap_qty
                    var_amt = recalc_amt - snap_amt
                    if abs(var_qty) > 0.001 or abs(var_amt) > 0.5:
                        variance_map[key] = (var_qty, var_amt)
                        variance_count += 1

            # ob_snapshot에 INSERT
            rolling_keys: set[tuple] = set()
            for r in rows:
                key = (r['SO_ID'], r['OS name'], r['Expected delivery date'] or '')
                rolling_keys.add(key)
                var_qty, var_amt = variance_map.get(key, (0, 0))

                # Variance를 반영한 Ending 계산
                # Start = 이전 스냅샷 ending (또는 롤링 Start)
                # Ending = Start + Input + Variance - Output
                start_qty = r['Value_Start_qty'] or 0
                start_amt = r['Value_Start_amount'] or 0
                input_qty = r['Value_Input_qty'] or 0
                input_amt = r['Value_Input_amount'] or 0
                output_qty = r['Value_Output_qty'] or 0
                output_amt = r['Value_Output_amount'] or 0

                if last_closed is not None:
                    # Start는 이전 스냅샷의 ending으로 고정
                    snap_ending = snap_map.get(key, (0, 0))
                    start_qty = snap_ending[0]
                    start_amt = snap_ending[1]
                    ending_qty = start_qty + input_qty + var_qty - output_qty
                    ending_amt = start_amt + input_amt + var_amt - output_amt
                else:
                    # 첫 스냅샷: 롤링 결과 그대로
                    ending_qty = r['Value_Ending_qty'] or 0
                    ending_amt = r['Value_Ending_amount'] or 0

                conn.execute("""
                    INSERT OR REPLACE INTO ob_snapshot (
                        snapshot_period, SO_ID, [OS name], [Expected delivery date],
                        ending_qty, ending_amount, start_qty, start_amount,
                        input_qty, input_amount, output_qty, output_amount,
                        variance_qty, variance_amount,
                        customer_name, item_name, 구분, 등록Period,
                        [AX Period], [Model code], Sector, snapshot_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    period, r['SO_ID'], r['OS name'], r['Expected delivery date'] or '',
                    ending_qty, ending_amt, start_qty, start_amt,
                    input_qty, input_amt, output_qty, output_amt,
                    var_qty, var_amt,
                    r['Customer name'], r['Item name'], r['구분'], r['등록Period'],
                    r['AX Period'], r['Model code'], r['Sector'], now_iso,
                ))

            # 퇴장 행: 전월 스냅샷에 있지만 rolling 결과에 없는 그룹
            # 소급 변경으로 Ending=0이 되어 HAVING 필터에서 제외된 건
            # Start=전월Ending, Variance=소급변경분, Ending≈0 으로 기록
            if last_closed is not None:
                exited_keys = set(snap_map.keys()) - rolling_keys
                for key in exited_keys:
                    snap_qty, snap_amt = snap_map[key]
                    if abs(snap_qty) < 0.001 and abs(snap_amt) < 0.5:
                        continue  # 전월 Ending도 0이면 스킵

                    recalc_qty, recalc_amt = recalc_map.get(key, (0, 0))
                    var_qty = recalc_qty - snap_qty
                    var_amt = recalc_amt - snap_amt
                    ending_qty = snap_qty + var_qty  # = recalc_qty (≈0)
                    ending_amt = snap_amt + var_amt  # = recalc_amt (≈0)

                    # 전월 스냅샷에서 메타데이터 가져오기
                    meta_row = conn.execute("""
                        SELECT customer_name, item_name, 구분, 등록Period,
                               [AX Period], [Model code], Sector
                        FROM ob_snapshot
                        WHERE snapshot_period = ? AND SO_ID = ?
                          AND [OS name] = ? AND [Expected delivery date] = ?
                    """, (last_closed, key[0], key[1], key[2])).fetchone()

                    conn.execute("""
                        INSERT OR REPLACE INTO ob_snapshot (
                            snapshot_period, SO_ID, [OS name], [Expected delivery date],
                            ending_qty, ending_amount, start_qty, start_amount,
                            input_qty, input_amount, output_qty, output_amount,
                            variance_qty, variance_amount,
                            customer_name, item_name, 구분, 등록Period,
                            [AX Period], [Model code], Sector, snapshot_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        period, key[0], key[1], key[2],
                        ending_qty, ending_amt, snap_qty, snap_amt,
                        0, 0, 0, 0,
                        var_qty, var_amt,
                        meta_row['customer_name'] if meta_row else None,
                        meta_row['item_name'] if meta_row else None,
                        meta_row['구분'] if meta_row else None,
                        meta_row['등록Period'] if meta_row else None,
                        meta_row['AX Period'] if meta_row else None,
                        meta_row['Model code'] if meta_row else None,
                        meta_row['Sector'] if meta_row else None,
                        now_iso,
                    ))
                    variance_count += 1

            # ob_snapshot_meta에 INSERT
            conn.execute("""
                INSERT OR REPLACE INTO ob_snapshot_meta (period, closed_at, note, is_active)
                VALUES (?, ?, ?, 1)
            """, (period, now_iso, note or None))

            conn.commit()

            msg = f"'{period}' 마감 완료 ({len(rows)}건)"
            if variance_count > 0:
                msg += f", Variance 감지 {variance_count}건"

            return SnapshotResult(period, True, msg, len(rows), variance_count)

        except Exception as e:
            conn.rollback()
            logger.error("스냅샷 실패: %s", e)
            return SnapshotResult(period, False, f"스냅샷 실패: {e}")
        finally:
            conn.close()

    def undo_snapshot(self, period: str) -> SnapshotResult:
        """최신 마감 취소 (최신 활성 마감만 가능)"""
        if not _validate_period_format(period):
            return SnapshotResult(period, False, f"잘못된 형식: '{period}' (yyyy-MM 필요)")

        conn = self._connect()
        try:
            create_snapshot_tables(conn)

            # 해당 period가 활성 마감인지 확인
            if not self.is_period_closed(period, conn):
                return SnapshotResult(period, False, f"'{period}'는 마감되지 않았습니다")

            # 최신 마감인지 확인
            last_closed = self.get_last_closed_period(conn)
            if last_closed != period:
                return SnapshotResult(
                    period, False,
                    f"최신 마감만 취소 가능합니다 (최신: '{last_closed}')"
                )

            # 스냅샷 데이터 삭제
            cursor = conn.execute(
                "DELETE FROM ob_snapshot WHERE snapshot_period = ?", (period,)
            )
            deleted = cursor.rowcount

            # 메타 비활성화
            conn.execute(
                "UPDATE ob_snapshot_meta SET is_active = 0 WHERE period = ?", (period,)
            )

            conn.commit()
            return SnapshotResult(period, True, f"'{period}' 마감 취소 완료 ({deleted}건 삭제)")

        except Exception as e:
            conn.rollback()
            logger.error("마감 취소 실패: %s", e)
            return SnapshotResult(period, False, f"마감 취소 실패: {e}")
        finally:
            conn.close()

    def list_snapshots(self) -> list[dict]:
        """마감 현황 조회 (Start/Input/Output/Variance/Ending 합계 포함)"""
        conn = self._connect()
        try:
            create_snapshot_tables(conn)
            rows = conn.execute("""
                SELECT m.period, m.closed_at, m.note, m.is_active,
                       COUNT(s.SO_ID) AS row_count,
                       COALESCE(SUM(s.start_amount), 0) AS total_start,
                       COALESCE(SUM(s.input_amount), 0) AS total_input,
                       COALESCE(SUM(s.output_amount), 0) AS total_output,
                       COALESCE(SUM(s.variance_amount), 0) AS total_variance,
                       COALESCE(SUM(s.ending_amount), 0) AS total_ending,
                       SUM(CASE WHEN ABS(s.variance_amount) > 0.5 THEN 1 ELSE 0 END) AS variance_rows
                FROM ob_snapshot_meta m
                LEFT JOIN ob_snapshot s ON m.period = s.snapshot_period
                GROUP BY m.period, m.closed_at, m.note, m.is_active
                ORDER BY m.period
            """).fetchall()

            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_last_closed_period(self, conn: sqlite3.Connection | None = None) -> str | None:
        """마지막 활성 마감 period 조회"""
        close_conn = conn is None
        if close_conn:
            conn = self._connect()
            create_snapshot_tables(conn)
        try:
            row = conn.execute("""
                SELECT MAX(period) AS last_period
                FROM ob_snapshot_meta
                WHERE is_active = 1
            """).fetchone()
            return row['last_period'] if row else None
        finally:
            if close_conn:
                conn.close()

    def is_period_closed(self, period: str, conn: sqlite3.Connection | None = None) -> bool:
        """해당 period가 활성 마감인지 확인"""
        close_conn = conn is None
        if close_conn:
            conn = self._connect()
            create_snapshot_tables(conn)
        try:
            row = conn.execute("""
                SELECT 1 FROM ob_snapshot_meta
                WHERE period = ? AND is_active = 1
            """, (period,)).fetchone()
            return row is not None
        finally:
            if close_conn:
                conn.close()

    def get_status(self) -> dict:
        """현재 상태 요약"""
        conn = self._connect()
        try:
            create_snapshot_tables(conn)
            last_closed = self.get_last_closed_period(conn)
            next_period = _next_period(last_closed) if last_closed else None

            active_count = conn.execute(
                "SELECT COUNT(*) AS cnt FROM ob_snapshot_meta WHERE is_active = 1"
            ).fetchone()['cnt']

            total_snapshot_rows = conn.execute(
                "SELECT COUNT(*) AS cnt FROM ob_snapshot"
            ).fetchone()['cnt']

            return {
                'last_closed_period': last_closed,
                'next_period': next_period,
                'active_snapshots': active_count,
                'total_snapshot_rows': total_snapshot_rows,
            }
        finally:
            conn.close()
