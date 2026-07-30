"""
DB 스키마 정의
==============

SQLite 테이블/PK 정의 및 스키마 관리.
시트별 테이블 설정과 DDL 생성을 담당합니다.
"""

from __future__ import annotations

import os
import socket
import sqlite3
import logging
from dataclasses import dataclass, field
from datetime import datetime

from po_generator.config import (
    SO_DOMESTIC_SHEET, SO_EXPORT_SHEET,
    PO_DOMESTIC_SHEET, PO_EXPORT_SHEET,
    DN_DOMESTIC_SHEET, DN_EXPORT_SHEET,
    PMT_DOMESTIC_SHEET, FX_SHEET,
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class SheetConfig:
    """시트별 동기화 설정"""
    sheet_name: str          # Excel 시트명
    table_name: str          # SQLite 테이블명
    pk_columns: tuple[str, ...]  # PK 컬럼 (복합키 지원)
    required_column: str     # NaN이면 행 스킵 (빈 행 필터링)
    needs_row_seq: bool = False  # _row_seq 자동 생성 여부
    row_seq_group: tuple[str, ...] = field(default_factory=tuple)  # _row_seq 그룹핑 컬럼


# 7개 시트 설정
SYNC_SHEETS: list[SheetConfig] = [
    SheetConfig(
        sheet_name=SO_DOMESTIC_SHEET,
        table_name='so_domestic',
        pk_columns=('SO_ID', 'Line item'),
        required_column='SO_ID',
    ),
    SheetConfig(
        sheet_name=SO_EXPORT_SHEET,
        table_name='so_export',
        pk_columns=('SO_ID', 'Line item'),
        required_column='SO_ID',
    ),
    SheetConfig(
        sheet_name=PO_DOMESTIC_SHEET,
        table_name='po_domestic',
        pk_columns=('PO_ID', 'Line item', '_row_seq'),
        required_column='PO_ID',
        needs_row_seq=True,
        row_seq_group=('PO_ID', 'Line item'),
    ),
    SheetConfig(
        sheet_name=PO_EXPORT_SHEET,
        table_name='po_export',
        pk_columns=('PO_ID', 'Line item', '_row_seq'),
        required_column='PO_ID',
        needs_row_seq=True,
        row_seq_group=('PO_ID', 'Line item'),
    ),
    # DN도 (DN_ID, SO_ID, Line item)이 유일하지 않다 — 같은 DN 문서·같은 SO 라인을
    # 두 행으로 나눠 적는 분할출고가 실제로 존재한다(2026-07-30 실측 2쌍). _row_seq 없이는
    # 뒤 행이 앞 행을 덮어써 매출·수량이 조용히 사라진다. PO와 동일한 처방.
    SheetConfig(
        sheet_name=DN_DOMESTIC_SHEET,
        table_name='dn_domestic',
        pk_columns=('DN_ID', 'SO_ID', 'Line item', '_row_seq'),
        required_column='DN_ID',
        needs_row_seq=True,
        row_seq_group=('DN_ID', 'SO_ID', 'Line item'),
    ),
    SheetConfig(
        sheet_name=DN_EXPORT_SHEET,
        table_name='dn_export',
        pk_columns=('DN_ID', 'SO_ID', 'Line item', '_row_seq'),
        required_column='DN_ID',
        needs_row_seq=True,
        row_seq_group=('DN_ID', 'SO_ID', 'Line item'),
    ),
    SheetConfig(
        sheet_name=PMT_DOMESTIC_SHEET,
        table_name='pmt_domestic',
        pk_columns=('\uc120\uc218\uae08_ID',),  # 선수금_ID
        required_column='\uc120\uc218\uae08_ID',  # 선수금_ID
    ),
]

# sync가 아는 시트 전체 = 세로형 공통 파이프라인(SYNC_SHEETS) + 가로형 전용 경로(FX).
# "--sheets에 설정에 없는 시트가 왔다" 경고 판정의 단일 소유자 — 특수 경로가 늘면 여기만 넓힌다.
KNOWN_SYNC_SHEETS: frozenset[str] = frozenset(
    {c.sheet_name for c in SYNC_SHEETS} | {FX_SHEET}
)

# _sync_log.change_type 어휘 — writer(sync_db)와 reader(dashboard 동기화 로그)가 공유.
# 멤버를 추가하면 양쪽이 자동으로 따라온다 ('재적재'가 writer에만 있던 갈라짐 방지).
SYNC_LOG_CHANGE_TYPES: tuple[str, ...] = ('신규', '수정', '삭제', '재적재')


def _get_table_pk(conn: sqlite3.Connection, table_name: str) -> tuple[str, ...]:
    """기존 테이블의 PK 컬럼 조회. 테이블이 없으면 빈 튜플."""
    try:
        cursor = conn.execute(f'PRAGMA table_info([{table_name}])')
        rows = cursor.fetchall()
        # table_info: (cid, name, type, notnull, dflt_value, pk)
        pk_cols = [(r[5], r[1]) for r in rows if r[5] > 0]
        pk_cols.sort()  # pk 순번 기준 정렬
        return tuple(c[1] for c in pk_cols)
    except sqlite3.OperationalError:
        return ()


def migrate_pk_if_changed(conn: sqlite3.Connection, config: 'SheetConfig') -> bool:
    """테이블 PK가 설정과 다르면 DROP 후 재생성 유도. 변경 여부 반환.

    DB는 Excel 백업이므로, PK 변경 시 테이블을 삭제해도
    다음 sync에서 전체 데이터가 다시 INSERT됩니다.
    """
    existing_pk = _get_table_pk(conn, config.table_name)
    if not existing_pk:
        return False  # 테이블 없음 → 마이그레이션 불필요

    if existing_pk == config.pk_columns:
        return False  # PK 동일

    logger.info(
        "%s: PK 변경 감지 (%s → %s), 테이블 재생성",
        config.table_name, existing_pk, config.pk_columns,
    )
    # 기존 테이블 백업 후 삭제 (스냅샷 등 파생 데이터 복구 가능)
    backup = f"{config.table_name}_bak"
    conn.execute(f'DROP TABLE IF EXISTS [{backup}]')
    conn.execute(f'ALTER TABLE [{config.table_name}] RENAME TO [{backup}]')
    logger.info("%s → %s 백업 완료", config.table_name, backup)
    return True


def _sanitize_col_name(col: str) -> str:
    """컬럼명을 SQLite 안전한 식별자로 변환.

    대괄호 이스케이프를 사용하므로 대부분의 문자열이 그대로 사용 가능.
    """
    return col.strip()


def create_table(conn: sqlite3.Connection, table_name: str,
                 columns: list[str], pk_columns: tuple[str, ...]) -> None:
    """테이블 생성 (없으면 생성, 있으면 무시)"""
    col_defs = []
    for col in columns:
        safe = _sanitize_col_name(col)
        col_defs.append(f'[{safe}] TEXT')

    pk_list = ', '.join(f'[{_sanitize_col_name(c)}]' for c in pk_columns)
    col_defs_str = ',\n  '.join(col_defs)

    # _sync_updated_at: 마지막 동기화 시각
    sql = f"""CREATE TABLE IF NOT EXISTS [{table_name}] (
  {col_defs_str},
  [_sync_updated_at] TEXT,
  PRIMARY KEY ({pk_list})
)"""
    conn.execute(sql)
    logger.debug("테이블 생성/확인: %s (PK: %s)", table_name, pk_columns)


def ensure_columns_exist(conn: sqlite3.Connection, table_name: str,
                         new_columns: list[str]) -> int:
    """기존 테이블에 없는 컬럼 추가. 추가된 컬럼 수 반환."""
    cursor = conn.execute(f'PRAGMA table_info([{table_name}])')
    existing = {row[1] for row in cursor.fetchall()}

    added = 0
    for col in new_columns:
        safe = _sanitize_col_name(col)
        if safe not in existing and safe != '_sync_updated_at':
            conn.execute(f'ALTER TABLE [{table_name}] ADD COLUMN [{safe}] TEXT')
            logger.debug("컬럼 추가: %s.[%s]", table_name, safe)
            added += 1

    return added


def get_table_row_count(conn: sqlite3.Connection, table_name: str) -> int:
    """테이블 행 수 조회"""
    try:
        cursor = conn.execute(f'SELECT COUNT(*) FROM [{table_name}]')
        return cursor.fetchone()[0]
    except sqlite3.OperationalError:
        return 0


def create_snapshot_tables(conn: sqlite3.Connection) -> None:
    """Order Book 스냅샷 테이블 생성 (없으면 생성)"""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ob_snapshot (
            snapshot_period TEXT NOT NULL,
            SO_ID TEXT NOT NULL,
            [OS name] TEXT NOT NULL,
            [Expected delivery date] TEXT NOT NULL DEFAULT '',
            ending_qty REAL NOT NULL DEFAULT 0,
            ending_amount REAL NOT NULL DEFAULT 0,
            start_qty REAL NOT NULL DEFAULT 0,
            start_amount REAL NOT NULL DEFAULT 0,
            input_qty REAL NOT NULL DEFAULT 0,
            input_amount REAL NOT NULL DEFAULT 0,
            output_qty REAL NOT NULL DEFAULT 0,
            output_amount REAL NOT NULL DEFAULT 0,
            variance_qty REAL NOT NULL DEFAULT 0,
            variance_amount REAL NOT NULL DEFAULT 0,
            customer_name TEXT,
            item_name TEXT,
            구분 TEXT,
            등록Period TEXT,
            [AX Period] TEXT,
            [Model code] TEXT,
            Sector TEXT,
            snapshot_at TEXT NOT NULL,
            PRIMARY KEY (snapshot_period, SO_ID, [OS name], [Expected delivery date])
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ob_snapshot_meta (
            period TEXT PRIMARY KEY,
            closed_at TEXT NOT NULL,
            note TEXT,
            is_active INTEGER NOT NULL DEFAULT 1
        )
    """)
    logger.debug("스냅샷 테이블 생성/확인 완료")


def create_fx_table(conn: sqlite3.Connection) -> None:
    """월별 환율 테이블 생성 (FX 시트 언피벗 결과: 통화 × 월 → 환율)"""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS fx (
            currency TEXT NOT NULL,
            ym TEXT NOT NULL,
            rate REAL NOT NULL,
            _sync_updated_at TEXT,
            PRIMARY KEY (currency, ym)
        )
    """)
    logger.debug("fx 테이블 생성/확인 완료")


# ─────────────────────────────────────────────────────────────
# v_dn_revenue — DN 매출 인식(월·금액) 단일 정의
#
# Order Book 계열 쿼리가 6곳(sql/ 4개 + snapshot.py + dashboard.py)에서
# 같은 DN 산식을 복제하고 있었다. 뷰로 한 번만 정의해 갈라짐을 막는다.
# Power Query `Order_Book` / `AX_매출대사`와 동일 규칙:
#   국내 매출인식일 = 세금계산서 발행일
#                   → 'N/A'(발행 불필요: 무상공급·FOC·반품)면 출고일
#                   → 선수금 세금계산서 + 출고 완료면 출고일
#                   → 없으면 NULL (매출 미인식 → Output 없음 = Backlog 잔류)
#   해외 매출인식일 = 선적일, KRW = 외화금액 × 선적월 환율 (없으면 시트 KRW로 폴백)
#   fx_variance     = 재환산액 − 시트 KRW (환율 재평가분, 국내는 항상 0)
# ─────────────────────────────────────────────────────────────
DN_REVENUE_VIEW_SQL = """
CREATE VIEW v_dn_revenue AS
SELECT
    DN_ID, SO_ID, [Line item], Qty, 매출월, 출고일,
    output_amount,
    output_amount - sheet_amount AS fx_variance,
    dn_cust, dn_item, dn_po, dn_brn, dn_market
FROM (
    -- ─── 국내: 매출인식일 기준 (KRW 거래 → 재환산 없음) ───
    SELECT
        d.DN_ID,
        d.SO_ID,
        CAST(d.[Line item] AS INTEGER) AS [Line item],
        CAST(d.Qty AS REAL)            AS Qty,
        SUBSTR(
            CASE
                WHEN TRIM(COALESCE(d.[세금계산서 발행일], '')) GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]*'
                    THEN d.[세금계산서 발행일]
                WHEN UPPER(TRIM(COALESCE(d.[세금계산서 발행일], ''))) = 'N/A'
                    THEN d.[출고일]
                WHEN TRIM(COALESCE(d.[선수금 세금계산서 발행일], '')) != ''
                     AND TRIM(COALESCE(d.[출고일], '')) != ''
                    THEN d.[출고일]
            END, 1, 7)                 AS 매출월,
        d.[출고일]                      AS 출고일,
        ROUND(CAST(d.[Total Sales] AS REAL)) AS output_amount,
        ROUND(CAST(d.[Total Sales] AS REAL)) AS sheet_amount,
        d.[Customer name] AS dn_cust, d.[Item] AS dn_item,
        d.[Customer PO] AS dn_po, d.[Business registration number] AS dn_brn,
        '국내' AS dn_market
    FROM dn_domestic d

    UNION ALL

    -- ─── 해외: 선적일 기준 + 선적월 환율 재환산 ───
    SELECT
        e.DN_ID,
        e.SO_ID,
        CAST(e.[Line item] AS INTEGER),
        CAST(e.Qty AS REAL),
        SUBSTR(e.[선적일], 1, 7),
        e.[출고일],
        CASE
            WHEN e.Currency = 'KRW' THEN ROUND(CAST(e.[Total Sales] AS REAL))
            WHEN f.rate IS NOT NULL AND TRIM(COALESCE(e.[Total Sales], '')) != ''
                THEN ROUND(CAST(e.[Total Sales] AS REAL) * f.rate)
            ELSE ROUND(CAST(e.[Total Sales KRW] AS REAL))
        END,
        ROUND(CAST(e.[Total Sales KRW] AS REAL)),
        e.[Customer name], e.[Item], e.[Customer PO], '', '해외'
    FROM dn_export e
    LEFT JOIN fx f
        ON f.currency = e.Currency
       AND f.ym = SUBSTR(e.[선적일], 1, 7)
)
"""

# v_dn_by_month — 매출 인식된 라인의 월별 집계 (Order Book Output의 공통 재료)
# "미인식 행은 Output이 아니다" 필터와 SUM 집계를 한 곳에 둔다 — 소비처(sql 4종 +
# snapshot.py + dashboard.py)가 같은 WHERE/GROUP BY를 각자 들고 있지 않게.
# 미인식 행 진단(매출월 IS NULL 조회)은 라인 레벨 v_dn_revenue를 직접 읽는다.
DN_BY_MONTH_VIEW_SQL = """
CREATE VIEW v_dn_by_month AS
SELECT SO_ID, [Line item], 매출월,
       SUM(Qty)           AS Output_qty,
       SUM(output_amount) AS Output_amount,
       SUM(fx_variance)   AS Output_fx,
       MIN(dn_cust) AS dn_cust, MIN(dn_item) AS dn_item,
       MIN(dn_po) AS dn_po, MIN(dn_brn) AS dn_brn, MIN(dn_market) AS dn_market
FROM v_dn_revenue
WHERE 매출월 IS NOT NULL AND 매출월 != ''
GROUP BY SO_ID, [Line item], 매출월
"""


def create_order_book_views(conn: sqlite3.Connection) -> None:
    """Order Book 계열 공통 뷰 생성/갱신 (정의 변경이 바로 반영되도록 DROP 후 재생성)"""
    create_fx_table(conn)  # 뷰가 참조하므로 먼저 보장
    conn.execute('DROP VIEW IF EXISTS v_dn_by_month')
    conn.execute('DROP VIEW IF EXISTS v_dn_revenue')
    conn.execute(DN_REVENUE_VIEW_SQL)
    conn.execute(DN_BY_MONTH_VIEW_SQL)
    logger.debug("v_dn_revenue / v_dn_by_month 뷰 생성/갱신 완료")


def ensure_sync_log_tables(conn: sqlite3.Connection) -> None:
    """_sync_runs + _sync_log v2 스키마 생성 (idempotent).

    구조:
    - _sync_runs : 동기화 세션 메타 (sync_id, started/ended_at, actor, host, dry_run, total_changes)
    - _sync_log  : 변경 이벤트 (sync_id FK, sheet, type, pk_json, pk_display, changes_json, row_snapshot_json)
                   record(레코드)당 1행. 신규/수정/삭제 정보는 changes_json 또는 row_snapshot_json으로 저장.
    """
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _sync_runs (
            sync_id        INTEGER PRIMARY KEY AUTOINCREMENT,
            started_at     TEXT NOT NULL,
            ended_at       TEXT,
            actor          TEXT,
            host           TEXT,
            dry_run        INTEGER NOT NULL DEFAULT 0,
            total_changes  INTEGER NOT NULL DEFAULT 0,
            note           TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _sync_log (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            sync_id           INTEGER NOT NULL,
            sheet_name        TEXT NOT NULL,
            change_type       TEXT NOT NULL,
            pk_json           TEXT NOT NULL,
            pk_display        TEXT NOT NULL,
            changes_json      TEXT,
            row_snapshot_json TEXT,
            FOREIGN KEY (sync_id) REFERENCES _sync_runs(sync_id)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sync_log_sync_id ON _sync_log (sync_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sync_log_sheet   ON _sync_log (sheet_name, sync_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sync_log_pk      ON _sync_log (pk_display)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sync_runs_started ON _sync_runs (started_at)")


def ensure_so_change_ack_table(conn: sqlite3.Connection) -> None:
    """SO 시트 무단 단가/수량 변경 확인(ack) 테이블 — idempotent."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _so_change_ack (
            sync_log_id INTEGER PRIMARY KEY,
            acked_at    TEXT NOT NULL,
            acked_by    TEXT,
            note        TEXT,
            FOREIGN KEY (sync_log_id) REFERENCES _sync_log(id)
        )
    """)


# 하위 호환용 별칭 — 기존 호출자(있다면)가 깨지지 않도록
ensure_sync_log_table = ensure_sync_log_tables


def _resolve_actor() -> str | None:
    """실행 사용자 식별 — Windows/Unix 어디서든 동작."""
    for env in ("USERNAME", "USER", "LOGNAME"):
        v = os.environ.get(env)
        if v:
            return v
    try:
        return os.getlogin()
    except OSError:
        return None


def create_sync_run(conn: sqlite3.Connection, dry_run: bool = False,
                    note: str | None = None,
                    started_at: str | None = None) -> int:
    """동기화 세션 시작 → _sync_runs 행 INSERT 후 sync_id 반환.

    started_at: 실제 동기화 시작시각(SyncSummary.started_at)을 넘기면 그대로 기록.
        None이면 호출 시각으로 대체. 로그 적재가 sync 종료 직후 일어나므로,
        넘기지 않으면 '로그쓰기 시점'이 기록되어 세션 시작/종료 구간이 왜곡된다.
    dry_run: 운영 경로에서는 항상 False — dry-run 동기화는 롤백되므로
        _sync_log/_sync_runs에 기록하지 않는다(유령 변경 방지). 컬럼은 과거
        마이그레이션 데이터 호환 목적으로만 유지된다.
    """
    started_at = started_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    actor = _resolve_actor()
    host = socket.gethostname()
    cur = conn.execute(
        "INSERT INTO _sync_runs (started_at, actor, host, dry_run, total_changes, note) "
        "VALUES (?, ?, ?, ?, 0, ?)",
        (started_at, actor, host, int(bool(dry_run)), note),
    )
    return cur.lastrowid


def finalize_sync_run(conn: sqlite3.Connection, sync_id: int,
                      total_changes: int) -> None:
    """동기화 세션 종료 — ended_at + total_changes 갱신."""
    ended_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE _sync_runs SET ended_at=?, total_changes=? WHERE sync_id=?",
        (ended_at, total_changes, sync_id),
    )


def get_sync_metadata(conn: sqlite3.Connection) -> dict[str, dict]:
    """_sync_meta 테이블에서 동기화 메타정보 조회"""
    try:
        cursor = conn.execute('SELECT table_name, last_sync, row_count FROM _sync_meta')
        return {
            row[0]: {'last_sync': row[1], 'row_count': row[2]}
            for row in cursor.fetchall()
        }
    except sqlite3.OperationalError:
        return {}


def update_sync_metadata(conn: sqlite3.Connection, table_name: str,
                         sync_time: str, row_count: int) -> None:
    """동기화 메타정보 업데이트"""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _sync_meta (
            table_name TEXT PRIMARY KEY,
            last_sync TEXT,
            row_count INTEGER
        )
    """)
    conn.execute("""
        INSERT INTO _sync_meta (table_name, last_sync, row_count)
        VALUES (?, ?, ?)
        ON CONFLICT(table_name) DO UPDATE SET
            last_sync = excluded.last_sync,
            row_count = excluded.row_count
    """, (table_name, sync_time, row_count))
