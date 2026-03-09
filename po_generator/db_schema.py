"""
DB 스키마 정의
==============

SQLite 테이블/PK 정의 및 스키마 관리.
시트별 테이블 설정과 DDL 생성을 담당합니다.
"""

from __future__ import annotations

import sqlite3
import logging
from dataclasses import dataclass, field

from po_generator.config import (
    SO_DOMESTIC_SHEET, SO_EXPORT_SHEET,
    PO_DOMESTIC_SHEET, PO_EXPORT_SHEET,
    DN_DOMESTIC_SHEET, DN_EXPORT_SHEET,
    PMT_DOMESTIC_SHEET,
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
        required_column='SO_ID',
        needs_row_seq=True,
        row_seq_group=('PO_ID', 'Line item'),
    ),
    SheetConfig(
        sheet_name=PO_EXPORT_SHEET,
        table_name='po_export',
        pk_columns=('PO_ID', 'Line item', '_row_seq'),
        required_column='SO_ID',
        needs_row_seq=True,
        row_seq_group=('PO_ID', 'Line item'),
    ),
    SheetConfig(
        sheet_name=DN_DOMESTIC_SHEET,
        table_name='dn_domestic',
        pk_columns=('DN_ID', 'Line item'),
        required_column='DN_ID',
    ),
    SheetConfig(
        sheet_name=DN_EXPORT_SHEET,
        table_name='dn_export',
        pk_columns=('DN_ID', 'Line item'),
        required_column='DN_ID',
    ),
    SheetConfig(
        sheet_name=PMT_DOMESTIC_SHEET,
        table_name='pmt_domestic',
        pk_columns=('\uc120\uc218\uae08_ID',),  # 선수금_ID
        required_column='\uc120\uc218\uae08_ID',  # 선수금_ID
    ),
]


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
    conn.execute(f'DROP TABLE [{config.table_name}]')
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
            [AX Project number] TEXT,
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
