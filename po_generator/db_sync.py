"""
Excel → SQLite 동기화 엔진
===========================

NOAH_SO_PO_DN.xlsx의 수동 입력 시트를 SQLite DB에 upsert 방식으로 동기화합니다.
"""

from __future__ import annotations

import sqlite3
import logging
import math
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from po_generator.config import NOAH_SO_PO_DN_FILE, DB_FILE
from po_generator.db_schema import (
    SheetConfig, SYNC_SHEETS,
    create_table, ensure_columns_exist,
    update_sync_metadata, get_table_row_count,
    migrate_pk_if_changed,
)

logger = logging.getLogger(__name__)


@dataclass
class SheetSyncResult:
    """시트별 동기화 결과"""
    sheet_name: str
    table_name: str
    total_rows: int = 0
    inserted: int = 0
    updated: int = 0
    skipped: int = 0
    errors: int = 0
    error_messages: list[str] = field(default_factory=list)
    unchanged: int = 0
    inserted_pks: list[tuple] = field(default_factory=list)
    updated_pks: list[tuple] = field(default_factory=list)
    # 신규 상세: [{pk: tuple, values: {col: value}}] — 비어있지 않은 값만
    inserted_details: list[dict] = field(default_factory=list)
    # 수정 상세: [{pk: tuple, changes: {col: (old, new)}}]
    updated_details: list[dict] = field(default_factory=list)
    # 삭제(prune): Excel에서 제거된 행
    pruned: int = 0
    pruned_pks: list[tuple] = field(default_factory=list)
    # 삭제 직전 행 스냅샷 — 감사/복구용 ([{pk: tuple, snapshot: {col: value}}])
    pruned_snapshots: list[dict] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return self.errors == 0


@dataclass
class SyncSummary:
    """전체 동기화 결과 요약"""
    results: list[SheetSyncResult] = field(default_factory=list)
    elapsed_seconds: float = 0.0
    source_file: str = ''
    db_file: str = ''

    @property
    def total_rows(self) -> int:
        return sum(r.total_rows for r in self.results)

    @property
    def total_inserted(self) -> int:
        return sum(r.inserted for r in self.results)

    @property
    def total_updated(self) -> int:
        return sum(r.updated for r in self.results)

    @property
    def total_pruned(self) -> int:
        return sum(r.pruned for r in self.results)

    @property
    def total_errors(self) -> int:
        return sum(r.errors for r in self.results)


def _sanitize_value(val):
    """pandas/numpy 값을 SQLite 호환 Python 타입으로 변환"""
    if val is None:
        return None
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        f = float(val)
        return None if math.isnan(f) or math.isinf(f) else f
    if isinstance(val, np.bool_):
        return bool(val)
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.isoformat()
    if isinstance(val, np.ndarray):
        return str(val.tolist())
    if pd.isna(val):
        return None
    return val


def _add_row_seq(df: pd.DataFrame, group_cols: tuple[str, ...]) -> pd.DataFrame:
    """그룹 내 순번(_row_seq) 부여. Excel 행 순서 기준."""
    existing = [c for c in group_cols if c in df.columns]
    if not existing:
        df['_row_seq'] = 1
        return df
    df = df.copy()
    df['_row_seq'] = df.groupby(existing, sort=False).cumcount() + 1
    return df


def _normalize_pk(pk: tuple) -> tuple:
    """PK 값을 문자열 튜플로 정규화 — Python set 비교 시 타입 불일치 방지."""
    return tuple('' if v is None else str(v) for v in pk)


class SyncEngine:
    """Excel → SQLite 동기화 엔진"""

    def __init__(self, excel_path: Path | None = None, db_path: Path | None = None):
        self.excel_path = excel_path or NOAH_SO_PO_DN_FILE
        self.db_path = db_path or DB_FILE

    def sync_all(self, dry_run: bool = False,
                 sheet_filter: list[str] | None = None) -> SyncSummary:
        """전체 시트 동기화.

        Args:
            dry_run: True면 실제 DB 변경 없이 시뮬레이션만 수행
            sheet_filter: 동기화할 시트명 리스트 (None이면 전체)

        Returns:
            SyncSummary: 동기화 결과 요약
        """
        start = datetime.now()
        summary = SyncSummary(
            source_file=self.excel_path.name,
            db_file=self.db_path.name,
        )

        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel 파일을 찾을 수 없습니다: {self.excel_path}")

        # Excel 파일 한 번만 오픈
        logger.info("Excel 파일 로딩: %s", self.excel_path.name)
        xls = pd.ExcelFile(self.excel_path)
        available_sheets = set(xls.sheet_names)

        # 대상 시트 필터링
        configs = SYNC_SHEETS
        if sheet_filter:
            filter_set = set(sheet_filter)
            configs = [c for c in configs if c.sheet_name in filter_set]
            not_found = filter_set - {c.sheet_name for c in configs}
            if not_found:
                logger.warning("설정에 없는 시트 무시: %s", not_found)

        # DB 연결 — dry-run도 실제 DB에 연결하여 정확한 diff 산출 후 롤백
        # isolation_level=None → 수동 트랜잭션 제어 (DDL 암묵적 COMMIT 방지)
        conn = sqlite3.connect(str(self.db_path), isolation_level=None)
        try:
            conn.execute('PRAGMA journal_mode=WAL')
            conn.execute('PRAGMA synchronous=NORMAL')
            conn.execute('BEGIN')

            for config in configs:
                if config.sheet_name not in available_sheets:
                    logger.warning("시트 없음, 스킵: %s", config.sheet_name)
                    result = SheetSyncResult(
                        sheet_name=config.sheet_name,
                        table_name=config.table_name,
                    )
                    result.error_messages.append(f"시트 '{config.sheet_name}' 없음")
                    summary.results.append(result)
                    continue

                result = self._sync_sheet(conn, xls, config, dry_run)
                summary.results.append(result)

            if dry_run:
                conn.rollback()
            elif summary.total_errors > 0:
                conn.rollback()
                logger.error(
                    "동기화 중단: %d건 에러 발생 → ROLLBACK (데이터 수정 후 재시도 필요)",
                    summary.total_errors,
                )
            else:
                conn.commit()
        finally:
            conn.close()
            xls.close()

        summary.elapsed_seconds = (datetime.now() - start).total_seconds()
        return summary

    def _sync_sheet(self, conn: sqlite3.Connection, xls: pd.ExcelFile,
                    config: SheetConfig, dry_run: bool) -> SheetSyncResult:
        """단일 시트 동기화"""
        result = SheetSyncResult(
            sheet_name=config.sheet_name,
            table_name=config.table_name,
        )

        try:
            # 1. DataFrame 로드
            df = pd.read_excel(xls, sheet_name=config.sheet_name, dtype=str,
                                 keep_default_na=False, na_values=[''])
            df.columns = [str(c).strip() for c in df.columns]

            # 2. 필수 컬럼 NaN인 행 제거 (빈 행 필터링)
            if config.required_column not in df.columns:
                msg = f"필수 컬럼 '{config.required_column}'이 시트에 없습니다"
                logger.error("%s: %s", config.sheet_name, msg)
                result.error_messages.append(msg)
                result.errors = 1
                return result

            df = df.dropna(subset=[config.required_column])
            df = df[df[config.required_column].str.strip() != '']
            result.total_rows = len(df)

            if result.total_rows == 0:
                logger.info("%s: 데이터 없음 — prune 확인", config.sheet_name)
                # 시트가 비었어도 DB 테이블에 잔류 행이 있으면 prune
                try:
                    row_count = get_table_row_count(conn, config.table_name)
                except Exception:
                    row_count = 0
                if row_count > 0:
                    safe_pk_cols = [f'[{c}]' for c in config.pk_columns]
                    # 전체 컬럼 + PK 함께 조회해서 스냅샷 확보
                    all_cols_info = conn.execute(
                        f'PRAGMA table_info([{config.table_name}])'
                    ).fetchall()
                    snapshot_cols = [c[1] for c in all_cols_info if c[1] != '_sync_updated_at']
                    safe_snap_cols = [f'[{c}]' for c in snapshot_cols]
                    db_rows = conn.execute(
                        f'SELECT {", ".join(safe_snap_cols)} FROM [{config.table_name}]'
                    ).fetchall()
                    stale_pks = []
                    snapshots = []
                    pk_idx = [snapshot_cols.index(c) for c in config.pk_columns if c in snapshot_cols]
                    for row in db_rows:
                        pk_tuple = _normalize_pk(tuple(row[i] for i in pk_idx))
                        stale_pks.append(pk_tuple)
                        snap = {col: row[i] for i, col in enumerate(snapshot_cols)
                                if row[i] is not None and str(row[i]) != ''}
                        snapshots.append({'pk': pk_tuple, 'snapshot': snap})
                    if stale_pks:
                        pk_placeholders = ' AND '.join(
                            f'[{c}] = ?' for c in config.pk_columns
                        )
                        for stale_pk in stale_pks:
                            conn.execute(
                                f'DELETE FROM [{config.table_name}] WHERE {pk_placeholders}',
                                list(stale_pk),
                            )
                    result.pruned = len(stale_pks)
                    result.pruned_pks = stale_pks
                    result.pruned_snapshots = snapshots
                    if stale_pks:
                        logger.info(
                            "%s: %d행 삭제(prune) — 시트 전체 비어있음",
                            config.sheet_name, result.pruned,
                        )
                return result

            # 3. _row_seq 생성 (필요한 시트만)
            if config.needs_row_seq:
                df = _add_row_seq(df, config.row_seq_group)

            # 4. 컬럼 목록 구성
            columns = list(df.columns)

            # 5. PK 변경 시 테이블 재생성 + 테이블 생성/컬럼 추가
            migrate_pk_if_changed(conn, config)
            create_table(conn, config.table_name, columns, config.pk_columns)
            added_cols = ensure_columns_exist(conn, config.table_name, columns)
            if added_cols > 0:
                logger.info("%s: %d개 새 컬럼 추가", config.sheet_name, added_cols)

            # 6. PK 기반 기존 데이터 조회
            pk_cols = config.pk_columns
            pk_placeholders = ' AND '.join(
                f'[{c}] = ?' for c in pk_cols
            )

            now_iso = datetime.now().isoformat()
            safe_cols = [f'[{c.strip()}]' for c in columns]

            # 7. Upsert 수행
            excel_pks: set[tuple] = set()
            for idx, row in df.iterrows():
                try:
                    # PK 값 추출 (required_column만 필수, 나머지 PK는 빈 문자열 허용)
                    pk_vals = []
                    skip = False
                    for pk_col in pk_cols:
                        val = _sanitize_value(row.get(pk_col))
                        if val is None or (isinstance(val, str) and val.strip() == ''):
                            if pk_col == config.required_column:
                                skip = True
                                break
                            val = ''  # 비필수 PK는 빈 문자열로 치환
                        pk_vals.append(val)

                    if skip:
                        result.skipped += 1
                        continue

                    # _row_seq 제외 원본 PK가 전부 빈값 → 빈 행으로 간주
                    real_pk_vals = [
                        v for v, c in zip(pk_vals, pk_cols) if c != '_row_seq'
                    ]
                    if real_pk_vals and all(v == '' for v in real_pk_vals):
                        result.skipped += 1
                        logger.debug(
                            "%s - 행 %d: PK 전부 빈값 → 스킵", config.sheet_name, idx,
                        )
                        continue

                    excel_pks.add(_normalize_pk(tuple(pk_vals)))

                    # 기존 행 조회 (전체 컬럼)
                    select_cols = ', '.join(safe_cols)
                    cursor = conn.execute(
                        f'SELECT {select_cols} FROM [{config.table_name}] WHERE {pk_placeholders}',
                        pk_vals,
                    )
                    existing_row = cursor.fetchone()

                    # 새 값 준비 (PK 컬럼은 None→'' 통일 — SQLite에서 NULL은 PK 비교 불가)
                    pk_set = set(pk_cols)
                    new_values = []
                    for c in columns:
                        val = _sanitize_value(row.get(c))
                        if c in pk_set and (val is None or (isinstance(val, str) and val.strip() == '')):
                            val = ''
                        new_values.append(val)

                    if existing_row is not None:
                        # 변경된 필드 감지
                        changes = {}
                        for i, col in enumerate(columns):
                            old_val = existing_row[i]
                            new_val = new_values[i]
                            # 둘 다 None/빈문자열이면 같은 것으로 취급
                            old_norm = None if old_val in (None, '', 'None') else str(old_val)
                            new_norm = None if new_val in (None, '', 'None') else str(new_val)
                            if old_norm != new_norm:
                                changes[col] = (old_val, new_val)

                        if not changes:
                            result.unchanged += 1
                            continue

                        # UPDATE (변경분 있을 때만)
                        values_with_ts = new_values + [now_iso]
                        set_clause = ', '.join(
                            f'{sc} = ?' for sc in safe_cols
                        )
                        conn.execute(
                            f'UPDATE [{config.table_name}] SET {set_clause}, '
                            f'[_sync_updated_at] = ? WHERE {pk_placeholders}',
                            values_with_ts + pk_vals,
                        )
                        result.updated += 1
                        result.updated_pks.append(tuple(pk_vals))
                        result.updated_details.append({
                            'pk': tuple(pk_vals),
                            'changes': changes,
                        })
                    else:
                        # INSERT
                        all_cols = safe_cols + ['[_sync_updated_at]']
                        placeholders = ', '.join('?' for _ in all_cols)
                        conn.execute(
                            f'INSERT INTO [{config.table_name}] '
                            f'({", ".join(all_cols)}) VALUES ({placeholders})',
                            new_values + [now_iso],
                        )
                        result.inserted += 1
                        result.inserted_pks.append(tuple(pk_vals))
                        # 신규 행의 비어있지 않은 값 기록
                        non_empty = {}
                        for i, c in enumerate(columns):
                            v = new_values[i]
                            if v is not None and str(v).strip() != '':
                                non_empty[c] = v
                        result.inserted_details.append({
                            'pk': tuple(pk_vals),
                            'values': non_empty,
                        })

                except Exception as e:
                    result.errors += 1
                    msg = f"행 {idx}: {e}"
                    result.error_messages.append(msg)
                    logger.warning("%s - %s", config.sheet_name, msg)

            # 8. Prune: Excel에서 삭제된 행 제거 (스냅샷 캡처 후 DELETE)
            safe_pk_cols = [f'[{c}]' for c in pk_cols]
            db_pks_cursor = conn.execute(
                f'SELECT {", ".join(safe_pk_cols)} FROM [{config.table_name}]'
            )
            db_pks = {_normalize_pk(row) for row in db_pks_cursor.fetchall()}
            stale_pks = db_pks - excel_pks
            if stale_pks:
                for stale_pk in stale_pks:
                    # 삭제 직전 행 스냅샷 캡처 (감사/복구용)
                    snap_row = conn.execute(
                        f'SELECT {", ".join(safe_cols)} FROM [{config.table_name}] '
                        f'WHERE {pk_placeholders}',
                        list(stale_pk),
                    ).fetchone()
                    if snap_row is not None:
                        snap = {col: snap_row[i] for i, col in enumerate(columns)
                                if snap_row[i] is not None and str(snap_row[i]) != ''}
                    else:
                        snap = {}
                    result.pruned_snapshots.append({'pk': stale_pk, 'snapshot': snap})

                    conn.execute(
                        f'DELETE FROM [{config.table_name}] WHERE {pk_placeholders}',
                        list(stale_pk),
                    )
                result.pruned = len(stale_pks)
                result.pruned_pks = [pk for pk in stale_pks]
                logger.info(
                    "%s: %d행 삭제(prune) — Excel에서 제거된 행",
                    config.sheet_name, result.pruned,
                )

            # 9. 메타 정보 업데이트 (dry-run 시에도 실행, rollback으로 원복)
            row_count = get_table_row_count(conn, config.table_name)
            update_sync_metadata(conn, config.table_name, now_iso, row_count)

            logger.info(
                "%s: %d행 처리 (신규 %d, 수정 %d, 삭제 %d, 동일 %d, 스킵 %d, 에러 %d)",
                config.sheet_name, result.total_rows,
                result.inserted, result.updated, result.pruned, result.unchanged,
                result.skipped, result.errors,
            )

        except Exception as e:
            result.errors += 1
            result.error_messages.append(str(e))
            logger.error("%s 동기화 실패: %s", config.sheet_name, e)

        return result
