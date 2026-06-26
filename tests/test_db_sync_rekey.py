"""
미완성 행 보류 + 재키잉(키변경) 인식 테스트
============================================

Excel→SQLite 동기화에서, 위치성 키 컬럼(Line item 등)을 비운 채 먼저 동기화하거나
나중에 편집하면 같은 행이 '삭제+신규'로 기록돼 "데이터는 있는데 삭제로 뜨는" 오해가
발생했다. 이를 근본적으로 막는 두 장치를 검증한다:

1. 미완성 행 보류 — _row_seq 외 PK(자연키) 컬럼이 비면 채워질 때까지 동기화 제외
2. 재키잉 인식 — 같은 sync 안 '내용 동일' 삭제+신규 쌍을 '키변경(수정)' 1건으로 합침
"""

import sqlite3

import pandas as pd
import pytest

from po_generator.db_sync import SyncEngine
from sync_db import _reconcile_rekeys, write_sync_log_to_db
from po_generator.db_sync import SheetSyncResult, SyncSummary

PO_PK = ('PO_ID', 'Line item', '_row_seq')


def _write_po_sheet(path, line_item):
    """PO_국내 시트 1행짜리 임시 Excel 생성."""
    df = pd.DataFrame([{
        'PO_ID': 'TEST-1', 'SO_ID': 'S1', 'Line item': line_item,
        'Item name': 'Widget', 'Customer name': 'ACME', 'Total ICO': '0',
    }])
    with pd.ExcelWriter(path) as w:
        df.to_excel(w, sheet_name='PO_국내', index=False)


# ── 1. 미완성 행 보류 (실엔진) ────────────────────────────────────────────

def test_empty_keycol_row_is_held(tmp_path):
    """Line item(키 구성요소)이 비면 동기화 보류 — 신규/삭제 모두 안 생김."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_po_sheet(xlsx, '')  # 빈 Line item
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['PO_국내']).results[0]
    assert r.inserted == 0
    assert r.pruned == 0
    assert r.skipped == 1


def test_filled_keycol_yields_clean_insert(tmp_path):
    """빈값으로 보류된 뒤 Line item을 채우면 깨끗한 신규 1건 (삭제 없음)."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_po_sheet(xlsx, '')
    SyncEngine(xlsx, db).sync_all(sheet_filter=['PO_국내'])  # 보류
    _write_po_sheet(xlsx, '1')                               # 채움
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['PO_국내']).results[0]
    assert r.inserted == 1
    assert r.pruned == 0
    # DB에 단 1행
    n = sqlite3.connect(db).execute('SELECT COUNT(*) FROM po_domestic').fetchone()[0]
    assert n == 1


# ── 2. 재키잉 인식 (순수 함수) ────────────────────────────────────────────

def test_rekey_value_edit_collapses_to_update():
    """Line item 1→12 편집: 삭제+신규 → 키변경(수정) 1건."""
    ins = [{'pk': ('TEST-1', '12', '1'),
            'values': {'Item name': 'Widget', 'SO_ID': 'S1', '_row_seq': '1'}}]
    prn = [{'pk': ('TEST-1', '1', '1'),
            'snapshot': {'Item name': 'Widget', 'SO_ID': 'S1', 'Line item': '1', '_row_seq': '1'}}]
    rekeys, rem_ins, rem_prn = _reconcile_rekeys(PO_PK, ins, prn)
    assert len(rekeys) == 1
    assert not rem_ins and not rem_prn
    assert rekeys[0]['pk'] == ('TEST-1', '12', '1')
    assert rekeys[0]['changes']['Line item'] == {'old': '1', 'new': '12'}


def test_rekey_fill_sparse_snapshot_matches_superset():
    """삭제 스냅샷이 신규 값의 부분집합이어도(키 채우며 값 추가) 매칭."""
    ins = [{'pk': ('TEST-1', '1', '1'),
            'values': {'Item name': 'Widget', 'SO_ID': 'S1', 'Line item': '1', 'Status': 'Open'}}]
    prn = [{'pk': ('TEST-1', '', '1'),
            'snapshot': {'Item name': 'Widget', 'SO_ID': 'S1'}}]  # Line item/Status 없음
    rekeys, rem_ins, rem_prn = _reconcile_rekeys(PO_PK, ins, prn)
    assert len(rekeys) == 1 and not rem_ins and not rem_prn
    assert rekeys[0]['changes']['Line item'] == {'old': None, 'new': '1'}


def test_guard_different_entity_not_paired():
    """문서ID(PO_ID)가 다르면 절대 묶이지 않음."""
    ins = [{'pk': ('TEST-1', '12', '1'), 'values': {'Item name': 'Widget'}}]
    prn = [{'pk': ('TEST-2', '1', '1'), 'snapshot': {'Item name': 'Widget'}}]
    rekeys, rem_ins, rem_prn = _reconcile_rekeys(PO_PK, ins, prn)
    assert not rekeys and len(rem_ins) == 1 and len(rem_prn) == 1


def test_guard_different_content_not_paired():
    """비키 컬럼 내용이 다르면 묶이지 않음."""
    ins = [{'pk': ('TEST-1', '12', '1'), 'values': {'Item name': 'Gadget'}}]
    prn = [{'pk': ('TEST-1', '1', '1'), 'snapshot': {'Item name': 'Widget'}}]
    rekeys, rem_ins, rem_prn = _reconcile_rekeys(PO_PK, ins, prn)
    assert not rekeys and len(rem_ins) == 1 and len(rem_prn) == 1


def test_guard_ambiguous_match_kept_as_delete():
    """삭제 1건에 신규 후보가 2건이면 모호 → 묶지 않고 삭제로 보존(안전)."""
    ins = [
        {'pk': ('TEST-1', '12', '1'), 'values': {'Item name': 'Widget'}},
        {'pk': ('TEST-1', '13', '1'), 'values': {'Item name': 'Widget'}},
    ]
    prn = [{'pk': ('TEST-1', '1', '1'), 'snapshot': {'Item name': 'Widget'}}]
    rekeys, rem_ins, rem_prn = _reconcile_rekeys(PO_PK, ins, prn)
    assert not rekeys
    assert len(rem_ins) == 2 and len(rem_prn) == 1


def test_empty_inputs_passthrough():
    """입력이 비면 그대로 통과(에러 없음)."""
    assert _reconcile_rekeys(PO_PK, [], []) == ([], [], [])
    ins = [{'pk': ('A', '1', '1'), 'values': {'x': '1'}}]
    assert _reconcile_rekeys(PO_PK, ins, []) == ([], ins, [])


# ── 3. 로그 기록 통합 (write_sync_log_to_db) ──────────────────────────────

def test_write_log_records_rekey_as_single_update(tmp_path):
    """재키잉이 _sync_log에 '수정' 1행으로만 남고, 삭제/신규는 안 남는다."""
    db = tmp_path / 'log.db'
    res = SheetSyncResult(sheet_name='PO_국내', table_name='po_domestic')
    res.inserted_details = [{'pk': ('TEST-1', '12', '1'),
                             'values': {'Item name': 'Widget', 'SO_ID': 'S1', '_row_seq': '1'}}]
    res.pruned_snapshots = [{'pk': ('TEST-1', '1', '1'),
                             'snapshot': {'Item name': 'Widget', 'SO_ID': 'S1',
                                          'Line item': '1', '_row_seq': '1'}}]
    summary = SyncSummary(results=[res], started_at='2026-06-26 13:12:00')

    write_sync_log_to_db(summary, db_path=db)

    rows = sqlite3.connect(db).execute(
        'SELECT change_type, pk_display, changes_json FROM _sync_log ORDER BY id'
    ).fetchall()
    types = [r[0] for r in rows]
    assert types == ['수정'], f"기대=['수정'] 실제={types}"
    assert 'TEST-1 | 12 | 1' == rows[0][1]
    assert 'Line item' in rows[0][2]
