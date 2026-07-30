"""
미완성 행 보류 + 재키잉(키변경) 인식 테스트
============================================

Excel→SQLite 동기화에서, 위치성 키 컬럼(Line item 등)을 비운 채 먼저 동기화하거나
나중에 편집하면 같은 행이 '삭제+신규'로 기록돼 "데이터는 있는데 삭제로 뜨는" 오해가
발생했다. 이를 근본적으로 막는 두 장치를 검증한다:

1. 미완성 행 보류 — _row_seq 외 PK(자연키) 컬럼이 비면 채워질 때까지 동기화 제외
2. 재키잉 인식 — 같은 sync 안 '내용 동일' 삭제+신규 쌍을 '키변경(수정)' 1건으로 합침
3. 중복 자연키 보존 — PO/DN은 자연키가 유일하지 않다(분할출고). _row_seq가 없으면
   뒤 행이 앞 행을 덮어써 매출·수량이 조용히 사라진다
4. PK 마이그레이션 로그 — PK 정의를 바꿔 전 행을 재적재할 때 수천 건의 가짜 '신규'로
   변경 이력을 덮지 않는다
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


# ── 4. 중복 자연키 보존 (분할출고) ────────────────────────────────────────

def _write_dn_sheet(path, rows):
    """DN_국내 시트 생성. rows = [(line_item, qty, total_sales), ...]"""
    df = pd.DataFrame([{
        'DN_ID': 'DND-1', 'SO_ID': 'S1', 'Line item': li,
        'Item': 'NOS160-MS-FC', 'Qty': qty, 'Total Sales': amt,
        'Customer name': 'ACME', '출고일': '2026-06-05',
    } for li, qty, amt in rows])
    with pd.ExcelWriter(path) as w:
        df.to_excel(w, sheet_name='DN_국내', index=False)


def test_dn_duplicate_natural_key_rows_both_survive(tmp_path):
    """같은 (DN_ID, SO_ID, Line item)을 두 행으로 나눈 분할출고가 둘 다 남는다.

    실측 사례: DND-2026-0511/SOD-2026-0232/2 가 Qty 11 + 12 두 행이었는데
    _row_seq 없던 시절 뒤 행이 앞 행을 덮어써 18,656,000원이 사라졌다.
    """
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_dn_sheet(xlsx, [('2', '11', '18656000'), ('2', '12', '20352000')])
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['DN_국내']).results[0]
    assert r.inserted == 2, f'두 행 모두 들어와야 한다 (실제 {r.inserted})'

    rows = sqlite3.connect(db).execute(
        'SELECT _row_seq, Qty, [Total Sales] FROM dn_domestic ORDER BY _row_seq'
    ).fetchall()
    assert [x[0] for x in rows] == ['1', '2']
    assert [x[1] for x in rows] == ['11', '12']
    assert sum(int(x[2]) for x in rows) == 39008000   # 금액 누락 없음


def test_dn_export_duplicate_natural_key_rows_both_survive(tmp_path):
    """DN_해외도 같은 처방 (선적 분할)."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    pd.DataFrame([{
        'DN_ID': 'DNO-1', 'SO_ID': 'S1', 'Line item': '1', 'Item': 'IQ3',
        'Qty': q, 'Currency': 'USD', 'Total Sales': '100',
        'Total Sales KRW': '150000', '선적일': '2026-07-21',
    } for q in ('3', '7')]).to_excel(xlsx, sheet_name='DN_해외', index=False)
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['DN_해외']).results[0]
    assert r.inserted == 2
    assert sqlite3.connect(db).execute('SELECT COUNT(*) FROM dn_export').fetchone()[0] == 2


def test_dn_second_sync_is_unchanged_not_rekeyed(tmp_path):
    """분할출고 행이 있어도 재동기화는 '동일' — _row_seq가 안정적으로 재부여된다."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_dn_sheet(xlsx, [('2', '11', '18656000'), ('2', '12', '20352000')])
    SyncEngine(xlsx, db).sync_all(sheet_filter=['DN_국내'])
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['DN_국내']).results[0]
    assert (r.inserted, r.updated, r.pruned) == (0, 0, 0)
    assert r.unchanged == 2


# ── 5. PK 마이그레이션 로그 ───────────────────────────────────────────────

def test_pk_migration_logs_single_reload_not_mass_insert(tmp_path):
    """PK 정의 변경 재적재는 '재적재' 1행만 남긴다 (행별 가짜 '신규' 금지).

    상세가 남아 있어도(엔진이 수집을 건너뛰지만, 방어) 접는지 확인한다.
    """
    db = tmp_path / 'log.db'
    res = SheetSyncResult(sheet_name='DN_국내', table_name='dn_domestic', total_rows=1370)
    res.pk_migrated = True
    res.inserted = 1370
    res.inserted_details = [{'pk': ('D1', 'S1', '1', str(i)), 'values': {'Qty': '11'}}
                            for i in range(1, 4)]
    write_sync_log_to_db(SyncSummary(results=[res], started_at='2026-07-30 21:00:00'),
                         db_path=db)
    rows = sqlite3.connect(db).execute(
        'SELECT change_type, pk_display, changes_json FROM _sync_log').fetchall()
    assert len(rows) == 1, f'1행이어야 한다 (실제 {len(rows)})'
    assert rows[0][0] == '재적재'
    assert '1370' in rows[0][2]


def test_normal_sync_still_logs_per_row(tmp_path):
    """마이그레이션이 아니면 종전대로 행별 기록 — 억제가 새다가 감사 이력을 먹지 않는다."""
    db = tmp_path / 'log.db'
    res = SheetSyncResult(sheet_name='DN_국내', table_name='dn_domestic', total_rows=2)
    res.inserted_details = [{'pk': ('D1', 'S1', '1', '1'), 'values': {'Qty': '11'}},
                            {'pk': ('D1', 'S1', '1', '2'), 'values': {'Qty': '12'}}]
    write_sync_log_to_db(SyncSummary(results=[res], started_at='2026-07-30 21:00:00'),
                         db_path=db)
    rows = sqlite3.connect(db).execute('SELECT change_type FROM _sync_log').fetchall()
    assert [r[0] for r in rows] == ['신규', '신규']
