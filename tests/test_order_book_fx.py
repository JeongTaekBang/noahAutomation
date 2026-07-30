"""
Order Book 환율 재환산 + 매출 귀속월 테스트
===========================================

Order Book Output을 `AX_매출대사`와 같은 기준으로 맞춘 두 축을 검증한다:

1. **해외 환율** — Output KRW = 외화금액 × 선적월 환율. 시트 `Total Sales KRW`(수주시점
   환율)와의 차액은 `fx_variance`로 빠져 **Ending은 재환산 전과 동일**해야 한다.
2. **국내 귀속월** — Output 귀속월 = 세금계산서 발행월 (N/A=발행불필요면 출고월,
   선수금+출고면 출고월, 아무것도 없으면 미인식 → Backlog 잔류).

산식은 `v_dn_revenue` 뷰 한 곳에만 있으므로(db_schema.py) 뷰를 직접 검증한다.
"""

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from po_generator.db_schema import create_fx_table, create_order_book_views
from po_generator.db_sync import SyncEngine

SQL_DIR = Path(__file__).resolve().parent.parent / 'sql'

DN_DOMESTIC_COLS = [
    'DN_ID', 'SO_ID', 'Line item', 'Qty', 'Total Sales',
    '출고일', '세금계산서 발행일', '선수금 세금계산서 발행일',
    'Customer name', 'Item', 'Customer PO', 'Business registration number',
]
DN_EXPORT_COLS = [
    'DN_ID', 'SO_ID', 'Line item', 'Qty', 'Currency',
    'Total Sales', 'Total Sales KRW', '출고일', '선적일',
    'Customer name', 'Item', 'Customer PO',
]
SO_DOMESTIC_COLS = [
    'SO_ID', 'Line item', 'Customer name', 'Customer PO', 'Item name', 'OS name',
    'Item qty', 'Sales amount', 'Period', 'AX Period', 'Model code', 'Sector',
    'Business registration number', 'Industry code', 'Expected delivery date', 'Status',
]
SO_EXPORT_COLS = [c if c != 'Sales amount' else 'Sales amount KRW' for c in SO_DOMESTIC_COLS]


def _mk_db(tmp_path, dn_dom=(), dn_exp=(), so_dom=(), so_exp=(), fx=()):
    """v_dn_revenue를 붙인 최소 DB 생성. 행은 dict로 주고 빈 컬럼은 NULL."""
    db = tmp_path / 'ob.db'
    conn = sqlite3.connect(str(db))
    conn.row_factory = sqlite3.Row
    for table, cols in (('dn_domestic', DN_DOMESTIC_COLS), ('dn_export', DN_EXPORT_COLS),
                        ('so_domestic', SO_DOMESTIC_COLS), ('so_export', SO_EXPORT_COLS)):
        conn.execute(f"CREATE TABLE {table} ({', '.join(f'[{c}] TEXT' for c in cols)})")
    for table, cols, rows in (('dn_domestic', DN_DOMESTIC_COLS, dn_dom),
                              ('dn_export', DN_EXPORT_COLS, dn_exp),
                              ('so_domestic', SO_DOMESTIC_COLS, so_dom),
                              ('so_export', SO_EXPORT_COLS, so_exp)):
        for row in rows:
            unknown = set(row) - set(cols)
            assert not unknown, f'{table}에 없는 컬럼: {unknown}'
            conn.execute(
                f"INSERT INTO {table} ({', '.join(f'[{c}]' for c in cols)}) "
                f"VALUES ({', '.join('?' * len(cols))})",
                [row.get(c) for c in cols])
    create_order_book_views(conn)
    for currency, ym, rate in fx:
        conn.execute('INSERT INTO fx (currency, ym, rate) VALUES (?, ?, ?)', (currency, ym, rate))
    conn.commit()
    return conn


def _rev(conn):
    """v_dn_revenue를 (SO_ID, Line item) → row 로."""
    return {(r['SO_ID'], r['Line item']): r for r in conn.execute('SELECT * FROM v_dn_revenue')}


# ═══════════════════════════════════════════════════════════════
# 1. 국내 매출 귀속월
# ═══════════════════════════════════════════════════════════════

def test_domestic_uses_tax_invoice_month_not_shipping_month(tmp_path):
    """세금계산서가 출고 다음 달이면 Output은 세금계산서 월로 귀속된다."""
    conn = _mk_db(tmp_path, dn_dom=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '2', 'Total Sales': '1000',
        '출고일': '2026-06-28 00:00:00', '세금계산서 발행일': '2026-07-03 00:00:00',
    }])
    assert _rev(conn)[('S1', 1)]['매출월'] == '2026-07'


def test_domestic_na_invoice_marker_recognizes_at_shipping_month(tmp_path):
    """'N/A' = 발행 불필요(무상공급·FOC·반품) → 출고월에 인식.

    이걸 미인식으로 두면 0원 무상공급 수량이 Backlog에 영구히 남는다.
    """
    conn = _mk_db(tmp_path, dn_dom=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '100', 'Total Sales': '0',
        '출고일': '2026-07-10 00:00:00', '세금계산서 발행일': 'N/A',
    }])
    row = _rev(conn)[('S1', 1)]
    assert row['매출월'] == '2026-07'
    assert row['Qty'] == 100


def test_domestic_advance_invoice_falls_back_to_shipping_month(tmp_path):
    """선수금 세금계산서 + 출고 완료 → 출고월에 수익인식."""
    conn = _mk_db(tmp_path, dn_dom=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '1', 'Total Sales': '500',
        '출고일': '2026-05-20 00:00:00', '선수금 세금계산서 발행일': '2026-03-02 00:00:00',
    }])
    assert _rev(conn)[('S1', 1)]['매출월'] == '2026-05'


def test_domestic_without_any_invoice_is_unrecognized(tmp_path):
    """세금계산서·선수금 둘 다 없으면 매출월 NULL — Output 없음 = Backlog 잔류."""
    conn = _mk_db(tmp_path, dn_dom=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '1', 'Total Sales': '407000',
        '출고일': '2026-07-03 00:00:00',
    }])
    assert _rev(conn)[('S1', 1)]['매출월'] is None


def test_domestic_never_has_fx_variance(tmp_path):
    """국내는 KRW 거래 — 재환산 없음, fx_variance 항상 0."""
    conn = _mk_db(tmp_path, dn_dom=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '3', 'Total Sales': '1234567',
        '출고일': '2026-07-01 00:00:00', '세금계산서 발행일': '2026-07-01 00:00:00',
    }])
    row = _rev(conn)[('S1', 1)]
    assert row['fx_variance'] == 0
    assert row['output_amount'] == 1234567


# ═══════════════════════════════════════════════════════════════
# 2. 해외 선적월 환율 재환산
# ═══════════════════════════════════════════════════════════════

def test_export_restates_at_shipping_month_rate(tmp_path):
    """실제 사례 SOO-2026-0188: USD 1,276 · 6월 수주 · 7월 선적.

    시트 KRW는 6월 환율(1,914,046)인데 매출은 7월 환율로 인식돼야 한다.
    """
    conn = _mk_db(tmp_path, dn_exp=[{
        'DN_ID': 'DNO-2026-0140', 'SO_ID': 'SOO-2026-0188', 'Line item': '1', 'Qty': '2',
        'Currency': 'USD', 'Total Sales': '1276', 'Total Sales KRW': '1914046',
        '선적일': '2026-07-21 00:00:00',
    }], fx=[('USD', '2026-06', 1500.036), ('USD', '2026-07', 1548.608203890816)])
    row = _rev(conn)[('SOO-2026-0188', 1)]
    assert row['매출월'] == '2026-07'
    assert row['output_amount'] == 1976024        # 1276 × 1548.6082 (라인 단위 반올림)
    assert row['fx_variance'] == 61978            # 1,976,024 − 1,914,046


def test_export_missing_rate_falls_back_to_sheet_krw(tmp_path):
    """선적월 환율이 아직 없으면 시트 KRW 유지 + Variance 0.

    Output이 조용히 0이 되면 매출이 사라진 것처럼 보이므로 폴백이 필수다.
    """
    conn = _mk_db(tmp_path, dn_exp=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '1',
        'Currency': 'EUR', 'Total Sales': '100', 'Total Sales KRW': '176881',
        '선적일': '2026-08-05 00:00:00',
    }], fx=[('EUR', '2026-07', 1768.809)])   # 8월 환율 없음
    row = _rev(conn)[('S1', 1)]
    assert row['output_amount'] == 176881
    assert row['fx_variance'] == 0


def test_export_krw_currency_uses_foreign_amount_column(tmp_path):
    """해외지만 통화가 KRW면 재환산 없이 Total Sales를 쓴다 (AX_매출대사와 동일)."""
    conn = _mk_db(tmp_path, dn_exp=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '1',
        'Currency': 'KRW', 'Total Sales': '500000', 'Total Sales KRW': '500000',
        '선적일': '2026-07-01 00:00:00',
    }], fx=[('USD', '2026-07', 1548.6)])
    row = _rev(conn)[('S1', 1)]
    assert row['output_amount'] == 500000
    assert row['fx_variance'] == 0


def test_export_unshipped_is_unrecognized(tmp_path):
    """미선적(선적일 없음) → 매출월 NULL → Output 없음 = Backlog 잔류."""
    conn = _mk_db(tmp_path, dn_exp=[{
        'DN_ID': 'D1', 'SO_ID': 'S1', 'Line item': '1', 'Qty': '1',
        'Currency': 'USD', 'Total Sales': '100', 'Total Sales KRW': '150000',
    }], fx=[('USD', '2026-07', 1548.6)])
    assert _rev(conn)[('S1', 1)]['매출월'] is None


# ═══════════════════════════════════════════════════════════════
# 3. 원장 항등식 — Ending이 재환산 도입 전과 같아야 한다
# ═══════════════════════════════════════════════════════════════

def test_order_book_ending_unchanged_by_restatement(tmp_path):
    """SO를 6월 환율로 수주하고 7월에 선적 → Output은 7월 환율, Variance가 차액 흡수.

    Ending = Start + Input − Output + Variance = 0 (환율차가 잔고를 오염시키지 않는다)
    """
    conn = _mk_db(
        tmp_path,
        so_exp=[{
            'SO_ID': 'SOO-1', 'Line item': '1', 'OS name': 'Noah NA', 'Item qty': '2',
            'Sales amount KRW': '1914046', 'Period': '2026-06',
            'Expected delivery date': '2026-06-29 00:00:00', 'Customer name': 'ACME',
        }],
        dn_exp=[{
            'DN_ID': 'D1', 'SO_ID': 'SOO-1', 'Line item': '1', 'Qty': '2', 'Currency': 'USD',
            'Total Sales': '1276', 'Total Sales KRW': '1914046',
            '선적일': '2026-07-21 00:00:00',
        }],
        fx=[('USD', '2026-07', 1548.608203890816)],
    )
    df = pd.read_sql_query((SQL_DIR / 'order_book.sql').read_text(encoding='utf-8'), conn)
    jul = df[df['Period'] == '2026-07'].iloc[0]
    assert jul['Value_Output_amount'] == 1976024
    assert jul['Value_Variance_amount'] == 61978
    assert jul['Value_Ending_amount'] == 0            # ← 환율차 흡수 확인
    assert jul['Value_Ending_qty'] == 0

    jun = df[df['Period'] == '2026-06'].iloc[0]
    assert jun['Value_Input_amount'] == 1914046
    assert jun['Value_Ending_amount'] == 1914046      # 6월엔 수주시점 환율로 잔고


def test_order_book_identity_holds_on_every_row(tmp_path):
    """모든 행에서 Ending = Start + Input − Output + Variance."""
    conn = _mk_db(
        tmp_path,
        so_dom=[{'SO_ID': 'SOD-1', 'Line item': '1', 'OS name': 'MA', 'Item qty': '5',
                 'Sales amount': '5000000', 'Period': '2026-05',
                 'Expected delivery date': '2026-06-30 00:00:00'}],
        dn_dom=[{'DN_ID': 'D1', 'SO_ID': 'SOD-1', 'Line item': '1', 'Qty': '3',
                 'Total Sales': '3000000', '출고일': '2026-06-10 00:00:00',
                 '세금계산서 발행일': '2026-06-10 00:00:00'}],
        so_exp=[{'SO_ID': 'SOO-1', 'Line item': '1', 'OS name': 'CVA', 'Item qty': '2',
                 'Sales amount KRW': '1914046', 'Period': '2026-06',
                 'Expected delivery date': '2026-07-01 00:00:00'}],
        dn_exp=[{'DN_ID': 'D2', 'SO_ID': 'SOO-1', 'Line item': '1', 'Qty': '2',
                 'Currency': 'USD', 'Total Sales': '1276', 'Total Sales KRW': '1914046',
                 '선적일': '2026-07-21 00:00:00'}],
        fx=[('USD', '2026-07', 1548.608203890816)],
    )
    df = pd.read_sql_query((SQL_DIR / 'order_book.sql').read_text(encoding='utf-8'), conn)
    assert len(df) > 0
    calc = (df['Value_Start_amount'] + df['Value_Input_amount']
            - df['Value_Output_amount'] + df['Value_Variance_amount'])
    assert (calc - df['Value_Ending_amount']).abs().max() < 0.5


def test_backlog_keeps_unrecognized_domestic_shipment(tmp_path):
    """세금계산서 미발행 출고분은 Output이 아니라 Backlog에 남는다."""
    conn = _mk_db(
        tmp_path,
        so_dom=[{'SO_ID': 'SOD-1', 'Line item': '1', 'OS name': 'MA', 'Item qty': '1',
                 'Sales amount': '407000', 'Period': '2026-07',
                 'Expected delivery date': '2026-07-14 00:00:00'}],
        dn_dom=[{'DN_ID': 'D1', 'SO_ID': 'SOD-1', 'Line item': '1', 'Qty': '1',
                 'Total Sales': '407000', '출고일': '2026-07-03 00:00:00'}],
    )
    df = pd.read_sql_query((SQL_DIR / 'order_book_backlog.sql').read_text(encoding='utf-8'), conn)
    assert len(df) == 1
    assert df.iloc[0]['잔여수량'] == 1


# ═══════════════════════════════════════════════════════════════
# 4. FX 시트 동기화
# ═══════════════════════════════════════════════════════════════

@pytest.mark.parametrize('header,expected', [
    ('2026-01', '2026-01'),
    (' 2026-12 ', '2026-12'),
    ('2026-13', None),
    ('FX', None),
    (pd.Timestamp('2026-03-01'), '2026-03'),
])
def test_fx_month_key_normalizes_headers(header, expected):
    """월 컬럼 헤더는 텍스트/날짜 둘 다 YYYY-MM으로 정규화, 아니면 None."""
    assert SyncEngine._fx_month_key(header) == expected


def _write_fx_sheet(path, rows):
    df = pd.DataFrame(rows, columns=['FX', '2026-06', '2026-07'])
    with pd.ExcelWriter(path) as w:
        df.to_excel(w, sheet_name='FX', index=False)


def test_fx_sync_unpivots_and_tracks_changes(tmp_path):
    """가로형 FX 시트 → fx(currency, ym, rate) 언피벗. 환율 변경은 '수정'으로 잡힌다."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_fx_sheet(xlsx, [['USD', 1500.036, 1548.608], ['EUR', 1743.959, None]])
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['FX']).results[0]
    assert (r.inserted, r.updated, r.pruned, r.errors) == (3, 0, 0, 0)  # 빈칸은 제외

    conn = sqlite3.connect(str(db))
    assert conn.execute("SELECT rate FROM fx WHERE currency='USD' AND ym='2026-07'"
                        ).fetchone()[0] == pytest.approx(1548.608)
    conn.close()

    # 환율 수정 + 통화 삭제
    _write_fx_sheet(xlsx, [['USD', 1500.036, 1549.0]])
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['FX']).results[0]
    assert (r.inserted, r.updated, r.pruned) == (0, 1, 1)
    assert r.updated_details[0]['changes']['rate'][1] == pytest.approx(1549.0)


def test_fx_sync_errors_when_month_columns_missing(tmp_path):
    """월 컬럼을 못 찾으면 조용히 빈 fx로 두지 않고 에러로 세운다.

    fx가 비면 뷰가 시트 KRW로 폴백해 '숫자가 틀린 채' 돌아간다.
    """
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    pd.DataFrame([['USD', 1500.0]], columns=['FX', 'Jan']).to_excel(
        xlsx, sheet_name='FX', index=False)
    r = SyncEngine(xlsx, db).sync_all(sheet_filter=['FX']).results[0]
    assert r.errors == 1
    assert 'YYYY-MM' in r.error_messages[0]


def test_sync_creates_fx_table_and_view(tmp_path):
    """동기화가 끝나면 fx 테이블과 v_dn_revenue 뷰가 준비돼 있다."""
    xlsx, db = tmp_path / 't.xlsx', tmp_path / 't.db'
    _write_fx_sheet(xlsx, [['USD', 1500.0, 1548.6]])
    SyncEngine(xlsx, db).sync_all(sheet_filter=['FX'])
    conn = sqlite3.connect(str(db))
    names = {r[0] for r in conn.execute("SELECT name FROM sqlite_master")}
    conn.close()
    assert 'fx' in names
    assert 'v_dn_revenue' in names


def test_create_order_book_views_is_idempotent(tmp_path):
    """뷰 정의가 바뀌어도 매 동기화에서 갱신되도록 DROP+CREATE — 반복 호출 안전."""
    conn = sqlite3.connect(str(tmp_path / 'v.db'))
    create_fx_table(conn)
    for _ in range(3):
        create_order_book_views(conn)
    assert conn.execute(
        "SELECT COUNT(*) FROM sqlite_master WHERE name='v_dn_revenue'").fetchone()[0] == 1
    conn.close()
