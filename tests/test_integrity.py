"""
데이터 정합성 검사 — 0이어야 하는 것들
=====================================

검사 6종은 `po_generator/integrity.py` 한 곳에 있고, 대시보드(SO/DN 불일치 카드)와
CLI(`check_data.py`)가 같은 함수를 부른다. 여기서 지키고 싶은 것은 둘이다:

1. **정상 업무가 걸리지 않는다** — 부분 출고, 반올림 노이즈, 하이픈 표기 차이, 빈 사업자번호,
   원화 해외 DN. 검사가 늘 빨가면 아무도 안 본다.
2. **검사가 못 돌았으면 0건으로 보이지 않는다** — 대시보드 로더 실패를 배너로 올리는 것과
   같은 이유 ("데이터 없음"과 "조회 실패"는 다른 사건이다).
"""

from __future__ import annotations

import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import check_data
import dashboard
from po_generator import integrity
from po_generator.db_schema import create_fx_table, update_sync_metadata

SO_COLS = [
    'SO_ID', 'Line item', 'Customer name', 'Item name', 'OS name', 'Sector',
    'Item qty', 'Sales Unit Price', 'Currency', 'Status', 'Business registration number',
    'Customer PO', 'PO receipt date',
]
PO_COLS = ['PO_ID', 'SO_ID', 'Customer name', 'Line item', 'Item name', 'Item qty', 'Status']
DN_DOMESTIC_COLS = [
    'DN_ID', 'SO_ID', 'Line item', 'Business registration number', 'Customer name',
    'Item', 'Qty', 'Unit Price', 'Total Sales', '출고일',
]
DN_EXPORT_COLS = [
    'DN_ID', 'SO_ID', 'Line item', 'Customer name', 'Item', 'Qty', 'Currency',
    'Unit Price', 'Total Sales', 'Total Sales KRW', '출고일', '선적일',
]


def so(so_id, line, **kw):
    """SO 라인 — 기본은 주문 10개 × 단가 100, 미출고"""
    row = {'SO_ID': so_id, 'Line item': str(line), 'Customer name': '고객A',
           'Item name': 'IQ3', 'OS name': 'IQ3', 'Sector': 'Water', 'Item qty': '10',
           'Sales Unit Price': '100', 'Currency': 'KRW', 'Status': '미출고',
           'Business registration number': '111-11-11111', 'Customer PO': 'CPO-1',
           'PO receipt date': '2026-08-01 00:00:00'}
    row.update(kw)
    return row


def po(po_id, so_id, line, **kw):
    """PO 라인 — 공장 발주 (SO_ID로 SO와 이어진다)"""
    row = {'PO_ID': po_id, 'SO_ID': so_id, 'Customer name': '고객A', 'Line item': str(line),
           'Item name': 'IQ3', 'Item qty': '10', 'Status': 'Confirmed'}
    row.update(kw)
    return row


def dn(dn_id, so_id, line, qty, total, **kw):
    """국내 DN 라인 — 단가는 금액/수량"""
    row = {'DN_ID': dn_id, 'SO_ID': so_id, 'Line item': str(line),
           'Business registration number': '111-11-11111', 'Customer name': '고객A',
           'Item': 'IQ3', 'Qty': str(qty), 'Unit Price': str(total / qty if qty else 0),
           'Total Sales': str(total), '출고일': '2026-09-01'}
    row.update(kw)
    return row


def dn_exp(dn_id, so_id, line, qty, total, **kw):
    """해외 DN 라인 — 기본 EUR, 9월 선적"""
    row = {'DN_ID': dn_id, 'SO_ID': so_id, 'Line item': str(line), 'Customer name': 'ACME',
           'Item': 'IQ3', 'Qty': str(qty), 'Currency': 'EUR',
           'Unit Price': str(total / qty if qty else 0), 'Total Sales': str(total),
           'Total Sales KRW': str(total * 1500), '출고일': '2026-09-01', '선적일': '2026-09-05'}
    row.update(kw)
    return row


def _mk_db(path=':memory:', so_dom=(), so_exp=(), po_dom=(), po_exp=(), dn_dom=(), dn_ex=(), fx=()):
    """검사가 읽는 테이블만 담은 최소 DB. 값은 실DB처럼 TEXT."""
    conn = sqlite3.connect(str(path))
    tables = (('so_domestic', SO_COLS, so_dom), ('so_export', SO_COLS, so_exp),
              ('po_domestic', PO_COLS, po_dom), ('po_export', PO_COLS, po_exp),
              ('dn_domestic', DN_DOMESTIC_COLS, dn_dom), ('dn_export', DN_EXPORT_COLS, dn_ex))
    for table, cols, rows in tables:
        conn.execute(f"CREATE TABLE {table} ({', '.join(f'[{c}] TEXT' for c in cols)})")
        for row in rows:
            unknown = set(row) - set(cols)
            assert not unknown, f'{table}에 없는 컬럼: {unknown}'
            conn.execute(
                f"INSERT INTO {table} ({', '.join(f'[{c}]' for c in cols)}) "
                f"VALUES ({', '.join('?' * len(cols))})",
                [row.get(c) for c in cols])
    create_fx_table(conn)
    for currency, ym, rate in fx:
        conn.execute('INSERT INTO fx (currency, ym, rate) VALUES (?, ?, ?)', (currency, ym, rate))
    conn.commit()
    return conn


def _check(key):
    return next(c for c in integrity.CHECKS if c.key == key)


def _find(key, conn):
    return _check(key).find(conn)


# ═══════════════════════════════════════════════════════════════
# SO ↔ DN 라인 (과다출고 / 단가 불일치) — 대시보드와 공유하는 산식
# ═══════════════════════════════════════════════════════════════
class TestOverDelivery:

    def test_partial_delivery_is_normal(self):
        """DN 누계 < SO 수량 = 부분 출고 — 정상"""
        conn = _mk_db(so_dom=[so('SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400), dn('DND-2', 'SOD-1', 1, 6, 600)])
        assert _find('over_delivery', conn).empty

    def test_over_delivery_is_flagged(self):
        """분할 출고 누계가 주문수량을 넘으면 걸린다 — 초과분까지"""
        conn = _mk_db(so_dom=[so('SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 6, 600), dn('DND-2', 'SOD-1', 1, 6, 600)])
        rows = _find('over_delivery', conn)
        assert rows[['SO_ID', 'Line', '초과']].values.tolist() == [['SOD-1', 1, 2.0]]

    def test_export_market_is_checked_too(self):
        conn = _mk_db(so_exp=[so('SOO-1', 1, Currency='EUR')],
                      dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 11, 1100)])
        rows = _find('over_delivery', conn)
        assert rows[['시장', 'SO_ID']].values.tolist() == [['해외', 'SOO-1']]


class TestPriceMismatch:

    def test_matching_price_is_normal(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 300)])
        assert _find('price_mismatch', conn).empty

    def test_rounding_noise_is_ignored(self):
        """1원 미만 또는 0.1% 미만 차이는 반올림 — 잡으면 매번 빨갛다"""
        conn = _mk_db(
            so_dom=[so('SOD-1', 1, **{'Sales Unit Price': '333.333'}),
                    so('SOD-2', 1, **{'Item qty': '1000', 'Sales Unit Price': '10000'})],
            dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 1000),            # 0.001원 차이
                    dn('DND-2', 'SOD-2', 1, 1000, 10_000_000 + 5_000)])  # 5,000원이지만 0.05%
        assert _find('price_mismatch', conn).empty

    def test_mismatch_is_flagged_with_diff(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)], dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 360)])
        rows = _find('price_mismatch', conn)
        assert rows[['SO_ID', '기대 금액', 'DN 금액', '차이']].values.tolist() == [
            ['SOD-1', 300.0, 360.0, 60.0]]

    def test_export_compares_in_original_currency(self):
        """외화끼리 비교한다 — 원화 환산을 섞으면 환율 노이즈가 불일치로 보인다"""
        conn = _mk_db(so_exp=[so('SOO-1', 1, Currency='EUR', **{'Sales Unit Price': '10'})],
                      dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 5, 50, **{'Total Sales KRW': '74321'})])
        assert _find('price_mismatch', conn).empty

    def test_cancelled_so_line_is_left_to_its_own_check(self):
        """취소 라인은 SO 단가와 대조하지 않는다 — '취소·보류 SO 라인의 DN'이 따로 잡는다"""
        conn = _mk_db(so_dom=[so('SOD-1', 1, Status='Cancelled')],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 999)])
        assert _find('price_mismatch', conn).empty
        assert len(_find('dn_on_excluded_so', conn)) == 1


class TestDashboardSharesFormula:
    """대시보드 'SO/DN 불일치' 카드와 CLI가 같은 산식을 쓴다 — 사본이 생기면 숫자가 갈린다"""

    def test_dashboard_loader_returns_integrity_result(self, tmp_path, monkeypatch):
        path = tmp_path / 'noah_data.db'
        _mk_db(path, so_dom=[so('SOD-1', 1), so('SOD-2', 1)],
               dn_dom=[dn('DND-1', 'SOD-1', 1, 11, 1100), dn('DND-2', 'SOD-2', 1, 3, 360)]).close()
        monkeypatch.setattr(dashboard, '_conn', lambda: sqlite3.connect(str(path)))
        dashboard.load_so_dn_anomalies.clear()
        try:
            got = dashboard.load_so_dn_anomalies()
        finally:
            dashboard.load_so_dn_anomalies.clear()
        with sqlite3.connect(str(path)) as conn:
            want = integrity.so_dn_line_anomalies(conn)
        pd.testing.assert_frame_equal(got, want)
        assert set(got['SO_ID']) == {'SOD-1', 'SOD-2'}

    def test_dashboard_has_no_copy_of_the_query(self):
        """산식 사본 금지 — `v_dn_revenue`를 한 곳에 모은 것과 같은 이유"""
        source = Path(dashboard.__file__).read_text(encoding='utf-8')
        assert 'WITH so_lines AS' not in source


# ═══════════════════════════════════════════════════════════════
# DN → SO 참조
# ═══════════════════════════════════════════════════════════════
class TestDnWithoutSoLine:

    def test_dn_line_missing_from_so_is_flagged(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 1, 100), dn('DND-1', 'SOD-1', 2, 1, 100)])
        rows = _find('dn_without_so_line', conn)
        assert rows[['DN_ID', 'SO_ID', 'Line']].values.tolist() == [['DND-1', 'SOD-1', 2]]

    def test_line_item_text_format_does_not_matter(self):
        """'1'과 '1.0'은 같은 라인 — 동기화 경로에 따라 둘 다 나온다"""
        conn = _mk_db(so_dom=[so('SOD-1', '1.0')], dn_dom=[dn('DND-1', 'SOD-1', '1', 1, 100)])
        assert _find('dn_without_so_line', conn).empty

    def test_markets_do_not_cross(self):
        """국내 DN은 국내 SO에서만 찾는다 — 해외 SO에 같은 번호가 있어도 짝이 아니다"""
        conn = _mk_db(so_exp=[so('SOD-1', 1)], dn_dom=[dn('DND-1', 'SOD-1', 1, 1, 100)])
        assert len(_find('dn_without_so_line', conn)) == 1


class TestDnOnExcludedSo:

    @pytest.mark.parametrize('status', ['Cancelled', 'Hold'])
    def test_excluded_status_is_flagged(self, status):
        conn = _mk_db(so_dom=[so('SOD-1', 1, Status=status), so('SOD-1', 2)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 1, 100), dn('DND-1', 'SOD-1', 2, 1, 100)])
        rows = _find('dn_on_excluded_so', conn)
        assert rows[['SO_ID', 'Line', 'SO Status']].values.tolist() == [['SOD-1', 1, status]]

    def test_export_is_checked_too(self):
        conn = _mk_db(so_exp=[so('SOO-1', 1, Status='Cancelled')],
                      dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10)])
        assert _find('dn_on_excluded_so', conn)['시장'].tolist() == ['해외']


# ═══════════════════════════════════════════════════════════════
# SO → PO 참조 (대시보드 'PO 미등록'과 같은 정의)
# ═══════════════════════════════════════════════════════════════
class TestSoWithoutPo:

    def test_so_without_any_po_is_flagged(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1), so('SOD-1', 2, **{'Item name': 'ADAPTER'}),
                              so('SOD-2', 1)],
                      po_dom=[po('ND-2', 'SOD-2', 1)])
        rows = _find('so_without_po', conn)
        assert rows[['시장', 'SO_ID', 'Customer PO', '접수일', '라인 수', '품목']].values.tolist() == [
            ['국내', 'SOD-1', 'CPO-1', '2026-08-01', 2, 'IQ3, ADAPTER']]

    def test_line_gap_is_not_a_missing_po(self):
        """PO는 부속·운임을 한 라인에 합쳐 적는다 — SO 2라인 / PO 1라인은 정상"""
        conn = _mk_db(so_exp=[so('SOO-1', 1), so('SOO-1', 2, **{'Item name': 'FEDEX COST'})],
                      po_exp=[po('NO-1', 'SOO-1', 1)])
        assert _find('so_without_po', conn).empty

    def test_cancelled_po_still_counts_as_registered(self):
        """대시보드 'PO 미등록'과 같다 — PO 행이 있으면 등록된 것 (전부 취소는 '발주취소')"""
        conn = _mk_db(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1, Status='Cancelled')])
        assert _find('so_without_po', conn).empty

    def test_cancelled_or_held_so_is_skipped(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1, Status='Cancelled'), so('SOD-2', 1, Status='Hold'),
                              so('SOD-3', 1, Status='Cancelled'), so('SOD-3', 2)])
        rows = _find('so_without_po', conn)
        assert rows[['SO_ID', '라인 수']].values.tolist() == [['SOD-3', 1]], '살아 있는 라인만 센다'

    def test_shipped_so_is_still_checked(self):
        """대시보드 커버리지는 출고 완료를 빼지만 여기선 센다 — 나간 주문에 PO가 없으면 매입 누락"""
        conn = _mk_db(so_dom=[so('SOD-1', 1, Status='출고 완료')],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 10, 1000)])
        assert _find('so_without_po', conn)['SO_ID'].tolist() == ['SOD-1']

    def test_markets_do_not_cross(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)], po_exp=[po('NO-1', 'SOD-1', 1)])
        assert _find('so_without_po', conn)['시장'].tolist() == ['국내']

    def test_many_items_are_shortened(self):
        conn = _mk_db(so_dom=[so('SOD-1', i, **{'Item name': f'P{i}'}) for i in range(1, 6)])
        assert _find('so_without_po', conn)['품목'].tolist() == ['P1, P2, P3 외 2']


class TestPoWithoutSo:

    def test_po_pointing_to_missing_so_is_flagged_per_po(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)],
                      po_dom=[po('ND-1', 'SOD-1', 1),
                              po('ND-2', 'SOD-9', 1), po('ND-2', 'SOD-9', 2, **{'Item name': 'ADAPTER'})])
        rows = _find('po_without_so', conn)
        assert rows[['시장', 'PO_ID', 'SO_ID', '라인 수', '품목', 'Status']].values.tolist() == [
            ['국내', 'ND-2', 'SOD-9', 2, 'IQ3, ADAPTER', 'Confirmed']]

    def test_cancelled_po_is_still_flagged(self):
        """SO 행을 지우면 취소 기록까지 사라진다 — 취소된 PO라도 짝이 없으면 걸린다"""
        conn = _mk_db(po_dom=[po('ND-1', 'SOD-1', 1, Status='Cancelled')])
        assert _find('po_without_so', conn)['Status'].tolist() == ['Cancelled']

    def test_blank_so_id_is_flagged(self):
        conn = _mk_db(po_dom=[po('ND-1', '', 1)])
        assert _find('po_without_so', conn)['SO_ID'].tolist() == ['(공란)']

    def test_markets_do_not_cross(self):
        conn = _mk_db(so_exp=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)])
        assert _find('po_without_so', conn)['시장'].tolist() == ['국내']

    def test_po_only_line_is_normal(self):
        """SO에 없는 PO **라인**(외주 가공비 등)은 정상 — SO_ID만 있으면 된다"""
        conn = _mk_db(so_dom=[so('SOD-1', 1)],
                      po_dom=[po('ND-1', 'SOD-1', 1), po('ND-1', 'SOD-1', 4, **{'Item name': '하부 가공'})])
        assert _find('po_without_so', conn).empty


# ═══════════════════════════════════════════════════════════════
# 한 DN에 거래처 둘 (DN 번호 재사용)
# ═══════════════════════════════════════════════════════════════
class TestDnMixedCustomers:

    def test_two_domestic_customers_are_flagged(self):
        """실측 DND-2026-0328: 코콘 + 한일전자"""
        conn = _mk_db(dn_dom=[
            dn('DND-1', 'SOD-1', 1, 1, 100, **{'Customer name': '코콘',
                                               'Business registration number': '421-87-03388'}),
            dn('DND-1', 'SOD-2', 1, 1, 100, **{'Customer name': '한일전자',
                                               'Business registration number': '698-81-00878'}),
        ])
        rows = _find('dn_mixed_customers', conn)
        assert rows['DN_ID'].tolist() == ['DND-1']
        assert '코콘' in rows['거래처'].iloc[0] and '한일전자' in rows['거래처'].iloc[0]
        assert rows['SO_ID'].iloc[0] == 'SOD-1, SOD-2'

    def test_hyphen_format_is_the_same_customer(self):
        conn = _mk_db(dn_dom=[
            dn('DND-1', 'SOD-1', 1, 1, 100, **{'Business registration number': '421-87-03388'}),
            dn('DND-1', 'SOD-1', 2, 1, 100, **{'Business registration number': '4218703388'}),
        ])
        assert _find('dn_mixed_customers', conn).empty

    def test_blank_biz_no_is_not_another_customer(self):
        """모르는 것을 '다른 거래처'로 세지 않는다 — 묶음 메일과 같은 규칙"""
        conn = _mk_db(dn_dom=[
            dn('DND-1', 'SOD-1', 1, 1, 100),
            dn('DND-1', 'SOD-1', 2, 1, 100, **{'Business registration number': ''}),
        ])
        assert _find('dn_mixed_customers', conn).empty

    def test_export_uses_so_customer_code(self):
        """해외 DN엔 거래처 키가 없다 — SO_해외의 고객코드로 본다 (OC 메일 조인키와 같다)"""
        conn = _mk_db(
            so_exp=[so('SOO-1', 1, **{'Business registration number': 'C-0054'}),
                    so('SOO-2', 1, **{'Business registration number': 'C-0034'}),
                    so('SOO-3', 1, **{'Business registration number': 'C-0054'})],
            dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10), dn_exp('DNO-1', 'SOO-2', 1, 1, 10),
                   dn_exp('DNO-2', 'SOO-1', 1, 1, 10), dn_exp('DNO-2', 'SOO-3', 1, 1, 10)])
        rows = _find('dn_mixed_customers', conn)
        assert rows[['시장', 'DN_ID']].values.tolist() == [['해외', 'DNO-1']]


# ═══════════════════════════════════════════════════════════════
# 환율 없는 외화 DN — v_dn_revenue가 수주시점 환율로 조용히 폴백하는 경우
# ═══════════════════════════════════════════════════════════════
class TestFxMissing:

    def test_missing_rate_is_flagged_per_currency_month(self):
        conn = _mk_db(dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10), dn_exp('DNO-2', 'SOO-2', 1, 2, 30)])
        rows = _find('fx_missing', conn)
        assert rows[['통화', '선적월', '행 수', '외화 금액']].values.tolist() == [
            ['EUR', '2026-09', 2, 40.0]]
        assert rows['DN_ID'].iloc[0] == 'DNO-1, DNO-2'

    def test_rate_present_is_normal(self):
        conn = _mk_db(dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10)], fx=[('EUR', '2026-09', 1600)])
        assert _find('fx_missing', conn).empty

    def test_krw_and_unshipped_are_normal(self):
        """원화는 환산할 것이 없고, 선적 전이면 아직 매출이 아니다"""
        conn = _mk_db(dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10, Currency='KRW'),
                             dn_exp('DNO-2', 'SOO-2', 1, 1, 10, 선적일='')])
        assert _find('fx_missing', conn).empty

    def test_blank_currency_is_flagged(self):
        """통화 공란도 뷰에선 같은 폴백을 탄다 — 조용히 넘어가면 안 된다"""
        conn = _mk_db(dn_ex=[dn_exp('DNO-1', 'SOO-1', 1, 1, 10, Currency='')],
                      fx=[('EUR', '2026-09', 1600)])
        assert _find('fx_missing', conn)['통화'].tolist() == ['(공란)']


# ═══════════════════════════════════════════════════════════════
# 실행
# ═══════════════════════════════════════════════════════════════
class TestRunChecks:

    def test_clean_db_is_all_ok(self):
        conn = _mk_db(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 300)])
        results = integrity.run_checks(conn)
        assert [r.check.key for r in results] == [c.key for c in integrity.CHECKS]
        assert all(r.ok for r in results)

    def test_failed_check_is_not_zero(self):
        """조회가 죽은 검사는 '0건'이 아니라 실패다 — 나머지 검사는 계속 돈다"""
        conn = _mk_db(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)],
                      dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 300)])
        conn.execute('DROP TABLE fx')
        results = {r.check.key: r for r in integrity.run_checks(conn)}
        assert results['fx_missing'].error and not results['fx_missing'].ok
        assert results['fx_missing'].rows is None
        assert all(r.ok for k, r in results.items() if k != 'fx_missing')

    def test_check_keys_and_titles_are_unique(self):
        keys = [c.key for c in integrity.CHECKS]
        titles = [c.title for c in integrity.CHECKS]
        assert len(set(keys)) == len(keys) and len(set(titles)) == len(titles)

    def test_guide_lists_every_check(self):
        """운영 가이드의 '검사 목록' 표 = CHECKS — 검사를 더하고 문서를 빠뜨리면 여기서 걸린다

        걸리면 무엇을 고치는지가 문서에만 있으면, 새 검사가 [문제]로 떴을 때 볼 데가 없다.
        """
        guide = Path(integrity.__file__).resolve().parent.parent / 'docs' / '정합성검사_가이드.md'
        text = guide.read_text(encoding='utf-8')
        section = text.split('\n## 검사 목록\n', 1)[1].split('\n## ', 1)[0]
        documented = set(re.findall(r'^\| \*\*(.+?)\*\* \|', section, flags=re.M))
        assert documented == {c.title for c in integrity.CHECKS}


# ═══════════════════════════════════════════════════════════════
# CLI (check_data.py)
# ═══════════════════════════════════════════════════════════════
class TestCli:

    @pytest.fixture
    def env(self, tmp_path, monkeypatch):
        """DB·Excel·출력 폴더를 전부 임시로. 마지막 동기화는 2026-10-13 10:00."""
        db = tmp_path / 'noah_data.db'
        excel = tmp_path / 'NOAH_SO_PO_DN.xlsx'
        excel.write_bytes(b'')
        out = tmp_path / 'generated_check'
        monkeypatch.setattr(check_data, 'DB_FILE', db)
        monkeypatch.setattr(check_data, 'NOAH_SO_PO_DN_FILE', excel)
        monkeypatch.setattr(check_data, 'CHECK_OUTPUT_DIR', out)
        monkeypatch.setattr('sys.argv', ['check_data.py'])
        synced = datetime(2026, 10, 13, 10, 0)
        os.utime(excel, (synced.timestamp() - 3600,) * 2)  # 동기화 1시간 전에 저장

        def build(**tables):
            conn = _mk_db(db, **tables)
            update_sync_metadata(conn, 'so_domestic', synced.isoformat(), 1)
            conn.commit()
            conn.close()
        return build, excel, out, synced

    CLEAN = dict(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)],
                 dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 300)])
    MISMATCH = dict(so_dom=[so('SOD-1', 1)], po_dom=[po('ND-1', 'SOD-1', 1)],
                    dn_dom=[dn('DND-1', 'SOD-1', 1, 3, 360)])

    @staticmethod
    def _result(out):
        return out / check_data.RESULT_FILE_NAME

    @staticmethod
    def _info(path):
        """요약 시트 표 아래 정보 블록 {이름: 값}"""
        from openpyxl import load_workbook
        ws = load_workbook(path)['요약']
        return {r[0]: r[1] for r in ws.iter_rows(values_only=True)
                if r[0] in ('검사 시각', '데이터 기준', '주의')}

    def test_clean_run_still_writes_result(self, env, capsys):
        """정상이어도 결과 파일을 쓴다 — 안 쓰면 직전의 '문제' 파일이 최신처럼 남는다"""
        build, _, out, _ = env
        build(**self.CLEAN)
        assert check_data.main() == 0
        text = capsys.readouterr().out
        assert '데이터 기준: 2026-10-13 10:00' in text
        assert '[주의]' not in text
        sheets = pd.read_excel(self._result(out), sheet_name=None)
        assert list(sheets) == ['요약']
        assert set(sheets['요약']['상태'].head(len(integrity.CHECKS))) == {'정상'}  # 표 아래는 정보 블록

    def test_findings_exit_one_and_write_detail(self, env, capsys):
        build, _, out, _ = env
        build(**self.MISMATCH)
        assert check_data.main() == 1
        assert 'SOD-1' in capsys.readouterr().out
        sheets = pd.read_excel(self._result(out), sheet_name=None)
        assert list(sheets)[0] == '요약'
        assert len(sheets) == 2  # 요약 + 걸린 검사 1장
        detail = next(df for name, df in sheets.items() if name != '요약')
        assert detail['SO_ID'].tolist() == ['SOD-1']

    def test_rerun_after_fix_replaces_stale_result(self, env):
        """사용자 보고(2026-10-13): 고치고 다시 돌렸는데 결과 파일이 그대로 '문제'였다

        문제가 있을 때만 파일을 만들던 시절 — 정상 재실행이 파일을 안 써서 15:11의 결과가
        최신처럼 남았다. 이제 같은 파일을 매번 덮어쓴다.
        """
        build, _, out, _ = env
        build(**self.MISMATCH)
        check_data.main()
        assert len(pd.read_excel(self._result(out), sheet_name=None)) == 2

        check_data.DB_FILE.unlink()
        build(**self.CLEAN)  # 시트를 고치고 DB Sync 한 상태
        assert check_data.main() == 0
        sheets = pd.read_excel(self._result(out), sheet_name=None)
        assert list(sheets) == ['요약'], '고친 뒤 결과에 지난 문제 시트가 남으면 안 된다'
        assert list(out.iterdir()) == [self._result(out)], '파일은 하나 — 어느 게 최신인지 고르게 하지 않는다'

    def test_result_says_when_and_on_which_data(self, env):
        """덮어쓰는 파일이라 언제 결과인지 파일 안에 적는다 (인쇄 머리글은 화면에 안 보인다)"""
        build, excel, out, synced = env
        build(**self.CLEAN)
        check_data.main()
        info = self._info(self._result(out))
        assert info['데이터 기준'].startswith('2026-10-13 10:00')
        assert info['검사 시각'][:4] == str(datetime.now().year)
        assert '주의' not in info

        os.utime(excel, (synced.timestamp() + 3600,) * 2)
        check_data.main()
        assert '2026-10-13 11:00' in self._info(self._result(out))['주의']

    def test_open_result_file_warns_and_keeps_exit_code(self, env, capsys, monkeypatch):
        """Excel에 열려 있으면 덮어쓰지 못한다 — 조용히 넘기면 열린 창이 낡은 결과를 보여 준다"""
        build, _, out, _ = env
        build(**self.CLEAN)

        def locked(*_a, **_k):
            raise PermissionError(13, 'Permission denied')
        monkeypatch.setattr(check_data, 'write_report', locked)
        assert check_data.main() == 0
        assert '열려 있어' in capsys.readouterr().out

    def test_excel_newer_than_sync_warns(self, env, capsys):
        """DB가 Excel보다 낡았으면 결과가 '이미 고친 것'을 보여줄 수 있다 — 먼저 알린다"""
        build, excel, _, synced = env
        build()
        os.utime(excel, (synced.timestamp() + 3600,) * 2)
        check_data.main()
        assert '[주의] Excel이 마지막 동기화 이후 수정됐습니다' in capsys.readouterr().out

    def test_failed_check_is_reported_and_exits_one(self, env, capsys):
        build, _, out, _ = env
        build()
        conn = sqlite3.connect(str(check_data.DB_FILE))
        conn.execute('DROP TABLE fx')
        conn.commit()
        conn.close()
        assert check_data.main() == 1
        assert '[실패]' in capsys.readouterr().out
        sheets = pd.read_excel(self._result(out), sheet_name=None)
        assert '실패' in sheets['요약']['상태'].tolist()

    def test_missing_db_exits_one_without_creating_it(self, env, capsys):
        """sqlite3.connect는 없는 파일을 만들어 버린다 — 빈 DB로 '모두 정상'을 내면 안 된다"""
        assert check_data.main() == 1
        assert not check_data.DB_FILE.exists()
        assert 'DB 파일이 없습니다' in capsys.readouterr().out
