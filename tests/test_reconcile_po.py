"""reconcile_po — Excel_vs_AX 상세 시트 회귀 테스트

핵심 불변식: Excel_vs_AX의 AX 합계는 항상 GRN 총액과 같아야 한다.
요약 시트는 Service를 잔차(GRN 총액 − 국내 − 해외)로 계산하므로 AX 합계가 정의상
GRN 총액이다. 두 시트의 Diff 총액이 어긋나면 사용자는 어느 쪽을 믿을지 알 수 없다.
"""

import pandas as pd
import pytest

import reconcile_po as R


def _grn(rows):
    return pd.DataFrame(rows, columns=['Purchase order', 'Cost amount physical',
                                       'Item name'])


def _delivery(rows):
    return pd.DataFrame(rows, columns=['RCK ODER', 'Type', 'AX PO',
                                       '계산서금액', 'Customer'])


def _po(rows):
    return pd.DataFrame(rows, columns=['구분', 'AX PO', 'Total ICO',
                                       'PO_ID', 'Customer name'])


@pytest.fixture
def collided():
    """같은 AX PO가 PO_국내 라인(오타)과 직접출고 Service 건에 동시에 존재.

    2026-08 실측: ND-0668 L5~8의 AX PO가 P024791 대신 P024792로 입력되어,
    비와이 직접출고 P024792와 충돌했다.
    """
    df_po = _po([
        ('국내', 'P024792', 1_524_920, 'ND-0668', '주식회사 스칸텍'),
        ('국내', 'P024788', 1_524_920, 'ND-0668', '주식회사 스칸텍'),
    ])
    delivery = _delivery([
        ('P024792', 'Service', 'P024792', 2_530_020, '비와이'),
        ('ND-0668', 'Product', 'P024788', 1_524_920, '스칸텍'),
    ])
    grn = _grn([
        ('P024792', 2_530_020, 'Local hand station (lhs) b7631'),
        ('P024788', 1_524_920, 'Ytc limit switch'),
        ('P024791', 1_524_920, 'Ytc limit switch'),   # 오타로 고아가 된 진짜 배치
    ])
    return df_po, delivery, grn


def test_ax_total_equals_grn_total(collided):
    """AX 합계 == GRN 총액 — 요약 시트와 tie-out되는 조건"""
    detail = R._build_excel_vs_ax(*collided)
    assert detail['AX'].sum() == pytest.approx(collided[2]['Cost amount physical'].sum())


def test_grn_not_duplicated_across_classes(collided):
    """AX PO가 두 분류에 걸려도 GRN은 국내 우선 1행에만 귀속 (요약과 동일)"""
    detail = R._build_excel_vs_ax(*collided)
    rows = detail[detail['AX PO'] == 'P024792']
    assert len(rows) == 2, "분류별 행은 그대로 두 줄로 보인다"
    dom = rows[rows['분류'] == 'Product(국내)'].iloc[0]
    svc = rows[rows['분류'] == 'Service'].iloc[0]
    assert dom['AX'] == 2_530_020        # GRN은 국내 행에만
    assert svc['AX'] == 0
    # Excel은 양쪽 다 원래 금액을 유지 (요약의 excel_dom / excel_service와 동일)
    assert dom['Excel'] == 1_524_920
    assert svc['Excel'] == 2_530_020


def test_class_subtotals_match_summary_buckets(collided):
    """분류별 소계가 요약 시트의 버킷 계산과 일치"""
    df_po, delivery, grn = collided
    detail = R._build_excel_vs_ax(df_po, delivery, grn)
    by = detail.groupby('분류')[['Excel', 'AX']].sum()

    # 요약 시트와 같은 방식으로 직접 계산
    dom_keys = set(df_po.loc[df_po['구분'] == '국내', 'AX PO'])
    gp, ga = grn['Purchase order'], grn['Cost amount physical']
    assert by.loc['Product(국내)', 'Excel'] == df_po['Total ICO'].sum()
    assert by.loc['Product(국내)', 'AX'] == ga[gp.isin(dom_keys)].sum()
    assert by.loc['Service', 'Excel'] == 2_530_020
    assert by.loc['Service', 'AX'] == ga.sum() - ga[gp.isin(dom_keys)].sum()


def test_grn_only_po_becomes_service_row(collided):
    """Excel 어디에도 없는 GRN(오타로 고아가 된 건)은 Service 행으로 노출"""
    detail = R._build_excel_vs_ax(*collided)
    orphan = detail[detail['AX PO'] == 'P024791']
    assert len(orphan) == 1
    row = orphan.iloc[0]
    assert row['분류'] == 'Service'
    assert row['Excel'] == 0
    assert row['AX'] == 1_524_920
    assert row['참조'] == 'Ytc limit switch'   # GRN Item name으로 채움


def test_collision_logs_warning(collided, caplog):
    """AX PO 충돌은 조용히 넘어가지 않고 경고로 남는다"""
    with caplog.at_level('WARNING', logger='reconcile_po'):
        R._build_excel_vs_ax(*collided)
    assert 'P024792' in caplog.text


def test_no_collision_is_unchanged():
    """충돌이 없는 평범한 경우엔 기존 동작 그대로"""
    df_po = _po([('국내', 'P1', 100, 'ND-1', 'A'),
                 ('해외', 'P2', 200, 'NO-1', 'B')])
    delivery = _delivery([('P3', 'YTC', 'P3', 300, 'C'),
                          ('P4', 'Service', 'P4', 400, 'D')])
    grn = _grn([('P1', 100, 'x'), ('P2', 190, 'y'),
                ('P3', 300, 'z'), ('P4', 380, 'w')])
    detail = R._build_excel_vs_ax(df_po, delivery, grn)
    assert len(detail) == 4
    assert detail['Excel'].sum() == 1000
    assert detail['AX'].sum() == 970
    assert detail['Diff'].sum() == -30
    assert list(detail['분류']) == ['Product(국내)', 'Product(국내)',
                                   'Product(해외)', 'Service']
