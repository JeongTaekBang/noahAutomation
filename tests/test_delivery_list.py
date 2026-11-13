"""
delivery_list.py 테스트
=======================

공장 출고리스트 Delivery 시트의 헤더 맞춤.

출고리스트는 공장이 만드는 파일이라 헤더가 예고 없이 바뀐다 (2026-11 P11: 영문·혼용 → 한글).
읽는 곳이 둘(`dn_recorder` = DN 출고기록 자동 입력, `reconcile_po` = 매입대사)이라 한 곳에서
예전 이름으로 맞춘다 — 여기서 틀리면 두 기능이 같이 틀린다.
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

import reconcile_po
from po_generator import delivery_list as D
from po_generator import dn_recorder

# 실제 파일의 헤더 순서 그대로
OLD_HEADERS = ['DATE', 'JOB NO', 'BUYER', 'Customer', 'SECTOR', 'AMOUNT', 'RCK ODER',
               '납품 예정일', '납품완료', '계산서금액', 'SO_ID', 'Type', '비고']
NEW_HEADERS = ['구분', '수주번호', '고객명', '발주처', 'SECTOR', '발주 총 금액', '발주번호',
               '납품 일자', '납품된 항목 금액', '분할 납품', 'SO_ID', 'Type', '비고']

NEW_ROW = ['', 'L261521-1R', 'RCK', '동광브이엔티', 'CPI', 8_778_800, 'ND-0955',
           dt.datetime(2026, 11, 2), 5_267_280, None, 'SOD-2026-0953', 'Product',
           '분할납품 2차']


def _empty(headers: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=headers)


def _write(path, headers: list[str], rows: list[list]) -> None:
    pd.DataFrame(rows, columns=headers).to_excel(
        path, sheet_name=D.DELIVERY_SHEET, index=False)


# === 헤더 맞춤 ==============================================================

def test_새_양식_헤더를_예전_이름으로_맞춘다():
    df = D.normalize_columns(_empty(NEW_HEADERS))

    assert list(df.columns) == [
        'DATE', 'JOB NO', 'BUYER', 'Customer', 'SECTOR', 'AMOUNT', 'RCK ODER',
        '납품완료', '계산서금액', '분할 납품', 'SO_ID', 'Type', '비고']


def test_예전_양식은_그대로다():
    df = _empty(OLD_HEADERS)

    assert D.normalize_columns(df) is df


def test_고객명은_거래처가_아니다():
    """새 양식의 `고객명`은 공장 입장의 고객(= RCK)이라 전 행이 'RCK'다 — 거래처는 `발주처`.
    이름만 보고 `고객명`을 Customer로 풀면 확인 목록·대사 결과의 거래처가 전부 'RCK'가 된다"""
    df = D.normalize_columns(pd.DataFrame([NEW_ROW], columns=NEW_HEADERS))

    assert df[D.COL_CUSTOMER].tolist() == ['동광브이엔티']
    assert df['BUYER'].tolist() == ['RCK']
    assert '고객명' not in D.HEADER_ALIASES[D.COL_CUSTOMER]


def test_예전_이름이_이미_있으면_별칭_열을_건드리지_않는다():
    """같은 이름의 열이 둘이면 `df['계산서금액']`이 DataFrame이 되어 엉뚱한 곳에서 터진다"""
    df = D.normalize_columns(_empty(['SO_ID', '계산서금액', '납품된 항목 금액']))

    assert list(df.columns) == ['SO_ID', '계산서금액', '납품된 항목 금액']


@pytest.mark.parametrize('header,canonical', [
    ('납품일자', D.COL_SHIP_DATE),
    ('납품된 항목 금액 ', D.COL_INVOICE_AMOUNT),
    ('발주 번호', D.COL_RCK_ORDER),
    ('RCK  ODER', D.COL_RCK_ORDER),
])
def test_공백만_다른_헤더도_알아본다(header, canonical):
    df = D.normalize_columns(_empty(['SO_ID', header]))

    assert list(df.columns) == ['SO_ID', canonical]


# === 파일 읽기 ==============================================================

def test_필요한_열이_없으면_찾은_이름과_실제_헤더를_알려준다(tmp_path):
    """헤더가 또 바뀌면 `KeyError: '납품완료'` 한 줄이 아니라 무엇이 없는지 보여야 한다"""
    path = tmp_path / '2026리스트_RCK_P12.xlsx'
    _write(path, ['SO_ID', '출고 날짜', '납품된 항목 금액'],
           [['SOD-2026-0001', dt.datetime(2026, 12, 1), 1000]])

    with pytest.raises(ValueError) as exc:
        dn_recorder.load_delivery(path)

    message = str(exc.value)
    assert "'납품완료' / '납품 일자'" in message      # 찾은 이름 (예전 / 새 양식)
    assert '출고 날짜' in message                     # 시트에 실제로 있는 헤더
    assert '계산서금액' not in message                # 있는 열은 말하지 않는다


def test_DN_입력은_새_양식을_예전_이름으로_읽는다(tmp_path):
    path = tmp_path / '2026리스트_RCK_P11.xlsx'
    service = ['RSS', 'L261649', 'RCK', '창진기업', 'CPI', 1_096_630, 'P026066',
               dt.datetime(2026, 11, 2), 1_096_630, None, None, 'Service', None]
    _write(path, NEW_HEADERS, [NEW_ROW, service])

    df = dn_recorder.load_delivery(path)

    assert len(df) == 1                               # SO_ID 없는 서비스 출고는 DN 대상 아님
    row = df.iloc[0]
    assert row[dn_recorder.COL_SO_ID] == 'SOD-2026-0953'
    assert row[dn_recorder.COL_SHIP_DATE] == pd.Timestamp(2026, 11, 2)
    assert row[dn_recorder.COL_INVOICE_AMOUNT] == 5_267_280   # 발주 총 금액(8,778,800)이 아니다
    assert row[dn_recorder.COL_CUSTOMER] == '동광브이엔티'
    assert row[dn_recorder.COL_RCK_ORDER] == 'ND-0955'
    assert row[dn_recorder.COL_JOB_NO] == 'L261521-1R'


def test_매입대사도_새_양식을_예전_이름으로_읽는다(tmp_path):
    path = tmp_path / '2026리스트_RCK_P11.xlsx'
    _write(path, NEW_HEADERS, [NEW_ROW])

    df = reconcile_po.load_delivery(path)

    row = df.iloc[0]
    assert row['RCK ODER'] == 'ND-0955'
    assert row['계산서금액'] == 5_267_280
    assert row['Customer'] == '동광브이엔티'
    assert {'Type', 'SECTOR', 'SO_ID'} <= set(df.columns)


def test_두_양식은_같은_내용이면_같게_읽힌다(tmp_path):
    """양식만 다르고 내용이 같은 두 파일 — 읽은 결과의 공통 열이 값까지 같다"""
    old_row = ['', 'L261521-1R', 'RCK', '동광브이엔티', 'CPI', 8_778_800, 'ND-0955',
               dt.datetime(2026, 10, 30), dt.datetime(2026, 11, 2), 5_267_280,
               'SOD-2026-0953', 'Product', '분할납품 2차']
    old_path, new_path = tmp_path / 'old.xlsx', tmp_path / 'new.xlsx'
    _write(old_path, OLD_HEADERS, [old_row])
    _write(new_path, NEW_HEADERS, [NEW_ROW])

    old, new = dn_recorder.load_delivery(old_path), dn_recorder.load_delivery(new_path)

    shared = [c for c in old.columns if c in new.columns]
    assert set(OLD_HEADERS) - set(shared) == {'납품 예정일'}     # 새 양식에서 빠진 열
    # DATE(구분)는 빈 칸뿐이라 읽으면 dtype만 갈린다 — 값 비교에서 뺀다
    shared.remove('DATE')
    pd.testing.assert_frame_equal(old[shared], new[shared])
