"""
dn_recorder.py 테스트
=====================

출고리스트 → `DN_국내` 추가 행 계산.

회귀 방지의 핵심은 **자동 입력에 잘못된 수량이 섞이지 않는 것**이다.
이 시트는 매출의 단일 소스(`v_dn_revenue`)라 틀린 수량이 들어가면 Order Book·대시보드·
매출대사가 전부 같이 틀린다. 그래서 확신이 없는 건은 넣지 말고 `reviews`로 빼야 한다.
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from po_generator import dn_recorder as R


# === 픽스처 ================================================================

SO_A = 'SOD-2026-0100'
BIZ_A = '615-81-88675'
SHIP = dt.datetime(2026, 8, 6)


def _frame(defaults: dict, rows: list[dict]) -> pd.DataFrame:
    """행이 없어도 **열은 있는** DataFrame — 시트를 실제로 읽으면 그렇게 나온다"""
    return pd.DataFrame([{**defaults, **row} for row in rows],
                        columns=list(defaults))


def make_delivery(rows: list[dict]) -> pd.DataFrame:
    """출고리스트 Delivery 시트 형태 (load_delivery를 통과한 뒤 모습)"""
    df = _frame({
        R.COL_SO_ID: SO_A,
        R.COL_SHIP_DATE: SHIP,
        R.COL_INVOICE_AMOUNT: 800_000,
        R.COL_CUSTOMER: '엔이에스',
        R.COL_RCK_ORDER: 'ND-0100',
        'AMOUNT': 800_000,
    }, rows)
    df[R.COL_SHIP_DATE] = pd.to_datetime(df[R.COL_SHIP_DATE])
    return df


def make_so(rows: list[dict]) -> pd.DataFrame:
    return _frame({
        'SO_ID': SO_A,
        'Line item': 1,
        'Item name': 'NOS100-M',
        'Item qty': 10,
        'Business registration number': BIZ_A,
        'Customer name': '엔이에스 주식회사',
        'Currency': 'KRW',
        # 파워쿼리 캐시라 출고 판정에는 못 쓰지만 취소/보류는 이 컬럼으로만 알 수 있다
        'Status': '미출고',
    }, rows)


def make_po(rows: list[dict]) -> pd.DataFrame:
    return _frame({
        'PO_ID': 'ND-0100',
        'SO_ID': SO_A,
        'Line item': 1,
        'Item name': 'NOS100-M',
        'Item qty': 10,
        'ICO Unit': 80_000,
        'Total ICO': 800_000,
        'Status': 'Invoiced P08',
    }, rows)


def make_dn(rows: list[dict]) -> pd.DataFrame:
    return _frame({
        'DN_ID': 'DND-2026-0001',
        'SO_ID': SO_A,
        'Line item': 1,
        'Qty': 1,
        'Currency': 'KRW',
        '출고일': dt.datetime(2026, 1, 5),
        '세금계산서 발행일': dt.datetime(2026, 1, 5),
        'Remarks': None,
        'Business registration number': BIZ_A,
        'Seq': 1,
    }, rows)


def qty_map(plan: R.Plan) -> dict[int, int]:
    return {line.line_item: line.qty for line in plan.lines}


def run(delivery, so, po, dn, period='P08', **kw) -> R.Plan:
    return R.build_plan(period, delivery, so, po, dn, **kw)


# === 라인/수량 결정 =========================================================

def test_마지막_출고는_SO_전_라인의_잔량을_쓴다():
    """PO에 남은 게 없으면 SO 기준 — PO가 부속을 1라인에 합쳐 적기 때문"""
    delivery = make_delivery([{}])
    so = make_so([
        {'Line item': 1, 'Item qty': 10},
        {'Line item': 2, 'Item qty': 2, 'Item name': '부싱가공'},  # PO엔 없는 부속
    ])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'SO_ID': 'SOD-2026-0001', 'Qty': 1}])

    plan = run(delivery, so, po, dn)

    assert qty_map(plan) == {1: 10, 2: 2}
    assert not plan.reviews


def test_부분_출고는_PO_Invoiced_라인만_쓴다():
    """PO에 아직 안 나간 라인이 있으면 이번 달 Invoiced 라인/수량 그대로"""
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 240_000}])
    so = make_so([
        {'Line item': 1, 'Item qty': 10},
        {'Line item': 2, 'Item qty': 5},
    ])
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 240_000},
        {'Line item': 2, 'Item qty': 5, 'Status': 'Confirmed'},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert qty_map(plan) == {1: 3}
    assert not plan.reviews


def test_기출고_수량을_잔량에서_뺀다():
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 4, '출고일': dt.datetime(2026, 7, 1)}])

    plan = run(delivery, so, po, dn)

    assert qty_map(plan) == {1: 6}


def test_나중_출고는_잔량에서_빼지_않는다():
    """지난 기간을 다시 돌릴 때 아직 일어나지도 않은 출고가 잔량을 갉아먹으면 안 된다"""
    delivery = make_delivery([{R.COL_SHIP_DATE: dt.datetime(2026, 8, 6)}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 4, '출고일': dt.datetime(2026, 9, 1),
                   'SO_ID': 'SOD-2026-0999'}])

    plan = run(delivery, so, po, dn)

    assert qty_map(plan) == {1: 10}


def test_SO가_취소된_라인은_PO가_Invoiced여도_제외한다():
    """판매가 취소돼도 공장은 이미 만든 것을 계산서로 넘기기도 한다 — 그건 매입만 발생한 것
    (2026-07 SOD-2026-0364 L5: SO는 Cancelled인데 PO L5 'IP66 TEST 시료 값'은 Invoiced)."""
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 1_800_000}])
    so = make_so([
        {'Line item': 1, 'Item qty': 6},
        {'Line item': 5, 'Item qty': 1, 'Item name': 'IP66 TEST 인증비',
         'Status': 'Cancelled'},
    ])
    po = make_po([
        {'Line item': 1, 'Item qty': 6, 'Total ICO': 800_000},
        {'Line item': 5, 'Item qty': 1, 'Total ICO': 1_000_000},   # 시료값은 매입 발생
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews          # 금액 대조에는 1,000,000이 그대로 들어가야 맞는다
    assert qty_map(plan) == {1: 6}


def test_SO가_Hold인_라인도_제외한다():
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 800_000}])
    so = make_so([
        {'Line item': 1, 'Item qty': 6},
        {'Line item': 2, 'Item qty': 1, 'Status': 'Hold'},
    ])
    po = make_po([{'Line item': 1, 'Item qty': 6, 'Total ICO': 800_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert qty_map(plan) == {1: 6}


def test_전_라인이_취소면_확인필요():
    """'SO 없음'(오타)과 '전부 취소'는 사람에게 다른 얘기다"""
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10, 'Status': 'Cancelled'}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])

    plan = run(delivery, so, po, make_dn([]))

    assert [r.reason for r in plan.reviews] == ['기록할 라인 없음']


def test_잔량이_0이면_확인필요가_아니라_건너뜀이다():
    """사람이 출고일을 하루 이틀 다르게 적어 둔 경우 (2026-07 SOD-2026-0790:
    출고리스트 7/29, DN 7/27). 확인 목록에 올리면 헛경보만 쌓인다."""
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 10, '출고일': dt.datetime(2026, 8, 4)}])

    plan = run(delivery, so, po, dn)

    assert not plan.reviews
    assert [r.reason for r in plan.skipped] == ['이미 출고 완료']


def test_취소된_PO_라인은_제외한다():
    """Cancelled를 안 빼면 취소분이 출고로 잡힌다 (2026-05/06 실측 2건)"""
    delivery = make_delivery([{}])
    so = make_so([
        {'Line item': 1, 'Item qty': 10},
        {'Line item': 2, 'Item qty': 3},
    ])
    po = make_po([
        {'Line item': 1, 'Item qty': 10},
        {'Line item': 2, 'Item qty': 3, 'Status': 'Cancelled', 'Total ICO': 111},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert qty_map(plan) == {1: 10}


def test_SO에_없는_PO_라인은_DN에_넣지_않는다():
    """DN은 매출 장부다 — PO에만 있는 라인은 고객에게 판 게 아니라 매입만 발생한 것
    (2026-05 SOD-2026-0188 L4 'De-cluch Gear Box Bushing 하부 가공' 450,000원)."""
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 1_250_000}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])          # SO는 1라인뿐
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 800_000},
        {'Line item': 4, 'Item qty': 1, 'Total ICO': 450_000,  # 외주 가공비
         'Item name': '하부 가공'},
        {'Line item': 1, 'Item qty': 7, 'Status': 'Confirmed'},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews          # 금액 대조에는 450,000이 그대로 들어가야 맞는다
    assert qty_map(plan) == {1: 3}   # DN에는 안 들어간다


def test_여러_날_출고에서도_SO에_없는_라인은_뺀다():
    """금액 배정은 PO 전 라인으로 하고, DN에 쓸 때만 뺀다"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 750_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 500_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 3}, {'Line item': 2, 'Item qty': 5}])
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 300_000},
        {'Line item': 4, 'Item qty': 1, 'Total ICO': 450_000},   # SO에 없음
        {'Line item': 2, 'Item qty': 5, 'Total ICO': 500_000},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews
    assert {(l.ship_date, l.line_item): l.qty for l in plan.lines} == {
        (pd.Timestamp(2026, 8, 4), 1): 3,
        (pd.Timestamp(2026, 8, 6), 2): 5,
    }


def test_분할발주는_같은_라인을_합산한다():
    """PO는 같은 라인을 여러 행으로 나눠 적는다 (_row_seq가 PK에 있는 이유와 같은 사정)"""
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 800_000}])
    so = make_so([{'Line item': 1, 'Item qty': 10},
                  {'Line item': 2, 'Item qty': 1}])
    po = make_po([
        {'Line item': 1, 'Item qty': 4, 'Total ICO': 320_000},
        {'Line item': 1, 'Item qty': 6, 'Total ICO': 480_000},
        {'Line item': 2, 'Item qty': 1, 'Status': 'Confirmed'},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert qty_map(plan) == {1: 10}


# === 자기검증 ==============================================================

def test_금액이_어긋나면_확인필요로_뺀다():
    delivery = make_delivery([{R.COL_INVOICE_AMOUNT: 750_000}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Total ICO': 800_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.lines
    assert [r.reason for r in plan.reviews] == ['금액 불일치']


def test_여러_날_출고는_PO_금액으로_날짜에_배정한다():
    """PO는 분할출고마다 행을 따로 두므로 ICO 합이 날짜별 계산서금액과 맞는다"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 300_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 500_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 3}, {'Line item': 2, 'Item qty': 5}])
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 300_000},
        {'Line item': 2, 'Item qty': 5, 'Total ICO': 500_000},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews
    by_date = {(l.ship_date, l.line_item): l.qty for l in plan.lines}
    assert by_date == {
        (pd.Timestamp(2026, 8, 4), 1): 3,
        (pd.Timestamp(2026, 8, 6), 2): 5,
    }
    # 날짜마다 DN_ID가 따로 붙는다
    assert len({l.dn_id for l in plan.lines}) == 2


def test_같은_라인이_여러_날_나눠_나가도_배정한다():
    """분할발주된 같은 라인 — 수량이 다르면 금액으로 갈린다 (2026-06 SOD-2026-0232 형태)"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 240_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 560_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 240_000},
        {'Line item': 1, 'Item qty': 7, 'Total ICO': 560_000},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews
    assert {(l.ship_date, l.qty) for l in plan.lines} == {
        (pd.Timestamp(2026, 8, 4), 3), (pd.Timestamp(2026, 8, 6), 7)}


def test_단가_정정_행은_출고로_치지_않는다():
    """출고 한 번 + 나중에 금액만 정정한 행 (2026-05 SOD-2026-0306:
    5/11 9,956,592 → 5/13 `L260441-1R`로 103,616 = 차액). 합계가 PO ICO와 맞으면
    첫 날 한 건으로 기록하고 뒤 행은 건너뛴다."""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 780_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 20_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 8}])
    po = make_po([{'Line item': 1, 'Item qty': 8, 'Total ICO': 800_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews
    assert len(plan.lines) == 1
    line = plan.lines[0]
    assert (line.line_item, line.qty) == (1, 8)
    assert line.ship_date == pd.Timestamp(2026, 8, 4)      # 실제 출고일
    assert [r.reason for r in plan.skipped] == ['단가 정정 행']


def test_반품이_섞이면_확인필요로_뺀다():
    """출고 → 반품 → 재출고를 어떻게 적을지는 사람이 정한다 (2026-03 SOD-2026-0280)"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 3), R.COL_INVOICE_AMOUNT: 800_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: -800_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 850_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 9}])
    po = make_po([{'Line item': 1, 'Item qty': 9, 'Total ICO': 850_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.lines
    assert {r.reason for r in plan.reviews} == {'반품 포함'}


def test_합계도_안_맞으면_확인필요로_뺀다():
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 400_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 300_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Total ICO': 800_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.lines
    assert {r.reason for r in plan.reviews} == {'같은 SO 여러 날 출고'}


def test_배정이_갈리면_확인필요로_뺀다():
    """같은 금액의 서로 다른 라인 — 어느 쪽이 어느 날인지 금액으로는 못 가른다"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 400_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 400_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 1}, {'Line item': 2, 'Item qty': 2}])
    po = make_po([
        {'Line item': 1, 'Item qty': 1, 'Total ICO': 400_000},
        {'Line item': 2, 'Item qty': 2, 'Total ICO': 400_000},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.lines
    assert {r.reason for r in plan.reviews} == {'같은 SO 여러 날 출고'}


def test_같은_날_두_줄이면_한_출고로_합친다():
    """출고리스트에 같은 주문·같은 날이 두 줄로 적히기도 한다 (2026-05 SOD-2026-0467)"""
    delivery = make_delivery([
        {R.COL_INVOICE_AMOUNT: 300_000},
        {R.COL_INVOICE_AMOUNT: 500_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10, 'Total ICO': 800_000}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.reviews
    assert len({l.dn_id for l in plan.lines}) == 1
    assert qty_map(plan) == {1: 10}


def test_앞_출고_때문에_뒤_출고가_사라지지_않는다():
    """중복 방어는 실행 전 시트 상태만 봐야 한다 (2026-05 SOD-2026-0188에서 터졌던 버그)"""
    delivery = make_delivery([
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 4), R.COL_INVOICE_AMOUNT: 240_000},
        {R.COL_SHIP_DATE: dt.datetime(2026, 8, 6), R.COL_INVOICE_AMOUNT: 80_000},
    ])
    so = make_so([{'Line item': 1, 'Item qty': 4}])
    po = make_po([
        {'Line item': 1, 'Item qty': 3, 'Total ICO': 240_000},
        {'Line item': 1, 'Item qty': 1, 'Total ICO': 80_000},
    ])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.skipped
    assert len(plan.lines) == 2


# === 금액 배정 단위 테스트 ==================================================

def test_split_by_amount_유일해를_찾는다():
    items = [(1, 3, 300.0), (2, 5, 500.0), (3, 1, 200.0)]
    assert R.split_by_amount(items, [500.0, 500.0]) is None   # 300+200 ↔ 500이 자리를 바꿀 수 있다
    assert R.split_by_amount(items, [800.0, 200.0]) == [{1: 3, 2: 5}, {3: 1}]


def test_split_by_amount_같은_결과를_주는_해는_모호가_아니다():
    """항목이 같은 값·같은 라인이면 어느 쪽을 골라도 결과가 같다 (2026-03 SOD-2026-0220)"""
    items = [(1, 1, 500.0), (1, 1, 500.0), (2, 1, 100.0)]
    assert R.split_by_amount(items, [600.0, 500.0]) == [{1: 1, 2: 1}, {1: 1}]


def test_split_by_amount_못_나누면_None():
    assert R.split_by_amount([(1, 10, 800.0)], [400.0, 400.0]) is None


def test_split_by_amount_결과가_갈리면_None():
    items = [(1, 1, 400.0), (2, 2, 400.0)]
    assert R.split_by_amount(items, [400.0, 400.0]) is None


def test_split_by_amount_항목이_너무_많으면_포기한다():
    items = [(i, 1, float(i)) for i in range(R.MAX_SPLIT_ITEMS + 1)]
    assert R.split_by_amount(items, [1.0, 2.0]) is None


def test_SO가_없으면_확인필요():
    delivery = make_delivery([{R.COL_SO_ID: 'SOD-2026-9999'}])
    plan = run(delivery, make_so([]), make_po([]), make_dn([]))

    assert not plan.lines
    assert [r.reason for r in plan.reviews] == ['SO 없음']


def test_이번_기간_Invoiced가_없으면_확인필요():
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Status': 'Confirmed'}])

    plan = run(delivery, so, po, make_dn([]))

    assert not plan.lines
    assert [r.reason for r in plan.reviews] == ['Invoiced P08 없음']


def test_출고일이_비면_확인필요():
    delivery = make_delivery([{R.COL_SHIP_DATE: None}])
    plan = run(delivery, make_so([{}]), make_po([{}]), make_dn([]))

    assert not plan.lines
    assert [r.reason for r in plan.reviews] == ['출고일 없음']


# === 멱등성 ================================================================

def test_이미_기록된_출고는_건너뛴다():
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 10, '출고일': SHIP}])

    plan = run(delivery, so, po, dn)

    assert not plan.lines
    assert not plan.reviews
    assert [r.reason for r in plan.skipped] == ['이미 입력됨']


def test_다른_날짜로_기록해_둔_건도_건너뛴다():
    """출고리스트 날짜와 다르게 적어 둔 건 — 키만 보면 못 걸러 통째로 중복 입력된다
    (2026-03 SOD-2026-0156: 출고리스트 3/23, DN 3/9)"""
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 10, '출고일': dt.datetime(2026, 8, 20)}])

    plan = run(delivery, so, po, dn)

    assert not plan.lines
    assert [r.reason for r in plan.skipped] == ['이미 입력됨(다른 출고일)']


def test_수량이_다른_뒤_회차는_건너뛰지_않는다():
    """누계로 보면 뒤 회차가 앞 회차에 가려 사라진다 (2026-07 SOD-2026-0713 실측:
    7/23 L1 1개를 8/6의 L1 2개가 덮었다). 라인·수량 조합이 똑같을 때만 건너뛴다."""
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 3}])
    po = make_po([
        {'Line item': 1, 'Item qty': 1, 'Total ICO': 800_000},
        {'Line item': 1, 'Item qty': 2, 'Status': 'Invoiced P09'},   # 다음 달 몫
    ])
    dn = make_dn([{'Line item': 1, 'Qty': 2, '출고일': dt.datetime(2026, 9, 6)}])

    plan = run(delivery, so, po, dn)

    assert not plan.skipped
    assert qty_map(plan) == {1: 1}


def test_일부만_기록돼_있으면_나머지를_넣는다():
    delivery = make_delivery([{}])
    so = make_so([{'Line item': 1, 'Item qty': 10}])
    po = make_po([{'Line item': 1, 'Item qty': 10}])
    dn = make_dn([{'Line item': 1, 'Qty': 3, '출고일': dt.datetime(2026, 7, 1)}])

    plan = run(delivery, so, po, dn)

    assert qty_map(plan) == {1: 7}


# === 채번 ==================================================================

def test_DN_ID는_출고리스트_1행당_1개이고_연번이다():
    delivery = make_delivery([
        {R.COL_SO_ID: 'SOD-2026-0101', R.COL_SHIP_DATE: dt.datetime(2026, 8, 4)},
        {R.COL_SO_ID: 'SOD-2026-0102', R.COL_SHIP_DATE: dt.datetime(2026, 8, 6)},
    ])
    so = make_so([
        {'SO_ID': 'SOD-2026-0101', 'Line item': 1, 'Item qty': 10},
        {'SO_ID': 'SOD-2026-0101', 'Line item': 2, 'Item qty': 1},
        {'SO_ID': 'SOD-2026-0102', 'Line item': 1, 'Item qty': 10},
    ])
    po = make_po([
        {'SO_ID': 'SOD-2026-0101'},
        {'SO_ID': 'SOD-2026-0102'},
    ])
    dn = make_dn([{'DN_ID': 'DND-2026-0723', 'SO_ID': 'SOD-2026-0001', 'Seq': 500}])

    plan = run(delivery, so, po, dn)

    assert plan.dn_ids == ['DND-2026-0724', 'DND-2026-0725']
    assert [l.seq for l in plan.lines] == [501, 502, 503]


def test_출고일_순서대로_채번한다():
    """출고리스트가 날짜순이 아니어도 DN_ID는 날짜순이어야 한다"""
    delivery = make_delivery([
        {R.COL_SO_ID: 'SOD-2026-0102', R.COL_SHIP_DATE: dt.datetime(2026, 8, 6)},
        {R.COL_SO_ID: 'SOD-2026-0101', R.COL_SHIP_DATE: dt.datetime(2026, 8, 4)},
    ])
    so = make_so([
        {'SO_ID': 'SOD-2026-0101', 'Item qty': 10},
        {'SO_ID': 'SOD-2026-0102', 'Item qty': 10},
    ])
    po = make_po([{'SO_ID': 'SOD-2026-0101'}, {'SO_ID': 'SOD-2026-0102'}])

    plan = run(delivery, so, po, make_dn([]))

    assert [(l.so_id, l.dn_id) for l in plan.lines] == [
        ('SOD-2026-0101', 'DND-2026-0001'),
        ('SOD-2026-0102', 'DND-2026-0002'),
    ]


def test_다른_연도_DN_ID는_채번에_끼어들지_않는다():
    dn = make_dn([{'DN_ID': 'DND-2025-0900'}, {'DN_ID': 'DND-2026-0007'}])
    assert R.next_dn_seq(dn, 2026) == 8
    assert R.next_dn_seq(dn, 2025) == 901
    assert R.next_dn_seq(dn, 2027) == 1


# === 세금계산서 발행일 / Remarks ==============================================

def test_세금계산서_발행일은_채우지_않는다():
    """사람이 직접 입력하는 칸이다 (2026-08-07 사용자 지시).
    발행 시점은 출고와 별개고, 이 날짜의 '월'이 Order Book 매출 인식월을 정한다 —
    공란이면 미인식(Backlog 잔류)이라 잘못된 월에 매출이 잡히지 않는다."""
    plan = run(make_delivery([{}]), make_so([{}]), make_po([{}]), make_dn([]))

    assert plan.lines[0].tax_date is None
    assert plan.lines[0].remarks is None


def test_월합_거래처는_Remarks를_물려받는다():
    """나중에 발행일을 채울 때 어느 날짜를 쓸지 알 수 있어야 한다"""
    dn = make_dn([{'Business registration number': BIZ_A,
                   'Remarks': '25일 마감, 월합세금계산서'}])

    plan = run(make_delivery([{}]), make_so([{}]), make_po([{}]), dn)

    assert plan.lines[0].tax_date is None
    assert plan.lines[0].remarks == '25일 마감, 월합세금계산서'


def test_하이픈_표기가_달라도_같은_거래처로_본다():
    dn = make_dn([{'Business registration number': '6158188675',
                   'Remarks': '월합 세금계산서'}])

    plan = run(make_delivery([{}]), make_so([{}]), make_po([{}]), dn)

    assert plan.lines[0].remarks == '월합 세금계산서'


# === 파일 핸들 =============================================================

def test_시트를_읽고_나면_파일을_붙잡고_있지_않는다(tmp_path):
    """`pd.ExcelFile`을 닫지 않으면 뒤이어 Excel이 같은 파일을 쓰기로 못 연다 —
    우리 프로세스가 우리를 막는다 (2026-08-07 실사용에서 두 번 터졌다).
    Windows는 열린 핸들이 있으면 rename을 막으므로 그걸로 확인한다."""
    book = tmp_path / "wb.xlsx"
    with pd.ExcelWriter(book) as writer:
        for sheet in ('SO_국내', 'PO_국내', 'DN_국내'):
            pd.DataFrame({'A': [1]}).to_excel(writer, sheet_name=sheet, index=False)

    frames = R.load_source_frames(book, ('SO_국내', 'PO_국내', 'DN_국내'))

    assert len(frames) == 3
    book.rename(tmp_path / "moved.xlsx")     # 핸들이 남아 있으면 PermissionError


# === 기간 코드 ==============================================================

@pytest.mark.parametrize('period,month', [('P01', 1), ('P8', 8), ('p12', 12)])
def test_기간_코드_파싱(period, month):
    assert R.period_month(period) == month


@pytest.mark.parametrize('bad', ['P00', 'P13', '8월', 'PP8', ''])
def test_잘못된_기간_코드는_거부한다(bad):
    with pytest.raises(ValueError):
        R.period_month(bad)


def test_정리된_상태_집합은_이번_달까지의_Invoiced와_Cancelled():
    assert R.settled_statuses(3) == {
        'Invoiced P01', 'Invoiced P02', 'Invoiced P03', 'Cancelled'}
