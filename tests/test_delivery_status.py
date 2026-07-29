"""
delivery_status.py 테스트
==========================

거래처 납기현황 조회 — 출고상태 파생, 잔량 계산, 거래처 조회, 요약 집계.

핵심 회귀 방지 대상은 **시트의 캐시 `Status`를 믿지 않는 것**이다.
`Status`는 파워쿼리 결과의 XLOOKUP 캐시라 새로고침이 밀리면 실제 출고와 어긋나고,
그대로 회신하면 이미 납품한 건을 "미출고"라고 고객에게 알리게 된다.
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

import delivery_status as ds


# === 픽스처 ================================================================

BIZ_A = '615-81-88675'
BIZ_B = '312-81-03840'


def make_so(rows: list[dict]) -> pd.DataFrame:
    """SO_국내 형태의 DataFrame 생성 (빠진 컬럼은 기본값으로 채움)"""
    defaults = {
        'SO_ID': 'SOD-2026-0001',
        'Line item': 1,
        'Business registration number': BIZ_A,
        'Customer name': '엔이에스 주식회사',
        'Customer PO': 'PO-001',
        'Item name': 'NOS100-M',
        'Item qty': 10,
        'Sales Unit Price': 100000,
        'Sales amount': 1000000,
        'PO receipt date': dt.datetime(2026, 1, 20),
        'EXW NOAH': dt.datetime(2026, 9, 14),
        'Expected delivery date': dt.datetime(2026, 9, 19),
        'Remarks': '2026-024N',
        'Status': '미출고',
    }
    return pd.DataFrame([{**defaults, **row} for row in rows])


def make_dn(rows: list[dict]) -> pd.DataFrame:
    """DN_국내 형태의 DataFrame 생성"""
    defaults = {
        'DN_ID': 'DND-2026-0001',
        'SO_ID': 'SOD-2026-0001',
        'Line item': 1,
        'Qty': 10,
        '출고일': dt.datetime(2026, 7, 27),
    }
    return pd.DataFrame([{**defaults, **row} for row in rows])


EMPTY_DN = pd.DataFrame(columns=['DN_ID', 'SO_ID', 'Line item', 'Qty', '출고일'])


def status_of(so: pd.DataFrame, dn: pd.DataFrame) -> list[str]:
    """계산된 출고상태 목록"""
    return list(ds.attach_ship_status(so, dn)['_출고상태'])


# === 출고상태 파생 =========================================================

class TestShipStatus:
    """파워쿼리 `SO_통합[출고완료]`와 같은 4단계 판정"""

    def test_dn_없으면_미출고(self):
        assert status_of(make_so([{}]), EMPTY_DN) == ['미출고']

    def test_수량_모자라면_부분출고(self):
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 4}])
        assert status_of(so, dn) == ['부분 출고']

    def test_수량_채웠고_출고일_있으면_출고완료(self):
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 10}])
        assert status_of(so, dn) == ['출고 완료']

    def test_수량_채웠지만_출고일_없으면_공장출고(self):
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 10, '출고일': None}])
        assert status_of(so, dn) == ['공장 출고']

    def test_분할납품은_합산해서_판정(self):
        """한 SO 라인을 여러 DN으로 나눠 출고하면 합계로 봐야 한다"""
        so = make_so([{'Item qty': 10}])
        dn = make_dn([
            {'DN_ID': 'DND-2026-0001', 'Qty': 6},
            {'DN_ID': 'DND-2026-0002', 'Qty': 4},
        ])
        assert status_of(so, dn) == ['출고 완료']

    def test_수량0_DN이_있으면_미출고가_아니라_부분출고(self):
        """파워쿼리는 `[출고수량] = null`로 미출고를 가른다.

        수량 0짜리 DN 행이 있으면 '출고 기록은 있다'는 뜻이므로 부분 출고가 맞다.
        수량으로만 판정하면 이 행이 미출고로 잘못 분류된다.
        """
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 0, '출고일': None}])
        assert status_of(so, dn) == ['부분 출고']

    def test_무상공급_부분출고도_수량으로_잡는다(self):
        """단가 0(FOC)이라 금액 기준으로는 영영 못 잡는 케이스"""
        so = make_so([{'Item qty': 576, 'Sales Unit Price': 0, 'Sales amount': 0}])
        dn = make_dn([{'Qty': 432}])
        work = ds.attach_ship_status(so, dn)
        assert work['_출고상태'].tolist() == ['부분 출고']
        assert work['_미출고수량'].tolist() == [144.0]
        assert work['_미출고금액'].tolist() == [0.0]

    def test_라인아이템_타입이_달라도_조인된다(self):
        """SO는 int, DN은 float/str로 읽히는 경우가 있다"""
        so = make_so([{'Line item': 2, 'Item qty': 10}])
        dn = make_dn([{'Line item': 2.0, 'Qty': 10}])
        assert status_of(so, dn) == ['출고 완료']

    def test_다른_라인의_출고는_섞이지_않는다(self):
        so = make_so([
            {'Line item': 1, 'Item qty': 10},
            {'Line item': 2, 'Item qty': 10},
        ])
        dn = make_dn([{'Line item': 1, 'Qty': 10}])
        assert status_of(so, dn) == ['출고 완료', '미출고']


class TestStaleStatusIgnored:
    """시트 캐시 Status를 판정에 쓰지 않는다"""

    def test_캐시가_미출고여도_DN이_있으면_출고완료(self):
        """새로고침이 밀려 Status가 '미출고'로 남은 실제 케이스 (삼신 SOD-2026-0651 등)"""
        so = make_so([{'Item qty': 10, 'Status': '미출고'}])
        dn = make_dn([{'Qty': 10}])
        work = ds.attach_ship_status(so, dn)
        assert work['_출고상태'].tolist() == ['출고 완료']
        # 회신 대상에서 빠져야 한다 — 이미 납품한 건이다
        assert ds.select_rows(work, ds.normalize_biz_no(BIZ_A), include_shipped=False).empty

    def test_캐시가_출고완료여도_잔량이_있으면_회신대상(self):
        so = make_so([{'Item qty': 54, 'Status': '출고 완료'}])
        dn = make_dn([{'Qty': 43}])
        work = ds.attach_ship_status(so, dn)
        rows = ds.select_rows(work, ds.normalize_biz_no(BIZ_A), include_shipped=False)
        assert rows['_출고상태'].tolist() == ['부분 출고']
        assert rows['_미출고수량'].tolist() == [11.0]

    def test_어긋난_행_수를_센다(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Item qty': 10, 'Status': '미출고'},     # 실제로는 출고완료 → 어긋남
            {'SO_ID': 'SOD-2', 'Item qty': 10, 'Status': '미출고'},     # DN 없음 → 일치
            {'SO_ID': 'SOD-3', 'Item qty': 10, 'Status': 'Cancelled'},  # 취소는 비교 제외
            {'SO_ID': 'SOD-4', 'Item qty': 10, 'Status': None},         # 수식 미계산은 비교 제외
        ])
        dn = make_dn([{'SO_ID': 'SOD-1', 'Qty': 10}])
        assert ds.count_stale_status(ds.attach_ship_status(so, dn)) == 1

    def test_빈_SO에도_깨지지_않는다(self):
        """빈 프레임에 apply(axis=1)을 걸면 Series가 아니라 DataFrame이 돌아온다"""
        empty_so = make_so([{}]).iloc[0:0]
        work = ds.attach_ship_status(empty_so, EMPTY_DN)
        assert work.empty
        assert '_출고상태' in work.columns
        assert ds.build_summary(work).empty
        assert ds.build_customer_list(work).empty

    def test_Status_컬럼이_아예_없어도_동작한다(self):
        so = make_so([{}]).drop(columns=['Status'])
        work = ds.attach_ship_status(so, EMPTY_DN)
        assert work['_출고상태'].tolist() == ['미출고']
        assert ds.count_stale_status(work) == 0
        assert len(ds.select_rows(work, ds.normalize_biz_no(BIZ_A), include_shipped=False)) == 1


class TestSelectRows:
    """회신 대상 선별"""

    def test_취소건은_제외한다(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Status': '미출고'},
            {'SO_ID': 'SOD-2', 'Status': 'Cancelled'},
            {'SO_ID': 'SOD-3', 'Status': 'Hold'},
        ])
        work = ds.attach_ship_status(so, EMPTY_DN)
        rows = ds.select_rows(work, ds.normalize_biz_no(BIZ_A), include_shipped=False)
        assert sorted(rows['_so_id']) == ['SOD-1']

    def test_다른_거래처는_섞이지_않는다(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Business registration number': BIZ_A},
            {'SO_ID': 'SOD-2', 'Business registration number': BIZ_B},
        ])
        work = ds.attach_ship_status(so, EMPTY_DN)
        rows = ds.select_rows(work, ds.normalize_biz_no(BIZ_A), include_shipped=False)
        assert sorted(rows['_so_id']) == ['SOD-1']

    def test_all_옵션은_출고완료도_포함(self):
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 10}])
        work = ds.attach_ship_status(so, dn)
        biz = ds.normalize_biz_no(BIZ_A)
        assert ds.select_rows(work, biz, include_shipped=False).empty
        assert len(ds.select_rows(work, biz, include_shipped=True)) == 1


# === 날짜 처리 =============================================================

class TestDateHandling:
    """`EXW NOAH`는 빈 칸이 NaN이 아니라 time(0,0)으로 읽힌다"""

    @pytest.mark.parametrize('value', [
        None, pd.NaT, float('nan'), '', '   ', dt.time(0, 0),
    ])
    def test_빈값은_None(self, value):
        assert ds.as_date(value) is None

    def test_날짜는_Timestamp로(self):
        assert ds.as_date(dt.datetime(2026, 9, 14)) == pd.Timestamp('2026-09-14')

    def test_빈_EXW는_처리중으로_찍힌다(self):
        so = make_so([{'EXW NOAH': dt.time(0, 0)}])
        work = ds.attach_ship_status(so, EMPTY_DN)
        summary = ds.build_summary(work)
        assert summary['NOAH 공장 출고일'].tolist() == [ds.DATE_TBD_LABEL]

    def test_주문_안에_날짜가_하나면_한_행(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10)},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 8, 10)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['NOAH 공장 출고일'].tolist() == ['2026-08-10']
        # 안 갈린 주문에는 품목 주석을 붙이지 않는다 — 표가 시끄러워진다
        assert summary['Remarks'].tolist() == ['2026-024N']


class TestSplitDelivery:
    """분할 납기 — 한 주문 안에서 EXW NOAH가 갈리면 날짜별로 행을 나눈다"""

    def test_미정이_섞이면_확정분과_분리된다(self):
        """대표 날짜 하나로 뭉개면 나머지 납기가 사라지거나 틀린 약속이 된다"""
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item qty': 8},
            {'Line item': 2, 'EXW NOAH': dt.time(0, 0), 'Item qty': 2},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['NOAH 공장 출고일'].tolist() == ['2026-08-10', ds.DATE_TBD_LABEL]
        assert summary['수량'].tolist() == [8.0, 2.0]

    def test_확정_날짜가_여럿이어도_나뉜다(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 9, 30)},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 8, 10)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['NOAH 공장 출고일'].tolist() == ['2026-08-10', '2026-09-30']

    def test_수량과_금액이_날짜별로_갈라진다(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10),
             'Item qty': 8, 'Sales Unit Price': 100000},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 9, 30),
             'Item qty': 2, 'Sales Unit Price': 500000},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['수량'].tolist() == [8.0, 2.0]
        assert summary['Sales 금액'].tolist() == [800000.0, 1000000.0]

    def test_비고가_겹치면_품목이_붙는다(self):
        """비스타콘트롤 케이스 — 비고가 같아 품목이 유일한 단서"""
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10),
             'Item name': 'NOS185-M-FC', 'Remarks': 'Rabigh#2'},
            {'Line item': 2, 'EXW NOAH': dt.time(0, 0),
             'Item name': 'Eye bolt', 'Remarks': 'Rabigh#2'},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['Remarks'].tolist() == [
            'Rabigh#2 (품목: NOS185-M-FC)',
            'Rabigh#2 (품목: Eye bolt)',
        ]

    def test_비고로_이미_구분되면_품목을_붙이지_않는다(self):
        """세진밸브 케이스 — 호선별로 비고가 갈려 있어 품목은 잡음"""
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 14),
             'Item name': 'NA015', 'Remarks': 'H2734'},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 10, 16),
             'Item name': 'NA015', 'Remarks': 'H2735'},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['Remarks'].tolist() == ['H2734', 'H2735']

    def test_비고가_비어도_품목만_붙는다(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': 'A', 'Remarks': ''},
            {'Line item': 2, 'EXW NOAH': dt.time(0, 0), 'Item name': 'B', 'Remarks': ''},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['Remarks'].tolist() == ['품목: A', '품목: B']

    def test_품목이_길면_길이로_접는다(self):
        """피엠에스는 품목명이 'MA02/0.75kW/43RPM/ON-OFF/SBWG-04-1SM'처럼 길다.

        개수로 자르면(예: 3개) 회신 표가 깨지므로 길이로 자른다.
        """
        long_names = [f'MA{n:02d}/0.75kW/43RPM/ON-OFF/SBWG-04-1SM' for n in range(1, 6)]
        so = make_so(
            [{'Line item': 1, 'EXW NOAH': dt.time(0, 0), 'Item name': 'Z'}]
            + [{'Line item': i + 2, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': name}
               for i, name in enumerate(long_names)]
        )
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        note = summary[summary['NOAH 공장 출고일'] == '2026-08-10'].iloc[0]['Remarks']
        assert long_names[0] in note          # 첫 품목은 반드시 남는다
        assert '외 ' in note and '종' in note   # 나머지는 종수로 접힌다
        assert len(note) < 110                # 표가 깨지지 않을 길이

    def test_품목이_짧으면_다_적는다(self):
        so = make_so(
            [{'Line item': 1, 'EXW NOAH': dt.time(0, 0), 'Item name': 'Z'}]
            + [{'Line item': n, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': f'ITEM-{n}'}
               for n in range(2, 5)]
        )
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        confirmed = summary[summary['NOAH 공장 출고일'] == '2026-08-10'].iloc[0]
        assert confirmed['Remarks'] == '2026-024N (품목: ITEM-2, ITEM-3, ITEM-4)'

    def test_품목명_하나가_한도보다_길어도_남긴다(self):
        """전부 '외 N종'만 나오면 단서가 없다"""
        huge = 'X' * (ds.SPLIT_ITEM_NOTE_MAX_CHARS + 40)
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': huge},
            {'Line item': 2, 'EXW NOAH': dt.time(0, 0), 'Item name': 'B'},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert huge in summary.loc[0, 'Remarks']

    def test_같은_품목이_여러_라인이면_한_번만_적는다(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': 'NOS185'},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item name': 'NOS185'},
            {'Line item': 3, 'EXW NOAH': dt.time(0, 0), 'Item name': 'Eye bolt'},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary.loc[0, 'Remarks'] == '2026-024N (품목: NOS185)'

    def test_다른_주문의_날짜는_영향을_주지_않는다(self):
        """SOD-1은 날짜가 하나뿐이니 품목 주석이 붙으면 안 된다"""
        so = make_so([
            {'SO_ID': 'SOD-1', 'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10)},
            {'SO_ID': 'SOD-2', 'Line item': 1, 'EXW NOAH': dt.datetime(2026, 9, 30)},
            {'SO_ID': 'SOD-2', 'Line item': 2, 'EXW NOAH': dt.time(0, 0)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert len(summary) == 3
        plain = summary[summary['NOAH 공장 출고일'] == '2026-08-10'].iloc[0]
        assert plain['Remarks'] == '2026-024N'
        assert summary['Remarks'].str.contains('품목:').sum() == 2


# === 요약/상세 집계 ========================================================

class TestSummary:
    """SO_ID 단위 요약 — 고객에게 보내는 표"""

    def test_컬럼_구성(self):
        work = ds.attach_ship_status(make_so([{}]), EMPTY_DN)
        assert list(ds.build_summary(work).columns) == [
            'Customer PO', 'Remarks', '수량', 'NOAH 공장 출고일', 'Sales 금액', 'PO receipt date',
        ]

    def test_같은_SO는_한_행으로_합산(self):
        so = make_so([
            {'Line item': 1, 'Item qty': 8, 'Sales Unit Price': 100000},
            {'Line item': 2, 'Item qty': 2, 'Sales Unit Price': 200000},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert len(summary) == 1
        assert summary.loc[0, '수량'] == 10
        assert summary.loc[0, 'Sales 금액'] == 8 * 100000 + 2 * 200000

    def test_SO가_다르면_같은_PO여도_행이_나뉜다(self):
        """호선별로 SO가 갈리는 주문 — 메일 표도 SO 단위로 나뉘어 있다"""
        so = make_so([
            {'SO_ID': 'SOD-1', 'Customer PO': 'PO-9', 'Remarks': 'SN00295호선'},
            {'SO_ID': 'SOD-2', 'Customer PO': 'PO-9', 'Remarks': 'SN00296호선'},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert len(summary) == 2
        assert sorted(summary['Remarks']) == ['SN00295호선', 'SN00296호선']

    def test_금액은_잔량_기준(self):
        """부분출고 건은 남은 수량만큼만 금액을 잡는다"""
        so = make_so([{'Item qty': 10, 'Sales Unit Price': 100000}])
        dn = make_dn([{'Qty': 4}])
        summary = ds.build_summary(ds.attach_ship_status(so, dn))
        assert summary.loc[0, '수량'] == 6
        assert summary.loc[0, 'Sales 금액'] == 600000

    def test_납기_확정건이_처리중보다_앞에_온다(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'EXW NOAH': dt.time(0, 0)},
            {'SO_ID': 'SOD-2', 'EXW NOAH': dt.datetime(2026, 9, 14)},
            {'SO_ID': 'SOD-3', 'EXW NOAH': dt.datetime(2026, 8, 31)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['NOAH 공장 출고일'].tolist() == ['2026-08-31', '2026-09-14', ds.DATE_TBD_LABEL]

    def test_빈_입력은_빈_요약(self):
        empty = ds.attach_ship_status(make_so([{}]), EMPTY_DN).iloc[0:0]
        summary = ds.build_summary(empty)
        assert summary.empty
        assert '수량' in summary.columns


class TestDetail:
    """라인 단위 상세 — 내부 확인용"""

    def test_컬럼_구성과_값(self):
        so = make_so([{'Item qty': 10, 'Sales Unit Price': 100000}])
        dn = make_dn([{'Qty': 4}])
        detail = ds.build_detail(ds.attach_ship_status(so, dn))
        assert list(detail.columns) == [
            'SO_ID', 'Line item', 'Customer PO', 'Remarks', 'Item name',
            '주문수량', '출고수량', '미출고수량', 'Sales Unit Price', '미출고금액',
            'PO receipt date', 'EXW NOAH', 'Expected delivery date', '출고상태',
        ]
        row = detail.iloc[0]
        assert (row['주문수량'], row['출고수량'], row['미출고수량']) == (10, 4.0, 6.0)
        assert row['출고상태'] == '부분 출고'
        assert row['PO receipt date'] == '2026-01-20'

    def test_라인번호_순으로_정렬(self):
        so = make_so([{'Line item': n} for n in (10, 2, 1)])
        detail = ds.build_detail(ds.attach_ship_status(so, EMPTY_DN))
        assert detail['Line item'].tolist() == ['1', '2', '10']


# === 거래처 조회 ===========================================================

class TestResolveCustomer:
    """사업자번호 또는 거래처명으로 거래처 확정"""

    @pytest.fixture
    def work(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Business registration number': BIZ_A, 'Customer name': '엔이에스 주식회사'},
            {'SO_ID': 'SOD-2', 'Business registration number': BIZ_B, 'Customer name': '(주)삼신'},
        ])
        return ds.attach_ship_status(so, EMPTY_DN)

    @pytest.mark.parametrize('query', ['615-81-88675', '6158188675', '615 81 88675'])
    def test_사업자번호_표기_차이를_흡수(self, work, query):
        hits, kind = ds.resolve_customer(work, query)
        assert hits == [('6158188675', '엔이에스 주식회사')]
        assert kind == '사업자번호로'

    def test_거래처명_부분일치(self, work):
        hits, kind = ds.resolve_customer(work, '엔이에스')
        assert hits == [('6158188675', '엔이에스 주식회사')]
        assert kind == '거래처명으로'

    def test_후보가_여럿이면_모두_반환(self, work):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Business registration number': BIZ_A, 'Customer name': '가나콘트롤'},
            {'SO_ID': 'SOD-2', 'Business registration number': BIZ_B, 'Customer name': '다라콘트롤'},
        ])
        hits, _ = ds.resolve_customer(ds.attach_ship_status(so, EMPTY_DN), '콘트롤')
        assert len(hits) == 2

    def test_없으면_빈_목록(self, work):
        assert ds.resolve_customer(work, '999-99-99999')[0] == []
        assert ds.resolve_customer(work, '존재하지않는회사')[0] == []

    def test_같은_사업자번호에_표기가_여럿이면_최빈값(self):
        so = make_so([
            {'SO_ID': 'SOD-1', 'Customer name': '엔이에스 주식회사'},
            {'SO_ID': 'SOD-2', 'Customer name': '엔이에스 주식회사'},
            {'SO_ID': 'SOD-3', 'Customer name': '엔이에스(주)'},
        ])
        hits, _ = ds.resolve_customer(ds.attach_ship_status(so, EMPTY_DN), BIZ_A)
        assert hits == [('6158188675', '엔이에스 주식회사')]


class TestFormatBizNo:
    def test_10자리는_하이픈_표기(self):
        assert ds.format_biz_no('6158188675') == '615-81-88675'

    def test_그_외는_원문_유지(self):
        assert ds.format_biz_no('12345') == '12345'


# === 거래처 목록 (--list) ==================================================

class TestCustomerList:
    def test_미출고가_있는_거래처만_금액순으로(self):
        so = make_so([
            # A: 미출고 10 x 100,000
            {'SO_ID': 'SOD-1', 'Business registration number': BIZ_A,
             'Customer name': 'A사', 'Item qty': 10, 'Sales Unit Price': 100000},
            # B: 미출고 10 x 500,000 (금액이 더 큼 → 위로)
            {'SO_ID': 'SOD-2', 'Business registration number': BIZ_B,
             'Customer name': 'B사', 'Item qty': 10, 'Sales Unit Price': 500000},
        ])
        listing = ds.build_customer_list(ds.attach_ship_status(so, EMPTY_DN))
        assert listing['거래처명'].tolist() == ['B사', 'A사']
        assert listing['미출고금액'].tolist() == [5000000.0, 1000000.0]
        assert listing['주문건수'].tolist() == [1, 1]

    def test_전량_출고된_거래처는_빠진다(self):
        so = make_so([{'Item qty': 10}])
        dn = make_dn([{'Qty': 10}])
        assert ds.build_customer_list(ds.attach_ship_status(so, dn)).empty


# === 표시 폭 ===============================================================

class TestPad:
    """한글은 콘솔에서 2칸을 차지한다 — 글자 수로 패딩하면 열이 어긋난다"""

    def test_한글_폭_계산(self):
        assert ds._display_width('가나') == 4
        assert ds._display_width('ab') == 2

    def test_폭에_맞춰_패딩(self):
        assert ds._display_width(ds._pad('가나', 10)) == 10
        assert ds._display_width(ds._pad('abcd', 10)) == 10

    def test_넘치면_말줄임(self):
        out = ds._pad('가나다라마바사', 6)
        assert '…' in out
        assert ds._display_width(out) == 6
