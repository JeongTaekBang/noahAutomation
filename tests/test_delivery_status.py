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
from unittest.mock import patch

import pandas as pd
import pytest

import delivery_status as ds
from po_generator.mail_cli import MailMode, MailOptions
from po_generator.mailer import MailConfigError, MailResult, Recipient


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
        # 기본값은 '정상'(요청납기 ≥ 공장출고일)이어야 한다.
        # 늦은 값을 기본으로 두면 무관한 테스트에까지 '요청납기 초과' 표식이 섞인다.
        'Requested delivery date': dt.datetime(2026, 9, 19),
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
        assert summary[ds.COL_EXW].tolist() == [ds.DATE_TBD_LABEL]

    def test_주문_안에_날짜가_하나면_한_행(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10)},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 8, 10)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary[ds.COL_EXW].tolist() == ['2026-08-10']
        # 안 갈린 주문에는 품목 주석을 붙이지 않는다 — 표가 시끄러워진다
        assert summary['Remarks'].tolist() == ['2026-024N']


class TestSplitDelivery:
    """분할 납기 — 한 주문 안에서 납기가 갈리면 날짜별로 행을 나눈다

    기준은 두 날짜 모두다: 고객 요청납기(Requested delivery date)와 공장 출고일(EXW NOAH).
    """

    def test_요청납기가_갈리면_나뉜다(self):
        """공장 출고일이 같아도 요청납기가 다르면 다른 약속이다"""
        so = make_so([
            {'Line item': 1, 'Item qty': 3, 'Requested delivery date': dt.datetime(2026, 9, 30)},
            {'Line item': 2, 'Item qty': 7, 'Requested delivery date': dt.datetime(2026, 8, 10)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary['Requested delivery date'].tolist() == ['2026-08-10', '2026-09-30']
        assert summary['수량'].tolist() == [7.0, 3.0]
        # 공장 출고일은 원래대로 하나
        assert set(summary[ds.COL_EXW]) == {'2026-09-14'}

    def test_요청납기가_같으면_안_나뉜다(self):
        so = make_so([
            {'Line item': 1, 'Item qty': 3},
            {'Line item': 2, 'Item qty': 7},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert len(summary) == 1
        assert summary.loc[0, '수량'] == 10.0

    def test_두_날짜가_모두_갈리면_조합만큼_나뉜다(self):
        so = make_so([
            {'Line item': 1, 'Requested delivery date': dt.datetime(2026, 8, 1),
             'EXW NOAH': dt.datetime(2026, 9, 1)},
            {'Line item': 2, 'Requested delivery date': dt.datetime(2026, 8, 1),
             'EXW NOAH': dt.datetime(2026, 10, 1)},
            {'Line item': 3, 'Requested delivery date': dt.datetime(2026, 8, 20),
             'EXW NOAH': dt.datetime(2026, 10, 1)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert len(summary) == 3
        pairs = set(zip(summary['Requested delivery date'], summary[ds.COL_EXW]))
        assert pairs == {
            ('2026-08-01', '2026-09-01'),
            ('2026-08-01', '2026-10-01'),
            ('2026-08-20', '2026-10-01'),
        }

    def test_미정이_섞이면_확정분과_분리된다(self):
        """대표 날짜 하나로 뭉개면 나머지 납기가 사라지거나 틀린 약속이 된다"""
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 8, 10), 'Item qty': 8},
            {'Line item': 2, 'EXW NOAH': dt.time(0, 0), 'Item qty': 2},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary[ds.COL_EXW].tolist() == ['2026-08-10', ds.DATE_TBD_LABEL]
        assert summary['수량'].tolist() == [8.0, 2.0]

    def test_확정_날짜가_여럿이어도_나뉜다(self):
        so = make_so([
            {'Line item': 1, 'EXW NOAH': dt.datetime(2026, 9, 30)},
            {'Line item': 2, 'EXW NOAH': dt.datetime(2026, 8, 10)},
        ])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary[ds.COL_EXW].tolist() == ['2026-08-10', '2026-09-30']

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
        note = summary[summary[ds.COL_EXW] == '2026-08-10'].iloc[0]['Remarks']
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
        confirmed = summary[summary[ds.COL_EXW] == '2026-08-10'].iloc[0]
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
        plain = summary[summary[ds.COL_EXW] == '2026-08-10'].iloc[0]
        assert plain['Remarks'] == '2026-024N'
        assert summary['Remarks'].str.contains('품목:').sum() == 2


# === 요약/상세 집계 ========================================================

class TestSummary:
    """SO_ID 단위 요약 — 고객에게 보내는 표"""

    def test_컬럼_구성(self):
        work = ds.attach_ship_status(make_so([{}]), EMPTY_DN)
        assert list(ds.build_summary(work).columns) == [
            'Customer PO', 'Remarks', '수량',
            'Requested delivery date', ds.COL_EXW,
            'Sales 금액', 'PO receipt date',
        ]

    def test_요청납기가_채워진다(self):
        work = ds.attach_ship_status(make_so([{}]), EMPTY_DN)
        assert ds.build_summary(work).loc[0, 'Requested delivery date'] == '2026-09-19'

    def test_요청납기가_비어_있으면_빈칸(self):
        so = make_so([{'Requested delivery date': None}])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary.loc[0, 'Requested delivery date'] == ''

    def test_요청납기가_문자열이어도_읽는다(self):
        """실데이터에 ISO 문자열로 들어간 셀이 38건 있다"""
        so = make_so([{'Requested delivery date': '2026-10-06'}])
        summary = ds.build_summary(ds.attach_ship_status(so, EMPTY_DN))
        assert summary.loc[0, 'Requested delivery date'] == '2026-10-06'

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
        assert summary[ds.COL_EXW].tolist() == ['2026-08-31', '2026-09-14', ds.DATE_TBD_LABEL]

    def test_요약합과_상세합이_같다(self):
        """그룹핑이 금액을 흘리거나 중복 집계하면 안 된다

        분리 기준을 두 번 바꿨다(EXW → EXW+요청납기). 키가 늘 때 merge fan-out이나
        행 누락이 생기면 고객에게 틀린 금액이 나간다.
        """
        so = make_so([
            {'SO_ID': 'SOD-1', 'Line item': 1, 'Item qty': 3, 'Sales Unit Price': 100000,
             'EXW NOAH': dt.datetime(2026, 9, 1)},
            {'SO_ID': 'SOD-1', 'Line item': 2, 'Item qty': 7, 'Sales Unit Price': 250000,
             'EXW NOAH': dt.datetime(2026, 10, 1)},
            {'SO_ID': 'SOD-1', 'Line item': 3, 'Item qty': 5, 'Sales Unit Price': 250000,
             'EXW NOAH': dt.datetime(2026, 10, 1),
             'Requested delivery date': dt.datetime(2026, 8, 1)},
            {'SO_ID': 'SOD-2', 'Line item': 1, 'Item qty': 2, 'Sales Unit Price': 500000},
        ])
        rows = ds.attach_ship_status(so, EMPTY_DN)
        summary, detail = ds.build_summary(rows), ds.build_detail(rows)
        assert summary['Sales 금액'].sum() == detail['미출고금액'].sum()
        assert summary['수량'].sum() == detail['미출고수량'].sum()
        # 전량 미출고이므로 원본 Sales amount 합과도 같아야 한다
        assert summary['Sales 금액'].sum() == 3 * 100000 + 7 * 250000 + 5 * 250000 + 2 * 500000

    def test_부분출고_출고분과_미출고분이_원본과_맞는다(self):
        so = make_so([{'Item qty': 3, 'Sales Unit Price': 1000000, 'Sales amount': 3000000}])
        dn = make_dn([{'Qty': 1}])
        rows = ds.attach_ship_status(so, dn)
        미출고 = ds.build_summary(rows).loc[0, 'Sales 금액']
        출고 = rows.iloc[0]['_단가'] * rows.iloc[0]['출고수량']
        assert 출고 + 미출고 == 3000000

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
            'PO receipt date', 'Requested delivery date', 'EXW NOAH',
            'Expected delivery date', '출고상태',
        ]
        row = detail.iloc[0]
        assert (row['주문수량'], row['출고수량'], row['미출고수량']) == (10, 4.0, 6.0)
        assert row['출고상태'] == '부분 출고'
        assert row['PO receipt date'] == '2026-01-20'
        assert row['Requested delivery date'] == '2026-09-19'

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

# === 메일 본문 ============================================================

def _summary(rows: list[dict]) -> pd.DataFrame:
    """메일 본문 테스트용 요약 (build_summary를 거쳐 실제 컬럼 구성을 얻는다)"""
    return ds.build_summary(ds.attach_ship_status(make_so(rows), EMPTY_DN))


class TestMailTables:
    """본문 표 — PO receipt date만 빼고 첨부와 같은 컬럼"""

    def test_평문표_컬럼(self):
        text = ds.build_text_table(_summary([{}]))
        header = text.splitlines()[0]
        for col in ('Customer PO', 'Remarks', '수량', ds.COL_EXW, 'Sales 금액'):
            assert col in header
        assert 'PO receipt date' not in header

    def test_금액은_천단위_구분(self):
        text = ds.build_text_table(_summary([{'Item qty': 10, 'Sales Unit Price': 123456}]))
        assert '1,234,560' in text

    def test_평문표_한글_열_정렬(self):
        """한글 비고가 섞여도 열이 어긋나면 안 된다"""
        text = ds.build_text_table(_summary([
            {'SO_ID': 'SOD-1', 'Remarks': '짧음'},
            {'SO_ID': 'SOD-2', 'Remarks': '아주 긴 한글 비고입니다'},
        ]))
        # 각주는 표가 아니므로 폭 비교에서 뺀다 (빈 줄 뒤부터가 각주)
        table = text.split('\n\n')[0]
        widths = {ds._display_width(l) for l in table.splitlines() if not l.startswith('-')}
        assert len(widths) == 1

    def test_HTML표_행수(self):
        html = ds.build_html_table(_summary([{'SO_ID': 'SOD-1'}, {'SO_ID': 'SOD-2'}]))
        assert html.count('<tr>') == 3  # 헤더 + 2행

    def test_HTML표_이스케이프(self):
        """거래처 비고에 & < 가 들어와도 표가 깨지면 안 된다 (예: 'S&T중공업')"""
        html = ds.build_html_table(_summary([{'Remarks': 'S&T중공업 <긴급>'}]))
        assert 'S&amp;T중공업 &lt;긴급&gt;' in html
        assert '<긴급>' not in html

    def test_HTML표_처리중_강조(self):
        html = ds.build_html_table(_summary([{'EXW NOAH': dt.time(0, 0)}]))
        assert ds.DATE_TBD_LABEL in html
        assert 'color:#c00000' in html


class TestLateHighlight:
    """요청납기를 넘긴 공장 출고일은 빨간색 — 고객이 가장 먼저 봐야 하는 줄"""

    LATE = {'Requested delivery date': dt.datetime(2026, 9, 14),
            'EXW NOAH': dt.datetime(2026, 10, 22)}
    ONTIME = {'Requested delivery date': dt.datetime(2026, 10, 22),
              'EXW NOAH': dt.datetime(2026, 9, 14)}

    @pytest.mark.parametrize('row,expected', [
        (LATE, True),
        (ONTIME, False),
        # 같은 날은 늦은 게 아니다
        ({'Requested delivery date': dt.datetime(2026, 9, 14),
          'EXW NOAH': dt.datetime(2026, 9, 14)}, False),
    ])
    def test_판정(self, row, expected):
        s = _summary([row])
        assert ds.is_late(s.iloc[0]) is expected

    def test_출고일_미정은_판정하지_않는다(self):
        """모르는 것을 늦었다고 표시하면 안 된다"""
        s = _summary([{'Requested delivery date': dt.datetime(2026, 1, 1),
                       'EXW NOAH': dt.time(0, 0)}])
        assert ds.is_late(s.iloc[0]) is False

    def test_요청납기_없으면_판정하지_않는다(self):
        s = _summary([{'Requested delivery date': None,
                       'EXW NOAH': dt.datetime(2026, 10, 22)}])
        assert ds.is_late(s.iloc[0]) is False

    def test_HTML은_해당_셀만_빨강_굵게(self):
        html = ds.build_html_table(_summary([self.LATE]))
        assert 'color:#c00000;font-weight:bold;">2026-10-22' in html
        # 요청납기 칸은 건드리지 않는다 (문제는 출고일이다)
        assert 'font-weight:bold;">2026-09-14' not in html

    def test_HTML_정상건은_강조_없음(self):
        html = ds.build_html_table(_summary([self.ONTIME]))
        assert 'font-weight:bold' not in html

    def test_평문은_별표와_각주(self):
        """색을 못 쓰는 대체본도 같은 정보를 담아야 한다"""
        text = ds.build_text_table(_summary([self.LATE]))
        assert '2026-10-22 *' in text
        assert '* 요청 납기일보다 공장 출고 예정일이 늦은 건' in text

    def test_평문_정상건은_각주_없음(self):
        text = ds.build_text_table(_summary([self.ONTIME]))
        assert '*' not in text

    def test_빈_요약에도_안_터진다(self):
        """지금 흐름에선 빈 요약이 메일까지 오지 않지만, 폭 계산이 빈 시퀀스에서 죽으면 안 된다"""
        empty = _summary([{}]).iloc[0:0]
        assert ds.build_text_table(empty) == ''
        html = ds.build_html_table(empty)
        assert html.count('<tr>') == 1  # 헤더만
        assert '<td' not in html


class TestHtmlBody:
    """본문 구조 — Outlook은 첫 블록 요소 뒤에 서명을 끼워 넣는다

    그래서 본문 전체가 최상위 블록 **하나** 안에 들어가야 서명이 맨 끝에 붙는다.
    """

    def test_최상위_블록이_하나다(self):
        """<body> 바로 아래에 형제 블록이 여러 개면 서명이 그 사이로 들어간다"""
        html = ds.build_html_body(f'인사말\n\n{ds._HTML_TABLE_TOKEN}\n\n맺음말', '<table>T</table>')
        inner = html[html.index('<body>') + len('<body>'):html.index('</body>')]
        assert inner.startswith('<table role="presentation"')
        assert inner.endswith('</table>')
        # 래퍼를 벗기면 그 안에 문단·표가 들어 있다
        assert inner.count('<td ') == 1

    def test_문단이_블록요소로_나온다(self):
        html = ds.build_html_body(f'인사말\n\n{ds._HTML_TABLE_TOKEN}\n\n맺음말', '<table></table>')
        assert html.count('<p ') == 2
        assert '인사말</p>' in html and '맺음말</p>' in html

    def test_표가_문단_사이_제자리에_들어간다(self):
        html = ds.build_html_body(f'앞\n\n{ds._HTML_TABLE_TOKEN}\n\n뒤', '<table>T</table>')
        assert html.index('앞') < html.index('<table>T') < html.index('뒤')

    def test_본문_텍스트는_이스케이프된다(self):
        html = ds.build_html_body('S&T중공업 <긴급>', '')
        assert 'S&amp;T중공업 &lt;긴급&gt;' in html

    def test_문단_안_줄바꿈은_유지(self):
        html = ds.build_html_body('첫줄\n둘째줄', '')
        assert '첫줄<br>둘째줄' in html

    def test_토큰이_남지_않는다(self):
        html = ds.build_html_body(f'앞\n\n{ds._HTML_TABLE_TOKEN}\n\n뒤', '<table></table>')
        assert ds._HTML_TABLE_TOKEN not in html


class TestMailSummary:
    """발송 흐름 — 메일 실패가 문서 생성을 뒤엎지 않는다"""

    @pytest.fixture
    def recipient(self):
        return Recipient(
            biz_no='6158188675', customer_name='엔이에스',
            to=('nes@example.com',), cc=('cc@example.com',),
        )

    @pytest.fixture
    def summary(self):
        return _summary([{}])

    def _opts(self, mode=MailMode.DRAFT):
        # df_customer가 채워져 있으면 customer_master()가 Excel을 읽지 않는다
        return MailOptions(mode=mode, df_customer=pd.DataFrame({'x': [1]}))

    def test_메일_꺼져_있으면_아무것도_안_한다(self, summary, tmp_path):
        with patch.object(ds, 'find_recipient') as finder:
            assert ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A',
                                   MailOptions.disabled()) is True
        finder.assert_not_called()

    def test_수신자_미등록이면_False(self, summary, tmp_path, capsys):
        with patch.object(ds, 'find_recipient', return_value=None):
            assert ds.mail_summary(summary, tmp_path / 'a.xlsx', '6158188675', '엔이에스',
                                   self._opts()) is False
        assert '수신자 미등록' in capsys.readouterr().out

    def test_설정오류는_문서를_뒤엎지_않는다(self, summary, tmp_path, capsys):
        with patch.object(ds, 'find_recipient', side_effect=MailConfigError('이메일 컬럼 없음')):
            assert ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A', self._opts()) is False
        assert '이메일 컬럼 없음' in capsys.readouterr().out

    def test_확인에서_거부하면_보내지_않는다(self, summary, recipient, tmp_path, capsys):
        with patch.object(ds, 'find_recipient', return_value=recipient), \
             patch.object(ds, 'confirm', return_value=False), \
             patch.object(ds, 'create_document_mail') as sender:
            assert ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A',
                                   self._opts(MailMode.ASK)) is False
        sender.assert_not_called()
        assert '메일 생략' in capsys.readouterr().out

    def test_수신자를_먼저_보여준다(self, summary, recipient, tmp_path, capsys):
        """오발송 차단 — 누구에게 나가는지 확인 전에 화면에 찍혀야 한다"""
        with patch.object(ds, 'find_recipient', return_value=recipient), \
             patch.object(ds, 'confirm', return_value=False):
            ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A', self._opts(MailMode.ASK))
        out = capsys.readouterr().out
        assert 'nes@example.com' in out and 'cc@example.com' in out

    def test_성공_경로(self, summary, recipient, tmp_path):
        result = MailResult(success=True, sent=False, recipient=recipient,
                            attachments=(tmp_path / 'a.xlsx',))
        with patch.object(ds, 'find_recipient', return_value=recipient), \
             patch.object(ds, 'create_document_mail', return_value=result) as sender:
            assert ds.mail_summary(summary, tmp_path / 'a.xlsx', '6158188675', '엔이에스',
                                   self._opts()) is True
        kwargs = sender.call_args.kwargs
        assert kwargs['attach_format'] == ds.DS_MAIL_ATTACH_FORMAT
        assert kwargs['doc_label'] == '납기현황'

    def test_HTML본문에_토큰이_남지_않는다(self, summary, recipient, tmp_path):
        """평문을 이스케이프한 뒤 표를 되돌리는데, 순서가 틀리면 토큰이 그대로 나간다"""
        with patch.object(ds, 'find_recipient', return_value=recipient), \
             patch.object(ds, 'create_document_mail',
                          return_value=MailResult(True, False, recipient)) as sender:
            ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A', self._opts())
        html = sender.call_args.kwargs['body_html']
        assert ds._HTML_TABLE_TOKEN not in html
        assert '<table' in html and '&lt;table' not in html

    def test_본문_치환자가_채워진다(self, summary, recipient, tmp_path):
        with patch.object(ds, 'find_recipient', return_value=recipient), \
             patch.object(ds, 'create_document_mail',
                          return_value=MailResult(True, False, recipient)) as sender:
            ds.mail_summary(summary, tmp_path / 'a.xlsx', '1', 'A', self._opts())
        extra = sender.call_args.kwargs['extra']
        assert extra['count'] == len(summary)
        assert 'Customer PO' in extra['table']


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
