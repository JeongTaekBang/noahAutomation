"""
거래명세표 PO No. 칸 — 호선명·현장명(SO Remarks) 병기 테스트
============================================================

하단 PO No. 옆과 본문 비고(C열)에 `SO_국내.Remarks`(호선명·현장명)를 싣는다.
처음엔 조선 기자재 거래처만 지정했지만(2026-08-25), 2026-11-03부터 **모든 거래처**다 —
Remarks는 고객에게 보이는 칸이고 내부 메모는 `SO_국내.Comment`에 적기로 했다.

지켜야 할 것 둘:
- **SO Comment(내부 메모)는 절대 거래명세표에 실리지 않는다** — `load_dn_data()`가
  SO 컬럼을 명시 목록으로만 가져오는 것이 그 보장이다.
- **DN 자체 Remarks(월합 세금계산서 문구)는 절대 PO 칸에 새지 않는다** — 이름이 같아서
  헷갈리기 쉬운 자리다 (`load_dn_data()`가 SO 쪽을 'SO Remarks'로 개명하는 이유).

품명(B열)에는 `SO_국내.Model code`(고객 품번)를 앞에 붙인다 (2026-11-04 — 에스제이브이).

`build_po_cell_text()`·`build_item_description()`은 COM 없는 순수 함수라 Excel 없이 검증한다.
"""

import pandas as pd
import pytest

from po_generator.config import DN_DOMESTIC_SHEET, PMT_DOMESTIC_SHEET, SO_DOMESTIC_SHEET
from po_generator.ts_generator import (
    build_item_description,
    build_po_cell_text,
    build_row_remark,
    resolve_so_remark_column,
)
from po_generator import utils as po_utils


SCANTECH = '주식회사 스칸텍'          # 시트 실표기 — 법인 접두가 붙는다
NES = '엔이에스 주식회사'             # 접미가 붙는 경우
CK = '씨앤케이엔지니어링(C&K ENG)'    # 예전 비대상 거래처 — 지금은 똑같이 붙는다


def order(customer=SCANTECH, po='OR26060022'):
    return pd.Series({'Customer name': customer, 'Customer PO': po})


def items(*rows):
    return pd.DataFrame(list(rows))


# === 호선명 병기 (DN 경로 = 'SO Remarks') ===

class TestVesselNameAppended:
    def test_single_po_with_vessel(self):
        """실측 DND-2026-0788: OR26060022 (H-8327)"""
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'OR26060022', 'SO Remarks': 'H-8327'},
        ))
        assert result == 'OR26060022 (H-8327)'

    def test_one_po_many_vessels_listed(self):
        """실측 DND-2026-0790: 한 발주에 호선 셋 — 하나라도 접으면 남은 배를 못 찾는다"""
        result = build_po_cell_text(order(po='SCT2605-134'), items(
            {'Customer PO': 'SCT2605-134', 'SO Remarks': '한화-H4394'},
            {'Customer PO': 'SCT2605-134', 'SO Remarks': '한화-H4395'},
            {'Customer PO': 'SCT2605-134', 'SO Remarks': '한화-H4396'},
        ))
        assert result == 'SCT2605-134 (한화-H4394, 한화-H4395, 한화-H4396)'

    def test_duplicate_vessel_not_repeated(self):
        """같은 호선의 라인이 여럿이어도 호선명은 한 번만"""
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'OR26060022', 'SO Remarks': 'H-8327'},
            {'Customer PO': 'OR26060022', 'SO Remarks': 'H-8327'},
        ))
        assert result == 'OR26060022 (H-8327)'

    def test_po_without_remark_stays_bare(self):
        """비고가 빈 발주는 지금처럼 번호만 (엔이에스 실측 DND-2026-0752)"""
        result = build_po_cell_text(order(NES, '2026-07-31-CH-01'), items(
            {'Customer PO': '2026-07-31-CH-01', 'SO Remarks': None},
        ))
        assert result == '2026-07-31-CH-01'

    def test_merged_doc_pairs_po_with_own_vessel(self):
        """월합: 발주번호마다 제 호선이 붙는다 — 통째로 뒤에 나열하면 짝을 모른다"""
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'OR26060022', 'SO Remarks': 'H-8327'},
            {'Customer PO': 'OR26060013', 'SO Remarks': 'SN2755'},
            {'Customer PO': 'SCT2605-002', 'SO Remarks': None},
        ))
        assert result == 'OR26060022 (H-8327), OR26060013 (SN2755), SCT2605-002'

    @pytest.mark.parametrize('customer', [
        SCANTECH, '주식회사 브이티엘', NES, '주식회사 파나시아', '한라IMS(주)',
        '제이케이콘트롤', CK, '',
    ])
    def test_every_customer_shows_remarks(self, customer):
        """거래처를 가리지 않는다 (2026-11-03 — 예전 지정 6곳 밖의 거래처·고객명 빈 행도)"""
        result = build_po_cell_text(order(customer), items(
            {'Customer PO': 'PO-1', 'SO Remarks': 'S590'},
        ))
        assert result == 'PO-1 (S590)'

    def test_site_name_appended(self):
        """현장 단위 거래처는 현장명이 실린다 (실측 DND-2026-0717 제이케이콘트롤)"""
        result = build_po_cell_text(order('제이케이콘트롤', 'JK2026071601'), items(
            {'Customer PO': 'JK2026071601', 'SO Remarks': 'SK 울산 데이터센터'},
        ))
        assert result == 'JK2026071601 (SK 울산 데이터센터)'


# === 붙으면 안 되는 자리 ===

class TestVesselNameNotAppended:
    def test_dn_own_remarks_never_leak(self):
        """DN 자체 Remarks(월합 문구)는 컬럼명이 같아도 PO 칸에 새지 않는다

        'SO Remarks'가 없고 'Remarks'만 있는 DN 아이템 — 그 'Remarks'는
        '25일 마감, 월합세금계산서' 같은 문구라 절대 발주번호 옆에 실리면 안 된다.
        """
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'OR26060022', 'Remarks': '25일 마감, 월합세금계산서'},
        ))
        assert result == 'OR26060022'

    def test_adv_path_uses_so_native_remarks(self):
        """선수금(ADV) 경로는 SO_국내 행 그대로라 'Remarks'가 곧 호선명이다"""
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'OR26060022', 'Remarks': 'H-8327'},
        ), doc_type='ADV')
        assert result == 'OR26060022 (H-8327)'

    def test_unknown_doc_type_falls_back_to_plain(self):
        result = build_po_cell_text(order(), items(
            {'Customer PO': 'PO-1', 'SO Remarks': 'H-8327'},
        ), doc_type='PMT')
        assert result == 'PO-1'


# === 본문 비고(C열)의 호선명 ===

class TestRowRemark:
    """행별 비고 우선순위: 월합 PO > 선수금 > 호선명 > 빈 값

    호선명은 **비고가 비어 있을 자리에만** 들어간다 — 월합의 행별 발주번호와
    선수금 표기는 문서의 성격을 말하는 기존 정보라 밀어내면 안 된다.
    """

    def test_vessel_fills_empty_remark(self):
        """일반 DN: 행마다 제 호선 (배 3척짜리 문서에서 행 구분이 목적)"""
        item = pd.Series({'Customer PO': 'SCT2605-134', 'SO Remarks': '한화-H4394'})
        assert build_row_remark(item, so_remark_col='SO Remarks') == '한화-H4394'

    def test_no_vessel_stays_empty(self):
        item = pd.Series({'Customer PO': 'PO-1', 'SO Remarks': None})
        assert build_row_remark(item, so_remark_col='SO Remarks') == ''

    def test_no_remark_column_stays_empty(self):
        """SO 비고 컬럼이 없는 경로(so_remark_col=None)는 기존 그대로 빈 비고"""
        item = pd.Series({'Customer PO': 'PO-1', 'SO Remarks': 'H-8327'})
        assert build_row_remark(item, so_remark_col=None) == ''

    def test_merged_po_beats_vessel(self):
        """월합은 행별 발주번호가 우선 — 발주↔호선 짝은 하단 PO No.가 보인다"""
        item = pd.Series({'Customer PO': 'OR26060022', 'SO Remarks': 'H-8327'})
        result = build_row_remark(item, use_po_as_remark=True, so_remark_col='SO Remarks')
        assert result == 'OR26060022'

    def test_advance_marker_beats_vessel(self):
        """선수금 표기는 문서 성격이다 — 호선명이 밀어내면 안 된다"""
        item = pd.Series({'Customer PO': 'PO-1', 'Remarks': 'H-8327'})
        result = build_row_remark(item, remark='선수금', so_remark_col='Remarks')
        assert result == '선수금'


class TestResolveSoRemarkColumn:
    """하단 병기와 본문 비고가 공유하는 관문"""

    def test_dn_path(self):
        df = items({'Customer PO': 'PO-1', 'SO Remarks': 'H-8327'})
        assert resolve_so_remark_column(df) == 'SO Remarks'

    def test_adv_path(self):
        df = items({'Customer PO': 'PO-1', 'Remarks': 'H-8327'})
        assert resolve_so_remark_column(df, doc_type='ADV') == 'Remarks'

    def test_unknown_doc_type_is_none(self):
        df = items({'Customer PO': 'PO-1', 'SO Remarks': 'H-8327'})
        assert resolve_so_remark_column(df, doc_type='PMT') is None

    def test_no_items_is_none(self):
        assert resolve_so_remark_column(None) is None

    def test_dn_path_ignores_dn_own_remarks(self):
        """'SO Remarks' 없이 'Remarks'만 있으면 DN 경로에선 None — 월합 문구이기 때문"""
        df = items({'Customer PO': 'PO-1', 'Remarks': '25일 마감, 월합세금계산서'})
        assert resolve_so_remark_column(df) is None


# === 품명(B열)의 고객 품번 ===

SJV = '주식회사 에스제이브이'


class TestItemDescription:
    """품명 = '고객 품번 + 품명' (OC·CI·PL의 Model number 앞붙임과 같은 모양)"""

    def test_model_code_prefixed(self):
        """실측 SOD-2026-1139 L1"""
        item = pd.Series({'Customer name': SJV, 'Model code': '66853010035',
                          'Item name': 'HS02-S180HI-N FAIL CLOSE SBWG-01DCM'})
        assert build_item_description(item) == '66853010035 HS02-S180HI-N FAIL CLOSE SBWG-01DCM'

    @pytest.mark.parametrize('blank', [None, float('nan'), '', '   '])
    def test_blank_model_code_keeps_item_name(self, blank):
        """품번 없는 행(지금은 거의 전부)은 예전 그대로 품명만"""
        item = pd.Series({'Model code': blank, 'Item name': 'IQ3'})
        assert build_item_description(item) == 'IQ3'

    def test_no_model_code_column_keeps_item_name(self):
        item = pd.Series({'Item name': 'IQ3'})
        assert build_item_description(item) == 'IQ3'

    def test_float_model_code_has_no_trailing_zero(self):
        """dtype을 놓친 경로에서 float로 읽혀도 '66853010035.0'이 찍히지 않는다"""
        item = pd.Series({'Model code': 66853010035.0, 'Item name': 'HS01-D140-N'})
        assert build_item_description(item) == '66853010035 HS01-D140-N'

    def test_ax_project_number_never_used(self):
        """별칭('model_code')에 AX Project number가 있다 — 그게 품명 앞에 찍히면 안 된다

        ADV 경로 아이템은 SO_국내 행 전체라 AX Project number를 갖고 있다.
        """
        item = pd.Series({'AX Project number': '82-17677', 'Item name': 'IQ3'})
        assert build_item_description(item) == 'IQ3'

    def test_model_code_without_item_name(self):
        item = pd.Series({'Model code': '66853010035', 'Item name': None})
        assert build_item_description(item) == '66853010035'


# === 기존 동작 회귀 (호선명과 무관한 PO 나열) ===

class TestExistingPoBehavior:
    def test_multiple_pos_joined_in_order(self):
        result = build_po_cell_text(order(CK), items(
            {'Customer PO': '26071402R0'},
            {'Customer PO': '26071403R0'},
            {'Customer PO': '26071402R0'},   # 중복은 한 번만
        ))
        assert result == '26071402R0, 26071403R0'

    def test_blank_and_nan_pos_dropped(self):
        result = build_po_cell_text(order(CK), items(
            {'Customer PO': 'PO-1'},
            {'Customer PO': None},
            {'Customer PO': '  '},
        ))
        assert result == 'PO-1'

    def test_no_po_column_falls_back_to_order_data(self):
        result = build_po_cell_text(order(CK, po='PO-77'), items({'Item name': 'X'}))
        assert result == 'PO-77'

    def test_none_items_falls_back_to_order_data(self):
        assert build_po_cell_text(order(CK, po='PO-77'), None) == 'PO-77'


# === load_dn_data가 'SO Remarks'를 실어 오는가 (ts_generator와의 결합 지점) ===

class TestLoadDnDataSoRemarks:
    """컬럼명 'SO Remarks'는 utils(생산)와 ts_generator(소비)의 약속이다.

    한쪽만 바뀌면 호선명이 조용히 사라진다 — 병합 결과로 계약을 못박는다.
    """

    @pytest.fixture
    def fake_workbook(self, tmp_path, monkeypatch):
        # Model code는 시트 실물처럼 **숫자 셀**로 쓴다 (빈칸이 섞여 dtype 없이 읽으면 float)
        so = pd.DataFrame([
            {'SO_ID': 'SOD-1', 'Line item': 1, 'Customer name': SCANTECH,
             'Customer PO': 'OR26060022', 'Model code': 66853010035,
             'Item name': 'IQ3', 'Item qty': 2,
             'Sales Unit Price': 1000, 'Business registration number': '123-45-67890',
             'Remarks': 'H-8327', 'Comment': 'LR 선급 진행 건'},
            {'SO_ID': 'SOD-1', 'Line item': 2, 'Customer name': SCANTECH,
             'Customer PO': 'OR26060022', 'Model code': None,
             'Item name': 'IQ5', 'Item qty': 1,
             'Sales Unit Price': 2000, 'Business registration number': '123-45-67890',
             'Remarks': None, 'Comment': 'Split from line 1'},
        ])
        dn = pd.DataFrame([
            {'DN_ID': 'DND-1', 'SO_ID': 'SOD-1', 'Line item': 1,
             'Remarks': '25일 마감, 월합세금계산서'},
            {'DN_ID': 'DND-1', 'SO_ID': 'SOD-1', 'Line item': 2,
             'Remarks': '25일 마감, 월합세금계산서'},
        ])
        pmt = pd.DataFrame([{'선수금_ID': 'ADV-1', 'SO_ID': 'SOD-1'}])
        path = tmp_path / 'NOAH_SO_PO_DN.xlsx'
        with pd.ExcelWriter(path) as writer:
            so.to_excel(writer, sheet_name=SO_DOMESTIC_SHEET, index=False)
            dn.to_excel(writer, sheet_name=DN_DOMESTIC_SHEET, index=False)
            pmt.to_excel(writer, sheet_name=PMT_DOMESTIC_SHEET, index=False)
        monkeypatch.setattr(po_utils, 'NOAH_SO_PO_DN_FILE', path)
        return path

    def test_so_remarks_column_present_with_so_values(self, fake_workbook):
        df = po_utils.load_dn_data()
        assert 'SO Remarks' in df.columns
        by_line = df.set_index('Line item')['SO Remarks']
        assert by_line[1] == 'H-8327'
        assert pd.isna(by_line[2])

    def test_dn_own_remarks_untouched(self, fake_workbook):
        """DN 자체 Remarks(월합 문구)는 이름 그대로, 값 그대로 남는다"""
        df = po_utils.load_dn_data()
        assert set(df['Remarks']) == {'25일 마감, 월합세금계산서'}

    def test_so_comment_never_loaded(self, fake_workbook):
        """SO Comment는 내부 메모 칸이다 — 거래명세표 데이터에 아예 실리지 않는다

        Remarks를 전 거래처에 표기하게 된 전제(2026-11-03)가 이것이다. 병합 컬럼을
        '전부'로 바꾸면 Comment가 어느 이름으로든 따라온다 — 값으로 못박는다.
        """
        df = po_utils.load_dn_data()
        assert not any('Comment' in str(c) for c in df.columns)
        values = set(df.astype(str).to_numpy().ravel())
        assert 'LR 선급 진행 건' not in values
        assert 'Split from line 1' not in values

    def test_model_code_loaded_as_text(self, fake_workbook):
        """고객 품번은 DN 경로에 문자열로 실린다 — float로 오면 '66853010035.0'이 찍힐 수 있다"""
        df = po_utils.load_dn_data()
        by_line = df.set_index('Line item')['Model code']
        assert by_line[1] == '66853010035'
        assert pd.isna(by_line[2])

    def test_dn_path_description_end_to_end(self, fake_workbook):
        """로더 → 품명: 품번 있는 라인만 앞에 붙는다"""
        df = po_utils.load_dn_data().set_index('Line item')
        assert build_item_description(df.loc[1]) == '66853010035 IQ3'
        assert build_item_description(df.loc[2]) == 'IQ5'

    def test_advance_path_model_code_loaded_as_text(self, fake_workbook):
        """선수금(ADV) 거래명세표도 같은 품번이 같은 모양으로 실린다"""
        _, so_items = po_utils.load_so_for_advance('ADV-1')
        so_items = so_items.set_index('Line item')
        assert so_items.loc[1, 'Model code'] == '66853010035'
        assert build_item_description(so_items.loc[1]) == '66853010035 IQ3'
        assert build_item_description(so_items.loc[2]) == 'IQ5'
