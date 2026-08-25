"""
거래명세표 PO No. 칸 — 호선명(SO Remarks) 병기 테스트
=====================================================

조선 기자재 4사(스칸텍·브이티엘·엔이에스·파나시아)는 발주가 호선 단위라
하단 PO No. 옆에 `SO_국내.Remarks`(호선명)를 병기한다 (2026-08-25 사용자 지시).

지켜야 할 것 둘:
- **지정 거래처가 아니면 절대 붙지 않는다** — 다른 거래처의 Remarks에는 내부 메모가 있다.
- **DN 자체 Remarks(월합 세금계산서 문구)는 절대 PO 칸에 새지 않는다** — 이름이 같아서
  헷갈리기 쉬운 자리다 (`load_dn_data()`가 SO 쪽을 'SO Remarks'로 개명하는 이유).

`build_po_cell_text()`는 COM 없는 순수 함수라 Excel 없이 검증한다.
"""

import pandas as pd
import pytest

from po_generator.config import DN_DOMESTIC_SHEET, SO_DOMESTIC_SHEET
from po_generator.ts_generator import build_po_cell_text
from po_generator import utils as po_utils


SCANTECH = '주식회사 스칸텍'          # 시트 실표기 — 법인 접두가 붙는다
NES = '엔이에스 주식회사'             # 접미가 붙는 경우
CK = '씨앤케이엔지니어링(C&K ENG)'    # 비대상 거래처


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
        SCANTECH, '주식회사 브이티엘', NES, '주식회사 파나시아',
    ])
    def test_all_four_customers_match(self, customer):
        result = build_po_cell_text(order(customer), items(
            {'Customer PO': 'PO-1', 'SO Remarks': 'S590'},
        ))
        assert result == 'PO-1 (S590)'


# === 붙으면 안 되는 자리 ===

class TestVesselNameNotAppended:
    def test_other_customers_never_show_remarks(self):
        """지정 4사가 아니면 Remarks가 있어도 붙지 않는다 (내부 메모 노출 방지)"""
        result = build_po_cell_text(order(CK, '26071402R0'), items(
            {'Customer PO': '26071402R0', 'SO Remarks': '내부 메모'},
        ))
        assert result == '26071402R0'

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
        so = pd.DataFrame([
            {'SO_ID': 'SOD-1', 'Line item': 1, 'Customer name': SCANTECH,
             'Customer PO': 'OR26060022', 'Item name': 'IQ3', 'Item qty': 2,
             'Sales Unit Price': 1000, 'Business registration number': '123-45-67890',
             'Remarks': 'H-8327'},
            {'SO_ID': 'SOD-1', 'Line item': 2, 'Customer name': SCANTECH,
             'Customer PO': 'OR26060022', 'Item name': 'IQ5', 'Item qty': 1,
             'Sales Unit Price': 2000, 'Business registration number': '123-45-67890',
             'Remarks': None},
        ])
        dn = pd.DataFrame([
            {'DN_ID': 'DND-1', 'SO_ID': 'SOD-1', 'Line item': 1,
             'Remarks': '25일 마감, 월합세금계산서'},
            {'DN_ID': 'DND-1', 'SO_ID': 'SOD-1', 'Line item': 2,
             'Remarks': '25일 마감, 월합세금계산서'},
        ])
        path = tmp_path / 'NOAH_SO_PO_DN.xlsx'
        with pd.ExcelWriter(path) as writer:
            so.to_excel(writer, sheet_name=SO_DOMESTIC_SHEET, index=False)
            dn.to_excel(writer, sheet_name=DN_DOMESTIC_SHEET, index=False)
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
