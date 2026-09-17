"""Final Invoice 파일명에 NO-XXXX(RCK PO) 표기 (2026-09-15)

`FI_{DN_ID}_{NO-XXXX}_{고객명}_{날짜}.xlsx` — NO-XXXX는 `DN_해외.RCK PO`
(RCK → NOAH 발주번호)다.

조회는 진짜 `FinderService`를 쓰고 데이터 로드만 대역으로 바꾼다. 단일 아이템이면
`items_df`가 None인 실데이터 모양을 그대로 지나가게 하려고 — fixture가 모양을
흉내내면 그 경로는 검증된 적이 없다 (2026-08-11 교훈). 생성기는 빈 파일만 쓴다
(Excel COM 없이).
"""

from __future__ import annotations

import re

import pandas as pd
import pytest

from po_generator.services import document_service
from po_generator.services.document_service import DocumentService
from po_generator.services.finder_service import FinderService

WG = 'WATERGATES GMBH'


def dn_row(dn_id: str, rck_po, customer_po, line: int = 1, customer: str = WG) -> dict:
    """`load_dn_export_data()` 한 행 (파일명에 쓰는 컬럼 위주)"""
    return {
        'DN_ID': dn_id,
        'SO_ID': 'SOO-2026-0001',
        'RCK PO': rck_po,
        'Customer name': customer,
        'Customer PO': customer_po,
        'Item': f'ITEM-{line}',
        'Line item': line,
        'Qty': 1,
    }


@pytest.fixture
def make_service(tmp_path, monkeypatch):
    """행 목록으로 DocumentService를 만든다 — 출력은 tmp_path, 생성은 빈 파일"""
    template = tmp_path / 'final_invoice.xlsx'
    template.write_bytes(b'template')
    monkeypatch.setattr(document_service, 'FI_TEMPLATE_FILE', template)
    monkeypatch.setattr(document_service, 'FI_OUTPUT_DIR', tmp_path / 'generated_fi')
    monkeypatch.setattr(
        document_service, 'create_fi_xlwings',
        lambda template_path, output_path, order_data, items_df: output_path.write_bytes(b'dummy'),
    )

    def _make(rows: list[dict]) -> DocumentService:
        finder = FinderService()
        df = pd.DataFrame(rows)
        monkeypatch.setattr(finder, 'load_dn_export_data', lambda: df)
        return DocumentService(finder=finder)

    return _make


def name_of(result) -> str:
    """실행일(yymmdd)을 뗀 파일명 — 충돌 접미사(`_1`)가 붙으면 여기서 걸린다"""
    assert result.success, result.message
    m = re.fullmatch(r'(.+)_\d{6}\.xlsx', result.output_file.name)
    assert m, f'예상 밖 파일명: {result.output_file.name}'
    return m.group(1)


class TestDnMode:
    """`create_fi.py DNO-…` — DN 단위 (복수 RCK PO면 발주번호별 분리)"""

    def test_single_item(self, make_service):
        """단일 아이템 DN — items_df가 None인 경로도 NO를 찾는다"""
        svc = make_service([dn_row('DNO-2026-0177', 'NO-0243', 'PO-007746',
                                   customer='SULLIVAN PROCESS CONTROLS LLC')])
        assert name_of(svc.generate_fi('DNO-2026-0177')) == \
            'FI_DNO-2026-0177_NO-0243_SULLIVAN_PROCESS_CONTROLS_LLC'

    def test_multi_item_one_rck_po(self, make_service):
        svc = make_service([dn_row('DNO-2026-0175', 'NO-0215', 'WGDBS2700486', line=n)
                            for n in (1, 2, 3)])
        assert name_of(svc.generate_fi('DNO-2026-0175')) == \
            'FI_DNO-2026-0175_NO-0215_WATERGATES_GMBH'

    def test_blank_rck_po_lines_do_not_leak_nan(self, make_service):
        """공란 라인이 섞여도 'nan'이 파일명에 새지 않는다"""
        svc = make_service([
            dn_row('DNO-2026-0175', 'NO-0215', 'WGDBS2700486', line=1),
            dn_row('DNO-2026-0175', None, 'WGDBS2700486', line=2),
            dn_row('DNO-2026-0175', '', 'WGDBS2700486', line=3),
        ])
        assert name_of(svc.generate_fi('DNO-2026-0175')) == \
            'FI_DNO-2026-0175_NO-0215_WATERGATES_GMBH'

    def test_split_by_rck_po(self, make_service):
        """발주번호별 분리 — Customer PO가 같은 발주번호끼리도 이름이 갈린다

        실측 DNO-2026-0020: NO-0009·0013·0027·0033이 모두 WGDBS2601233이라
        예전엔 `_1`/`_2` 접미사로만 구분됐다.
        """
        svc = make_service([
            dn_row('DNO-2026-0020', 'NO-0009', 'WGDBS2601233'),
            dn_row('DNO-2026-0020', 'NO-0013', 'WGDBS2601233'),
            dn_row('DNO-2026-0020', 'NO-0018', 'WGDBS2601282'),
        ])
        names = [name_of(svc.generate_fi('DNO-2026-0020', rck_po=no))
                 for no in ('NO-0009', 'NO-0013', 'NO-0018')]
        assert names == [
            'FI_DNO-2026-0020_NO-0009_WGDBS2601233_WATERGATES_GMBH',
            'FI_DNO-2026-0020_NO-0013_WGDBS2601233_WATERGATES_GMBH',
            'FI_DNO-2026-0020_NO-0018_WGDBS2601282_WATERGATES_GMBH',
        ]

    def test_split_without_customer_po_does_not_repeat_no(self, make_service):
        """Customer PO가 비면 예전엔 RCK PO로 폴백했다 — 이제 NO가 늘 붙으므로 두 번 쓰지 않는다"""
        svc = make_service([
            dn_row('DNO-2026-0100', 'NO-0100', None),
            dn_row('DNO-2026-0100', 'NO-0101', 'CPO-1'),
        ])
        assert name_of(svc.generate_fi('DNO-2026-0100', rck_po='NO-0100')) == \
            'FI_DNO-2026-0100_NO-0100_WATERGATES_GMBH'

    def test_no_rck_po_column_keeps_old_name(self, make_service):
        """RCK PO 컬럼이 없으면 붙일 NO가 없다 — 예전 이름 그대로"""
        row = dn_row('DNO-2026-0177', 'NO-0243', 'PO-007746')
        del row['RCK PO']
        svc = make_service([row])
        assert name_of(svc.generate_fi('DNO-2026-0177')) == 'FI_DNO-2026-0177_WATERGATES_GMBH'


class TestCustomerPoMode:
    """`create_fi.py --po …` — 발주번호 기준 통합 (복수 DN)"""

    def test_single_no(self, make_service):
        mrc = 'MRC Global(New Zealand) Ltd.'
        svc = make_service([
            dn_row('DNO-2026-0176', 'NO-0262', '4000469308', customer=mrc),
            dn_row('DNO-2026-0180', 'NO-0262', '4000469308', line=2, customer=mrc),
        ])
        assert name_of(svc.generate_fi_by_customer_po('4000469308')) == \
            'FI_4000469308_NO-0262_MRC_Global(New_Zealand)_Ltd.'

    def test_multiple_nos_all_listed_sorted(self, make_service):
        """발주번호가 여럿이면 전부 싣는다 — 실측 WGDBS2601233은 NO 5개"""
        svc = make_service([
            dn_row('DNO-2026-0020', 'NO-0033', 'WGDBS2601233', line=1),
            dn_row('DNO-2026-0020', 'NO-0009', 'WGDBS2601233', line=2),
            dn_row('DNO-2026-0021', 'NO-0013', 'WGDBS2601233', line=1),
            dn_row('DNO-2026-0021', 'NO-0009', 'WGDBS2601233', line=3),
        ])
        assert name_of(svc.generate_fi_by_customer_po('WGDBS2601233')) == \
            'FI_WGDBS2601233_NO-0009+NO-0013+NO-0033_WATERGATES_GMBH'
