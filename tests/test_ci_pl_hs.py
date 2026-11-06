"""CI·PL 서비스의 HS 규칙 — 어느 코드로 몇 장을 내나 (Excel COM 없이)

- 일반 거래처 = 수출신고 기본값 한 장
- SECTORIEL = **고객 코드 한 장만**, 파일명도 `CI_*`/`PL_*` 그대로 (2026-11-06 사용자 지시 —
  예전엔 기본판 + `CI-HS_*`/`PL-HS_*` 두 장이었다)
- SECTORIEL인데 `OS name`이 없으면(SO 조인 실패) 기본값으로 메우지 않고 만들지 않는다

조회는 진짜 `FinderService`, 데이터 로드·생성기만 대역 (test_coo의 서비스 절과 같은 방식)
"""

from __future__ import annotations

import re

import pandas as pd
import pytest

from po_generator.config import DEFAULT_HS_CODE
from po_generator.hs_code import HS_COLUMN, HS_ELECTRIC_ACTUATOR, OS_NAME_COLUMN
from po_generator.services import document_service
from po_generator.services.document_service import DocumentService
from po_generator.services.finder_service import FinderService

SECTORIEL = 'SECTORIEL SA'
DN_ID = 'DNO-2026-0190'


def dn_row(line=1, name='NA28 280Nm 24s. 230V50HZ', customer='WATERGATES GMBH') -> dict:
    """`load_dn_export_data()` 한 행"""
    return {
        'DN_ID': DN_ID, 'SO_ID': 'SOO-2026-0298', 'Customer name': customer,
        'Customer PO': 'PO-1', 'Item': name, 'Line item': line, 'Qty': 2,
        'Shipping Method': 'OCEAN', 'Bill to 1': customer, '출고일': pd.Timestamp('2026-10-02'),
    }


def so_rows(os_name: str | None = 'Noah NA') -> pd.DataFrame:
    df = pd.DataFrame(
        [{'SO_ID': 'SOO-2026-0298', 'Line item': n, 'Model number': f'M{n}'} for n in (1, 2)]
    )
    if os_name is not None:
        df[OS_NAME_COLUMN] = os_name
    return df


# (문서, 서비스 메서드, 생성기 이름, 템플릿 상수, 출력 폴더 상수)
DOCS = [
    ('CI', 'generate_ci', 'create_ci_xlwings', 'CI_TEMPLATE_FILE', 'CI_OUTPUT_DIR'),
    ('PL', 'generate_pl', 'create_pl_xlwings', 'PL_TEMPLATE_FILE', 'PL_OUTPUT_DIR'),
]


@pytest.fixture(params=DOCS, ids=[d[0] for d in DOCS])
def run(request, tmp_path, monkeypatch):
    """`run(rows, so=None) -> (result, 생성기 호출 목록, 출력 폴더)`"""
    prefix, method, creator, template_attr, out_attr = request.param

    template = tmp_path / f'{prefix}.xlsx'
    template.write_bytes(b'template')
    out_dir = tmp_path / f'generated_{prefix.lower()}'
    monkeypatch.setattr(document_service, template_attr, template)
    monkeypatch.setattr(document_service, out_attr, out_dir)
    # PL의 중량 보강은 PO_해외를 읽는다 — HS 규칙과 무관하니 비운다
    monkeypatch.setattr(document_service, 'build_po_line_weight_map', lambda: {})

    calls: list[dict] = []

    def fake_create(template_path, output_path, order_data, items_df):
        calls.append({'output_path': output_path, 'items_df': items_df})
        output_path.write_bytes(b'dummy')

    monkeypatch.setattr(document_service, creator, fake_create)

    def _run(rows: list[dict], so: pd.DataFrame | None = None):
        finder = FinderService()
        monkeypatch.setattr(finder, 'load_dn_export_data', lambda: pd.DataFrame(rows))
        so_df = so if so is not None else so_rows()
        monkeypatch.setattr(finder, 'load_so_export_data', lambda: so_df)
        result = getattr(DocumentService(finder=finder), method)(DN_ID)
        return prefix, result, calls, out_dir

    return _run


class TestHsRule:
    def test_일반_거래처는_수출신고_코드_한_장(self, run):
        prefix, result, calls, _ = run([dn_row(), dn_row(line=2, name='FEDEX COST')])
        assert result.success and not result.errors
        assert len(calls) == 1
        assert list(calls[0]['items_df'][HS_COLUMN]) == [DEFAULT_HS_CODE, '']
        assert result.output_file.name.startswith(f'{prefix}_')

    def test_SECTORIEL은_고객_코드_한_장만(self, run):
        """기본판을 따로 만들지 않는다 — 생성기 호출 1회, 부가 파일 없음"""
        rows = [dn_row(customer=SECTORIEL), dn_row(line=2, name='FEDEX COST', customer=SECTORIEL)]
        prefix, result, calls, out_dir = run(rows)
        assert result.success and not result.errors
        assert len(calls) == 1
        assert list(calls[0]['items_df'][HS_COLUMN]) == [HS_ELECTRIC_ACTUATOR, '']
        assert result.extra_files == []
        assert [p.name for p in out_dir.iterdir()] == [result.output_file.name]

    def test_SECTORIEL_파일명에_HS가_붙지_않는다(self, run):
        prefix, result, _, _ = run([dn_row(customer=SECTORIEL)])
        assert re.fullmatch(
            rf'{prefix}_{DN_ID}_SECTORIEL_SA_\d{{6}}\.xlsx', result.output_file.name
        )

    def test_SECTORIEL_미매칭은_경고로_알린다(self, run):
        """모르는 품목은 추측하지 않고 비운다 — 문서는 나가되 경고가 붙는다"""
        so = so_rows(os_name='Electric Spares')
        _, result, calls, _ = run([dn_row(name='CAPOT TIGE', customer=SECTORIEL)], so=so)
        assert result.success
        assert list(calls[0]['items_df'][HS_COLUMN]) == ['']
        assert result.warnings and 'HS 코드 미지정' in result.warnings[0]

    def test_SECTORIEL인데_OS_name이_없으면_만들지_않는다(self, run):
        """수출신고 기본값으로 메우지 않는다 — 고객이 받기로 한 코드가 아니다"""
        _, result, calls, _ = run([dn_row(customer=SECTORIEL)], so=so_rows(os_name=None))
        assert not result.success
        assert OS_NAME_COLUMN in result.errors[0]
        assert calls == []
