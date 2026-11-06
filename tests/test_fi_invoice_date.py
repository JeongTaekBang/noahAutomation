"""Final Invoice — Invoice Date = 선적일, xlsx + PDF (2026-11-06 사용자 요청)

예전엔 `dispatch_date` 별칭이 `출고일`을 먼저 찾아 **공장 출고일**이 찍혔다.
DN_해외는 `출고일`과 `선적일`이 따로 있고 실측 중앙값 16일 차이라, 이 테스트의 핵심은
"두 날짜가 다른 행에서 선적일이 나온다"이다.

판단(`resolve_invoice_date`)은 COM 없는 순수 함수로 직접 보고, 서비스는 진짜
`FinderService`에 데이터 로드만 대역으로 바꿔 생성기에 넘어간 날짜와 PDF를 본다
(`tests/test_fi_filename.py`와 같은 방식).
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from po_generator.fi_generator import resolve_invoice_date
from po_generator.services import document_service
from po_generator.services.document_service import DocumentService
from po_generator.services.finder_service import FinderService

TODAY = date(2026, 11, 6)


def lines(*ship_dates, dispatch='2026-10-20') -> pd.DataFrame:
    """DN_해외 라인 — 출고일은 늘 선적일과 다르게 둔다 (어느 쪽이 찍혔는지 갈리게)"""
    return pd.DataFrame({
        '출고일': [pd.Timestamp(dispatch)] * len(ship_dates),
        '선적일': [pd.Timestamp(d) if d else pd.NaT for d in ship_dates],
    })


class TestResolveInvoiceDate:
    def test_출고일이_아니라_선적일(self):
        assert resolve_invoice_date(lines('2026-11-05', '2026-11-05'), TODAY) == ('2026-11-05', [])

    def test_선적일이_여럿이면_가장_늦은_날_경고(self):
        """실측 DNO-2026-0015 — 한 DN을 2/20·3/02 두 번에 실었다"""
        day, warnings = resolve_invoice_date(lines('2026-03-02', '2026-02-20'), TODAY)
        assert day == '2026-03-02'
        assert len(warnings) == 1 and '2026-02-20, 2026-03-02' in warnings[0]

    def test_일부만_공란이면_날짜는_있는_것으로_경고(self):
        day, warnings = resolve_invoice_date(lines('2026-11-05', None), TODAY)
        assert day == '2026-11-05'
        assert len(warnings) == 1 and '1건' in warnings[0]

    def test_전부_공란이면_오늘_경고(self):
        """아직 선적 전 (실측 DNO-2026-0209) — 출고일로 메우지 않는다"""
        day, warnings = resolve_invoice_date(lines(None, None), TODAY)
        assert day == '2026-11-06'
        assert len(warnings) == 1 and '비어' in warnings[0]

    def test_선적일_컬럼이_없으면_오늘_경고(self):
        """dispatch_date 별칭처럼 출고일로 조용히 풀리지 않는다"""
        df = lines('2026-11-05').drop(columns=['선적일'])
        day, warnings = resolve_invoice_date(df, TODAY)
        assert day == '2026-11-06'
        assert len(warnings) == 1 and '컬럼' in warnings[0]

    def test_텍스트_날짜도_읽는다(self):
        df = pd.DataFrame({'선적일': ['2026-11-05']})
        assert resolve_invoice_date(df, TODAY) == ('2026-11-05', [])


def dn_row(dn_id: str, ship, line: int = 1, customer_po: str = 'PO-1') -> dict:
    return {
        'DN_ID': dn_id,
        'SO_ID': 'SOO-2026-0001',
        'RCK PO': 'NO-0001',
        'Customer name': 'WATERGATES GMBH',
        'Customer PO': customer_po,
        'Item': f'ITEM-{line}',
        'Line item': line,
        'Qty': 1,
        '출고일': pd.Timestamp('2026-10-20'),
        '선적일': pd.Timestamp(ship) if ship else pd.NaT,
    }


@pytest.fixture
def make_service(tmp_path, monkeypatch):
    """행 목록으로 DocumentService — 생성기는 받은 날짜를 기록하고 빈 파일만 쓴다"""
    template = tmp_path / 'final_invoice.xlsx'
    template.write_bytes(b'template')
    monkeypatch.setattr(document_service, 'FI_TEMPLATE_FILE', template)
    monkeypatch.setattr(document_service, 'FI_OUTPUT_DIR', tmp_path / 'generated_fi')

    captured: dict = {}

    def fake_create(template_path, output_path, order_data, items_df, invoice_date):
        captured['invoice_date'] = invoice_date
        output_path.write_bytes(b'dummy')

    def fake_pdf(xlsx_path):
        pdf = Path(xlsx_path).with_suffix('.pdf')
        pdf.write_bytes(b'%PDF')
        return pdf

    monkeypatch.setattr(document_service, 'create_fi_xlwings', fake_create)
    monkeypatch.setattr(document_service, 'export_pdf', fake_pdf)

    def _make(rows: list[dict]) -> DocumentService:
        finder = FinderService()
        monkeypatch.setattr(finder, 'load_dn_export_data', lambda: pd.DataFrame(rows))
        return DocumentService(finder=finder)

    return _make, captured, monkeypatch


class TestService:
    def test_DN_기준_선적일과_PDF(self, make_service):
        make, captured, _ = make_service
        result = make([dn_row('DNO-2026-0208', '2026-11-05')]).generate_fi('DNO-2026-0208')
        assert result.success and not result.errors and not result.warnings
        assert captured['invoice_date'] == '2026-11-05'
        assert [p.suffix for p in result.extra_files] == ['.pdf']
        assert result.extra_files[0].stem == result.output_file.stem
        assert result.extra_files[0].exists()

    def test_발주번호_기준도_같은_규칙(self, make_service):
        """복수 DN 통합 — 늦게 실은 DN의 선적일, 날짜가 갈린다는 경고"""
        make, captured, _ = make_service
        svc = make([
            dn_row('DNO-2026-0176', '2026-10-30', customer_po='4000469308'),
            dn_row('DNO-2026-0180', '2026-11-05', line=2, customer_po='4000469308'),
        ])
        result = svc.generate_fi_by_customer_po('4000469308')
        assert result.success and not result.errors
        assert captured['invoice_date'] == '2026-11-05'
        assert len(result.warnings) == 1
        assert [p.suffix for p in result.extra_files] == ['.pdf']

    def test_선적_전이면_경고가_결과에_실린다(self, make_service):
        make, _, _ = make_service
        result = make([dn_row('DNO-2026-0209', None)]).generate_fi('DNO-2026-0209')
        assert result.success
        assert len(result.warnings) == 1 and '선적일' in result.warnings[0]

    def test_PDF가_실패해도_xlsx는_남고_오류로_알린다(self, make_service):
        make, _, monkeypatch = make_service

        def broken_pdf(xlsx_path):
            raise RuntimeError('Excel 없음')

        monkeypatch.setattr(document_service, 'export_pdf', broken_pdf)
        result = make([dn_row('DNO-2026-0208', '2026-11-05')]).generate_fi('DNO-2026-0208')
        assert result.success and result.output_file.exists()
        assert result.extra_files == []
        assert len(result.errors) == 1 and 'PDF' in result.errors[0]
