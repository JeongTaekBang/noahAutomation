"""Final Invoice — Invoice No = INV 번호, DN 번호는 참조로 (2026-11-11 사용자 요청)

DN 하나에 PO가 여럿이면 FI를 발주번호별로 나눠 내는데 그 전부에 같은 DN 번호가 찍혔다
(실측 DN 215건 중 39건, `DNO-2026-0020`은 9장). 이제 `DN_해외`의 INV 번호가 인보이스
단위다 — 같은 번호를 넣은 행 전부가 한 장.

이 테스트가 지키는 것:
  - **어느 입구(INV·DN·발주번호)로 들어가도 같은 번호는 같은 문서** — 다른 DN에 있는
    행까지 실린다
  - INV 번호가 없는 지난 건은 예전 그대로 (다시 뽑아도 번호가 안 바뀐다)
  - 일부만 채운 DN, 거래처·통화가 섞인 번호는 만들지 않는다
  - INV 번호가 있는 라인이 DN 번호로도 나가지 않는다 (같은 청구가 번호 둘로)

판정(`fi_invoice`)은 COM 없는 순수 함수로 직접 보고, 서비스·CLI는 진짜 `FinderService`에
데이터 로드만 대역으로 바꾼다 (`tests/test_fi_filename.py`와 같은 방식).
"""

from __future__ import annotations

import inspect
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

import create_fi
from po_generator import fi_generator
from po_generator.config import COLUMN_ALIASES, FI_TEMPLATE_FILE
from po_generator.fi_invoice import (
    INVOICE_COLUMN_NAME,
    build_fi_references,
    check_invoice_lines,
    invoice_numbers,
    invoices_of,
    suggest_invoice_numbers,
)
from po_generator.services import GenerationStatus, document_service
from po_generator.services.document_service import DocumentService
from po_generator.services.finder_service import FinderService
from po_generator.utils import resolve_column

WG = 'WATERGATES GMBH'
INV = INVOICE_COLUMN_NAME


def row(dn_id: str, rck_po: str, customer_po, inv=None, *, line: int = 1,
        customer: str = WG, code: str = 'C-0054', currency: str = 'EUR',
        po_date: str | None = '2026-01-14') -> dict:
    """`load_dn_export_data()` 한 행 — INV 컬럼이 있는 시트"""
    return {
        'DN_ID': dn_id,
        'SO_ID': 'SOO-2026-0001',
        'RCK PO': rck_po,
        'Customer name': customer,
        'Business registration number': code,   # SO_해외에서 온 고객코드
        'Customer PO': customer_po,
        'PO receipt date': pd.Timestamp(po_date) if po_date else pd.NaT,
        'Currency': currency,
        'Item': f'ITEM-{line}',
        'Line item': line,
        'Qty': 1,
        '선적일': pd.Timestamp('2026-11-05'),
        INV: inv,
    }


def frame(*rows: dict) -> pd.DataFrame:
    return pd.DataFrame(list(rows))


# === 판정 (순수 함수) =========================================================

class TestInvoiceNumbers:
    def test_공백을_떼고_빈칸은_빈_문자열(self):
        df = frame(row('D1', 'NO-1', 'P1', ' INV-2026-0001 '), row('D1', 'NO-2', 'P2', None),
                   row('D1', 'NO-3', 'P3', ''))
        assert invoice_numbers(df).tolist() == ['INV-2026-0001', '', '']

    def test_컬럼이_없으면_전부_빈_문자열(self):
        df = frame(row('D1', 'NO-1', 'P1')).drop(columns=[INV])
        assert invoice_numbers(df).tolist() == ['']

    def test_같은_시트의_다른_Inv_컬럼으로_풀리지_않는다(self):
        """`RCK Inv no.`(1~3월의 다른 체계)·`NOAH Inv no.`(공장 인보이스)가 별칭에 들어가면
        지난 건을 다시 뽑을 때 그 값이 Invoice No로 찍힌다"""
        columns = ['DN_ID', 'NOAH Inv no.', 'RCK Inv no.']
        assert resolve_column(columns, 'invoice_no') is None
        assert not {'NOAH Inv no.', 'RCK Inv no.'} & set(COLUMN_ALIASES['invoice_no'])


class TestInvoicesOf:
    def test_전부_공란이면_예전_방식(self):
        df = frame(row('D1', 'NO-1', 'P1'), row('D1', 'NO-2', 'P2'))
        assert invoices_of(df, 'D1') == ([], None)

    def test_컬럼이_없어도_예전_방식(self):
        df = frame(row('D1', 'NO-1', 'P1')).drop(columns=[INV])
        assert invoices_of(df, 'D1') == ([], None)

    def test_전부_채웠으면_번호별(self):
        df = frame(row('D1', 'NO-2', 'P2', 'INV-2026-0002'),
                   row('D1', 'NO-1', 'P1', 'INV-2026-0001'),
                   row('D1', 'NO-1', 'P1', 'INV-2026-0001', line=2))
        assert invoices_of(df, 'D1') == (['INV-2026-0001', 'INV-2026-0002'], None)

    def test_일부만_채웠으면_거부하고_빈_라인을_알려준다(self):
        """빈 라인은 어느 인보이스에도 안 실린다 — 과소청구"""
        df = frame(row('D1', 'NO-1', 'P1', 'INV-2026-0001'),
                   row('D1', 'NO-2', 'P2'), row('D1', 'NO-3', 'P3'))
        numbers, error = invoices_of(df, 'D1')
        assert numbers == []
        assert 'D1' in error and '2건' in error and 'NO-2, NO-3' in error


class TestCheckInvoiceLines:
    def test_한_거래처_한_통화면_통과(self):
        df = frame(row('D1', 'NO-1', 'P1', 'I1'), row('D2', 'NO-2', 'P2', 'I1'))
        assert check_invoice_lines(df) == []

    def test_거래처가_둘이면_거부(self):
        """같은 번호를 남의 행에 넣었다 — 그대로 만들면 남의 품목·단가가 찍힌다"""
        df = frame(row('D1', 'NO-1', 'P1', 'I1'),
                   row('D2', 'NO-2', 'P2', 'I1', customer='SECTORIEL', code='C-0012'))
        errors = check_invoice_lines(df)
        assert len(errors) == 1 and WG in errors[0] and 'SECTORIEL' in errors[0]

    def test_이름_표기만_다른_같은_거래처는_하나로_센다(self):
        df = frame(row('D1', 'NO-1', 'P1', 'I1'),
                   row('D2', 'NO-2', 'P2', 'I1', customer='Watergates GmbH'))
        assert check_invoice_lines(df) == []

    def test_고객코드가_없으면_이름으로_본다(self):
        df = frame(row('D1', 'NO-1', 'P1', 'I1'),
                   row('D2', 'NO-2', 'P2', 'I1', customer='SECTORIEL'),
                   ).drop(columns=['Business registration number'])
        assert len(check_invoice_lines(df)) == 1

    def test_통화가_둘이면_거부(self):
        df = frame(row('D1', 'NO-1', 'P1', 'I1'),
                   row('D1', 'NO-2', 'P2', 'I1', currency='USD'))
        errors = check_invoice_lines(df)
        assert len(errors) == 1 and 'EUR, USD' in errors[0]


class TestBuildFiReferences:
    def test_INV_번호와_DN_참조(self):
        df = frame(row('DNO-2026-0020', 'NO-0009', 'WGDBS2601233', 'INV-2026-0001'))
        refs, warnings = build_fi_references(df, 'INV-2026-0001')
        assert (refs.invoice_no, refs.customer_po, refs.dn_ids) == \
            ('INV-2026-0001', 'WGDBS2601233', 'DNO-2026-0020')
        assert refs.customer_po_date == '2026-01-14'
        assert warnings == []

    def test_INV가_없으면_예전처럼_첫_DN이_Invoice_No(self):
        df = frame(row('DNO-2026-0176', 'NO-0262', '4000469308'),
                   row('DNO-2026-0180', 'NO-0262', '4000469308', line=2))
        refs, warnings = build_fi_references(df)
        assert refs.invoice_no == 'DNO-2026-0176'
        assert refs.dn_ids == 'DNO-2026-0176, DNO-2026-0180'   # 참조에는 전부 싣는다
        assert warnings == []

    def test_여러_DN에_걸친_INV는_경고(self):
        """의도한 통합일 수도, 같은 번호를 두 선적에 잘못 넣은 것일 수도 있다"""
        df = frame(row('DNO-2026-0176', 'NO-0262', '4000469308', 'I1'),
                   row('DNO-2026-0180', 'NO-0262', '4000469308', 'I1', line=2))
        refs, warnings = build_fi_references(df, 'I1')
        assert refs.dn_ids == 'DNO-2026-0176, DNO-2026-0180'
        assert len(warnings) == 1 and 'DN 2건' in warnings[0]

    def test_Customer_PO가_여럿이면_전부_싣고_경고(self):
        df = frame(row('D1', 'NO-1', 'PO-A', 'I1'), row('D1', 'NO-2', 'PO-B', 'I1'),
                   row('D1', 'NO-3', 'PO-A', 'I1'))
        refs, warnings = build_fi_references(df, 'I1')
        assert refs.customer_po == 'PO-A, PO-B'
        assert len(warnings) == 1 and 'Customer PO가 2개' in warnings[0]

    def test_PO_일자는_Customer_PO와_같은_순서로_하나씩(self):
        """한 날짜만 찍으면 나머지 발주의 날짜가 틀린다. 빈 것은 자리를 지켜 짝이 밀리지 않게"""
        df = frame(row('D1', 'NO-1', 'PO-A', 'I1', po_date='2026-01-14'),
                   row('D1', 'NO-2', 'PO-B', 'I1', po_date=None),
                   row('D1', 'NO-3', 'PO-C', 'I1', po_date='2026-02-03'),
                   row('D1', 'NO-4', 'PO-A', 'I1', po_date='2026-01-20'))   # 같은 발주의 다른 SO
        refs, _ = build_fi_references(df, 'I1')
        assert refs.customer_po == 'PO-A, PO-B, PO-C'
        assert refs.customer_po_date == '2026-01-14, -, 2026-02-03'

    def test_PO_일자가_없으면_빈칸(self):
        refs, _ = build_fi_references(frame(row('D1', 'NO-1', 'PO-A', 'I1', po_date=None)), 'I1')
        assert refs.customer_po_date == ''
        no_col = frame(row('D1', 'NO-1', 'PO-A', 'I1')).drop(columns=['PO receipt date'])
        assert build_fi_references(no_col, 'I1')[0].customer_po_date == ''

    def test_숫자로_읽힌_발주번호에_소수점이_붙지_않는다(self):
        """SECTORIEL 발주번호(88052)는 시트에서 숫자로 읽힌다"""
        refs, _ = build_fi_references(frame(row('D1', 'NO-1', 88052.0, 'I1')), 'I1')
        assert refs.customer_po == '88052'


class TestSuggestInvoiceNumbers:
    """INV 번호는 DN 번호를 따른다 (2026-11-11 사용자 결정) — 넣을 번호를 그 규칙으로 권한다

    '지금까지 쓴 가장 큰 번호 + 1'로 권하면 어긋난다: 0191만 번호가 있을 때 DNO-2026-0215에
    0192를 권했다.
    """

    def test_PO가_하나면_DN_번호_그대로(self):
        df = frame(row('DNO-2026-0215', 'NO-0329', '91821'))
        assert suggest_invoice_numbers(df) == [('INV-2026-0215', '91821')]

    def test_PO가_여럿이면_시트_순서대로_접미사(self):
        """실측 DNO-2026-0191 — 8124708(6줄)이 -1, 8125704가 -2"""
        df = frame(row('DNO-2026-0191', 'NO-0217', '8124708', line=1),
                   row('DNO-2026-0191', 'NO-0217', '8124708', line=2),
                   row('DNO-2026-0191', 'NO-0244', '8125704'))
        assert suggest_invoice_numbers(df) == [('INV-2026-0191-1', '8124708'),
                                               ('INV-2026-0191-2', '8125704')]

    def test_한_발주를_SO_여럿으로_받았으면_번호_하나(self):
        """단위는 RCK PO가 아니라 Customer PO"""
        df = frame(row('DNO-2026-0020', 'NO-0009', 'WGDBS2601233'),
                   row('DNO-2026-0020', 'NO-0013', 'WGDBS2601233'),
                   row('DNO-2026-0020', 'NO-0018', 'WGDBS2601282'))
        assert [no for no, _ in suggest_invoice_numbers(df)] == ['INV-2026-0020-1', 'INV-2026-0020-2']

    def test_DN이_여럿이거나_모양이_다르면_권하지_않는다(self):
        two_dns = frame(row('DNO-2026-0176', 'NO-0262', 'P1'), row('DNO-2026-0180', 'NO-0262', 'P1'))
        assert suggest_invoice_numbers(two_dns) == []
        assert suggest_invoice_numbers(frame(row('D1', 'NO-1', 'P1'))) == []


# === 서비스 ===================================================================

@pytest.fixture
def make_service(tmp_path, monkeypatch):
    """행 목록으로 DocumentService — 생성기는 받은 헤더 번호를 기록하고 빈 파일만 쓴다"""
    template = tmp_path / 'final_invoice.xlsx'
    template.write_bytes(b'template')
    out_dir = tmp_path / 'generated_fi'
    monkeypatch.setattr(document_service, 'FI_TEMPLATE_FILE', template)
    monkeypatch.setattr(document_service, 'FI_OUTPUT_DIR', out_dir)

    made: list[dict] = []

    def fake_create(template_path, output_path, order_data, items_df, invoice_date, refs):
        n_items = 1 if items_df is None else len(items_df)
        made.append({'file': output_path.name, 'refs': refs, 'items': n_items})
        output_path.write_bytes(b'dummy')

    def fake_pdf(xlsx_path):
        pdf = Path(xlsx_path).with_suffix('.pdf')
        pdf.write_bytes(b'%PDF')
        return pdf

    monkeypatch.setattr(document_service, 'create_fi_xlwings', fake_create)
    monkeypatch.setattr(document_service, 'export_pdf', fake_pdf)

    def _make(*rows: dict) -> DocumentService:
        finder = FinderService()
        df = frame(*rows)
        monkeypatch.setattr(finder, 'load_dn_export_data', lambda: df)
        return DocumentService(finder=finder)

    return _make, made


# 실측 모양: DNO-2026-0020은 Customer PO 하나(WGDBS2601233)를 SO 여러 개로 받았고
# 다른 PO도 같이 실렸다. 0021에는 같은 PO의 나머지가 실렸다고 치자.
SHEET = (
    row('DNO-2026-0020', 'NO-0009', 'WGDBS2601233', 'INV-2026-0001'),
    row('DNO-2026-0020', 'NO-0013', 'WGDBS2601233', 'INV-2026-0001'),
    row('DNO-2026-0020', 'NO-0018', 'WGDBS2601282', 'INV-2026-0002'),
    row('DNO-2026-0021', 'NO-0033', 'WGDBS2601233', 'INV-2026-0001'),
    row('DNO-2026-0030', 'NO-0040', 'WGDBS2601411'),                    # INV 넣기 전의 지난 건
    row('DNO-2026-0030', 'NO-0044', 'WGDBS2601424'),
)


class TestServiceByInvoice:
    def test_같은_번호의_행_전부가_한_장_DN을_가리지_않는다(self, make_service):
        make, made = make_service
        result = make(*SHEET).generate_fi_by_invoice('INV-2026-0001')
        assert result.success and not result.errors
        assert len(made) == 1 and made[0]['items'] == 3
        refs = made[0]['refs']
        assert refs.invoice_no == 'INV-2026-0001'
        assert refs.customer_po == 'WGDBS2601233'
        assert refs.dn_ids == 'DNO-2026-0020, DNO-2026-0021'
        assert any('DN 2건' in w for w in result.warnings)
        assert [p.suffix for p in result.extra_files] == ['.pdf']

    def test_파일명은_INV_번호로_시작한다(self, make_service):
        make, made = make_service
        make(*SHEET).generate_fi_by_invoice('INV-2026-0002')
        assert made[0]['file'].startswith(
            'FI_INV-2026-0002_NO-0018_WGDBS2601282_WATERGATES_GMBH_')

    def test_입력_표기가_아니라_시트에_적힌_번호를_찍는다(self, make_service):
        make, made = make_service
        result = make(*SHEET).generate_fi_by_invoice('  inv-2026-0002 ')
        assert result.success
        assert made[0]['refs'].invoice_no == 'INV-2026-0002'

    def test_없는_번호와_빈_번호(self, make_service):
        make, made = make_service
        svc = make(*SHEET)
        assert svc.generate_fi_by_invoice('INV-2026-0999').status == GenerationStatus.NOT_FOUND
        # 빈 키가 INV가 빈 행 전부와 맞으면 지난 건이 한 장으로 뭉친다
        assert svc.generate_fi_by_invoice('  ').status == GenerationStatus.NOT_FOUND
        assert made == []

    def test_거래처가_섞인_번호는_만들지_않는다(self, make_service):
        make, made = make_service
        svc = make(row('D1', 'NO-1', 'P1', 'INV-2026-0005'),
                   row('D2', 'NO-2', 'P2', 'INV-2026-0005', customer='SECTORIEL', code='C-0012'))
        result = svc.generate_fi_by_invoice('INV-2026-0005')
        assert not result.success and result.status == GenerationStatus.VALIDATION_ERROR
        assert 'SECTORIEL' in result.errors[0]
        assert made == []


class TestServiceLegacy:
    """INV 번호가 없는 라인 — 예전 그대로 (지난 건을 다시 뽑아도 번호가 안 바뀐다)"""

    def test_DN_기준은_DN_번호가_Invoice_No(self, make_service):
        make, made = make_service
        result = make(*SHEET).generate_fi('DNO-2026-0030', rck_po='NO-0040')
        assert result.success
        refs = made[0]['refs']
        assert (refs.invoice_no, refs.dn_ids) == ('DNO-2026-0030', 'DNO-2026-0030')
        assert made[0]['file'].startswith('FI_DNO-2026-0030_NO-0040_WGDBS2601411_')

    def test_INV가_있는_라인은_DN_번호로_내지_않는다(self, make_service):
        """같은 청구가 DN 번호·INV 번호 둘로 나가면 업체는 두 건으로 본다"""
        make, made = make_service
        svc = make(*SHEET)
        for result in (svc.generate_fi('DNO-2026-0020'),
                       svc.generate_fi('DNO-2026-0020', rck_po='NO-0018'),
                       svc.generate_fi_by_customer_po('WGDBS2601233')):
            assert not result.success and result.status == GenerationStatus.VALIDATION_ERROR
            assert 'INV-2026-000' in result.errors[0]
        assert made == []


# === CLI — 어느 입구로 들어가도 같은 번호는 같은 문서 =============================

def numbers_made(made: list[dict]) -> list[str]:
    return [m['refs'].invoice_no for m in made]


class TestCli:
    def test_INV_번호를_치면_그_한_장(self, make_service, capsys):
        make, made = make_service
        ok, notices = create_fi.generate_fi(make(*SHEET), 'INV-2026-0002', set())
        assert ok and numbers_made(made) == ['INV-2026-0002']

    def test_DN을_치면_그_DN의_인보이스_전부(self, make_service, capsys):
        make, made = make_service
        ok, notices = create_fi.generate_fi(make(*SHEET), 'DNO-2026-0020', set())
        assert ok and numbers_made(made) == ['INV-2026-0001', 'INV-2026-0002']
        assert made[0]['items'] == 3          # 0021에 있는 행까지 — INV로 친 것과 같은 문서
        assert any('INV-2026-0001' in n and 'DN 2건' in n for n in notices)

    def test_발주번호를_쳐도_같은_인보이스(self, make_service, capsys):
        make, made = make_service
        ok, _ = create_fi.generate_fi_by_po(make(*SHEET), 'WGDBS2601233', set())
        assert ok and numbers_made(made) == ['INV-2026-0001'] and made[0]['items'] == 3

    def test_한_실행에서_같은_인보이스를_두_번_만들지_않는다(self, make_service, capsys):
        """DN 두 건이 한 인보이스에 걸려 있다 — 두 번 만들면 `_1` 사본이 생긴다"""
        make, made = make_service
        svc, done = make(*SHEET), set()
        assert create_fi.generate_fi(svc, 'DNO-2026-0020', done)[0]
        assert create_fi.generate_fi(svc, 'DNO-2026-0021', done)[0]
        assert numbers_made(made) == ['INV-2026-0001', 'INV-2026-0002']

    def test_일부만_채운_DN은_아무것도_만들지_않는다(self, make_service, capsys):
        make, made = make_service
        svc = make(row('DNO-2026-0040', 'NO-1', 'P1', 'INV-2026-0003'),
                   row('DNO-2026-0040', 'NO-2', 'P2'))
        ok, _ = create_fi.generate_fi(svc, 'DNO-2026-0040', set())
        assert not ok and made == []
        assert '일부 라인에만' in capsys.readouterr().out

    def test_INV가_없는_복수_PO_DN은_예전대로_나누고_같은_번호라고_알린다(self, make_service, capsys):
        make, made = make_service
        ok, notices = create_fi.generate_fi(make(*SHEET), 'DNO-2026-0030', set())
        assert ok and numbers_made(made) == ['DNO-2026-0030', 'DNO-2026-0030']
        assert len(notices) == 1 and '2장이 모두 Invoice No DNO-2026-0030' in notices[0]
        # 넣을 번호를 PO별로 — DN 번호를 따른다
        assert '권장 번호: INV-2026-0030-1 = WGDBS2601411, INV-2026-0030-2 = WGDBS2601424' in notices[0]

    def test_INV가_없는_단일_PO_DN도_DN_번호로_나간다고_알린다(self, make_service, capsys):
        """알림이 없으면 번호를 저장하기 전에 돌린 사람은 'DNO로 나오네'만 본다 (2026-11-11 실사용)"""
        make, made = make_service
        svc = make(*SHEET, row('DNO-2026-0050', 'NO-0060', 'PO-X'))
        ok, notices = create_fi.generate_fi(svc, 'DNO-2026-0050', set())
        assert ok and numbers_made(made) == ['DNO-2026-0050']
        assert len(notices) == 1 and 'Invoice No가 DNO-2026-0050로 나갑니다' in notices[0]
        assert '저장한 뒤' in notices[0] and '권장 번호: INV-2026-0050' in notices[0]
        assert notices[0] in capsys.readouterr().out      # 요약뿐 아니라 그 자리에서도 보인다

    def test_INV가_없는_발주번호_통합도_알린다(self, make_service, capsys):
        make, made = make_service
        ok, notices = create_fi.generate_fi_by_po(make(*SHEET), 'WGDBS2601411', set())
        assert ok and numbers_made(made) == ['DNO-2026-0030']
        assert len(notices) == 1 and 'INV 번호가 없어' in notices[0]

    def test_INV_컬럼이_없으면_컬럼부터_만들라고_알린다(self, make_service, capsys):
        make, made = make_service
        rows = [{k: v for k, v in r.items() if k != INV} for r in SHEET[4:]]
        ok, notices = create_fi.generate_fi(make(*rows), 'DNO-2026-0030', set())
        assert ok and len(made) == 2
        assert f"'{INV}' 컬럼" in notices[0]

    def test_어디에도_없는_ID(self, make_service, capsys):
        make, made = make_service
        ok, _ = create_fi.generate_fi(make(*SHEET), 'DNO-2026-9999', set())
        assert not ok and made == []
        assert 'DN_ID·INV 번호' in capsys.readouterr().out


# === 템플릿 ===================================================================

class TestTemplate:
    """양식은 템플릿 파일이 소유한다 — 생성기가 쓰는 칸에 라벨이 있어야 한다"""

    def test_DN_번호_칸이_있다(self):
        ws = openpyxl.load_workbook(FI_TEMPLATE_FILE).worksheets[0]
        label = ws[fi_generator.CELL_DN_REF_LABEL].value
        assert label and fi_generator.DN_REF_LABEL in label
        merged = {str(r) for r in ws.merged_cells.ranges}
        assert {'A9:B9', 'C9:E9'} <= merged       # 위의 Customer PO 줄과 같은 모양
        assert fi_generator.CELL_DN_REF == 'C9'
        assert ws[fi_generator.CELL_DN_REF].value is None

    def test_번호_세_줄은_위에서부터_C7_C8_C9(self):
        """생성기가 세 줄을 한 번에 손질한다 — 칸이 흩어지면 엉뚱한 행의 높이를 잰다"""
        cells = (fi_generator.CELL_PO_NO, fi_generator.CELL_PO_DATE, fi_generator.CELL_DN_REF)
        start, cols = fi_generator.REF_START_ROW, fi_generator.REF_MERGED_COLS
        assert cells == tuple(f'{cols[0]}{start + n}' for n in range(3))
        ws = openpyxl.load_workbook(FI_TEMPLATE_FILE).worksheets[0]
        merged = {str(r) for r in ws.merged_cells.ranges}
        assert {f'{cols[0]}{start + n}:{cols[-1]}{start + n}' for n in range(3)} <= merged

    def test_여러_개가_실리는_칸은_줄바꿈과_행_높이를_확보한다(self):
        """병합 칸은 넘친 글자를 옆으로 흘리지 않고 잘라낸다. '셀에 맞춤'으로 줄이면
        PO 5개가 4pt 남짓으로 찍힌다 (2026-11-11 실측 DNO-2026-0020)"""
        assert 'layout_merged_lines(' in inspect.getsource(fi_generator)
        ws = openpyxl.load_workbook(FI_TEMPLATE_FILE).worksheets[0]
        for addr in (fi_generator.CELL_PO_NO, fi_generator.CELL_DN_REF):
            assert not ws[addr].alignment.shrink_to_fit, addr
