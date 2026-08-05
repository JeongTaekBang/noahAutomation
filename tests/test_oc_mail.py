"""
Order Confirmation 메일 발송 테스트 (해외 전용)
===============================================

국내 문서(거래명세표·납기현황)와 **조인키가 다르다** — 사업자번호가 아니라 고객코드로
`Customer_해외`를 찾는다. 이 파일은 그 차이가 무너지지 않는지를 지킨다.

메일 작성은 mock으로 대체하여 실제 메일이 나가지 않게 합니다.
"""

from types import SimpleNamespace

import pandas as pd
import pytest

import create_oc
from create_oc import _mail_oc, prepare_mail_options
from po_generator import mail_cli, mailer
from po_generator.config import (
    CUSTOMER_EXPORT_SHEET,
    OC_MAIL_ATTACH_FORMAT,
    OC_MAIL_BODY,
    OC_MAIL_SUBJECT,
    SUPPLIER_INFO,
)
from po_generator.mail_cli import MailMode, MailOptions
from po_generator.mailer import Recipient, find_recipient_overseas, render_template
from po_generator.utils import load_customer_overseas, normalize_customer_code

# 메일 백엔드 격리(user_settings 무시)는 conftest.isolate_mail_settings가 전 테스트에 적용
# 발주번호 수집(collect_customer_po)·모드 판정 규칙은 test_mailer.py가 단일 소유로 검증


@pytest.fixture
def df_customer():
    """Customer_해외 마스터 (load_customer_overseas 출력 형태)

    실제 시트와 같은 헤더를 쓴다 — 'C-code by 해외'와 '고객코드'가 **둘 다** 있고,
    SO_해외와 맞물리는 건 앞의 것이다 (뒤엣것은 AX 번호).
    """
    df = pd.DataFrame({
        'C-code by 해외': ['C-0095', 'C-0079', 'C-0001'],
        '고객코드': [2352.0, 1883.0, 9999.0],
        '고객명': ['WATERGATES GMBH', 'SECTORIEL SA', 'NO MAIL CO'],
        '수신자 이메일': ['buyer@watergates.de', 'a@sectoriel.fr; b@sectoriel.fr', None],
        '참조 이메일': [None, 'cc@sectoriel.fr, a@sectoriel.fr', None],
    })
    df['_고객코드_정규화'] = df['C-code by 해외'].map(normalize_customer_code)
    return df


@pytest.fixture
def order():
    """SO_해외 한 행 — 'Business registration number'에 고객코드가 들어 있다"""
    return pd.Series({
        'Business registration number': 'C-0095',
        'Customer name': 'WATERGATES',
        'Customer PO': '88826',
        'SO_ID': 'SOO-2026-0001',
    })


# === 고객코드 정규화 ===

class TestNormalizeCustomerCode:
    @pytest.mark.parametrize('raw, expected', [
        ('C-0095', 'C-0095'),
        ('  c-0095 ', 'C-0095'),
        ('c-0095', 'C-0095'),
        ('', ''),
        (None, ''),
        (float('nan'), ''),
        ('nan', ''),
    ])
    def test_normalize(self, raw, expected):
        assert normalize_customer_code(raw) == expected

    def test_사업자번호처럼_숫자만_남기지_않는다(self):
        """국내 정규화(normalize_biz_no)를 그대로 쓰면 'C-0095'가 '0095'가 된다"""
        assert normalize_customer_code('C-0095') == 'C-0095'


# === 수신자 조회 (Customer_해외) ===

class TestFindRecipientOverseas:
    def test_고객코드로_찾는다(self, df_customer):
        r = find_recipient_overseas('C-0095', df_customer)
        assert r is not None
        assert r.customer_name == 'WATERGATES GMBH'
        assert r.to == ('buyer@watergates.de',)

    def test_한_셀의_여러_주소를_분리한다(self, df_customer):
        r = find_recipient_overseas('C-0079', df_customer)
        assert r.to == ('a@sectoriel.fr', 'b@sectoriel.fr')

    def test_참조가_수신자와_겹치면_제외한다(self, df_customer):
        """a@ 는 To에 이미 있으므로 CC에서 빠져야 한다 (중복 수신 방지)"""
        r = find_recipient_overseas('C-0079', df_customer)
        assert r.cc == ('cc@sectoriel.fr',)

    def test_소문자_공백_키도_매칭된다(self, df_customer):
        assert find_recipient_overseas('  c-0079 ', df_customer) is not None

    def test_미등록_코드는_None(self, df_customer):
        assert find_recipient_overseas('C-9999', df_customer) is None

    def test_이메일_미입력이면_None(self, df_customer):
        assert find_recipient_overseas('C-0001', df_customer) is None

    def test_빈_키는_None(self, df_customer):
        assert find_recipient_overseas('', df_customer) is None
        assert find_recipient_overseas(None, df_customer) is None

    def test_정규화_컬럼이_없어도_동작한다(self, df_customer):
        """원본 시트를 그대로 넘겨도 KeyError 대신 조인이 되어야 한다"""
        raw = df_customer.drop(columns=['_고객코드_정규화'])
        assert find_recipient_overseas('C-0095', raw) is not None

    def test_고객코드_컬럼이_C_code_로_풀린다(self, df_customer):
        """'고객코드'(AX 번호)로 잘못 풀리면 전부 미매칭이 된다"""
        assert find_recipient_overseas('2352', df_customer) is None
        assert find_recipient_overseas('C-0095', df_customer) is not None

    def test_이메일_컬럼이_없으면_설정오류(self):
        df = pd.DataFrame({'C-code by 해외': ['C-0095'], '고객명': ['X']})
        with pytest.raises(mailer.MailConfigError, match='이메일 컬럼'):
            find_recipient_overseas('C-0095', df)


# === 영문 템플릿 ===

class TestEnglishTemplate:
    @pytest.fixture
    def recipient(self):
        return Recipient(
            customer_key='C-0095', customer_name='WATERGATES GMBH',
            to=('buyer@watergates.de',), cc=(),
        )

    def test_제목은_발신_조직과_고객_발주번호를_밝힌다(self, recipient):
        """고객은 자기 PO 번호로 메일을 찾는다 — 제목에 그게 바로 보여야 한다"""
        subject = render_template(
            OC_MAIL_SUBJECT, recipient, 'SOO-2026-0001', '2026-08-03', customer_po='88826',
        )
        assert 'Rotork Controls Korea' in subject
        assert 'Order Confirmation' in subject
        assert '88826' in subject

    def test_본문은_발주번호_확인과_자동발송_안내뿐이다(self, recipient):
        """인사말·서명 없이 짧게 — 2026-08-05 사용자 결정"""
        body = render_template(
            OC_MAIL_BODY, recipient, 'SOO-2026-0001', '2026-08-03', customer_po='88826',
        )
        assert '88826' in body
        assert 'Order Confirmation' in body
        assert 'Dear' not in body

    def test_본문에_자동발송_안내가_있다(self, recipient):
        body = render_template(OC_MAIL_BODY, recipient, 'SOO-1', '2026-08-03')
        assert 'automatically' in body.lower()

    def test_한글_상호가_들어가지_않는다(self, recipient):
        """해외 고객 메일 — 제목·본문 어디에도 {supplier}(한글)가 나오면 안 된다"""
        subject = render_template(OC_MAIL_SUBJECT, recipient, 'SOO-1', '2026-08-03')
        body = render_template(OC_MAIL_BODY, recipient, 'SOO-1', '2026-08-03')
        assert SUPPLIER_INFO.name not in subject
        assert SUPPLIER_INFO.name not in body

    def test_치환자가_남지_않는다(self, recipient):
        rendered = render_template(
            OC_MAIL_BODY, recipient, 'SOO-1', '2026-08-03', customer_po='88826',
        )
        assert '{' not in rendered

    def test_발주번호가_없으면_NA(self, recipient):
        body = render_template(OC_MAIL_BODY, recipient, 'SOO-1', '2026-08-03', customer_po='')
        assert mailer.NO_VALUE in body

    def test_첨부는_기본_PDF(self):
        assert OC_MAIL_ATTACH_FORMAT == 'pdf'


# === 확인 프롬프트 배선 (_mail_oc) ===

class TestMailOcPrompt:
    def _opts(self, df_customer, mode):
        # loader와 sheet_label은 항상 짝으로 — 한쪽만 해외면 안내 문구와 실제
        # 조회 시트가 갈린다 (df_customer가 있어 loader는 호출되지 않지만, 짝을 지킨다)
        return MailOptions(
            mode=mode, df_customer=df_customer,
            loader=load_customer_overseas, sheet_label=CUSTOMER_EXPORT_SHEET,
        )

    def test_확인에서_아니오면_메일을_만들지_않는다(self, order, df_customer, tmp_path, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'n')
        called = []
        monkeypatch.setattr(create_oc, 'create_document_mail', lambda **kw: called.append(kw))

        opts = self._opts(df_customer, MailMode.ASK)
        assert _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', opts) is False
        assert called == []

    def test_확인에서_예면_영문_템플릿으로_만든다(self, order, df_customer, tmp_path, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'y')
        called = []

        def fake_mail(**kw):
            called.append(kw)
            return mailer.MailResult(success=True, sent=False, recipient=kw['recipient'])

        monkeypatch.setattr(create_oc, 'create_document_mail', fake_mail)

        opts = self._opts(df_customer, MailMode.ASK)
        assert _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', opts) is True
        assert called[0]['subject_template'] == OC_MAIL_SUBJECT
        assert called[0]['body_template'] == OC_MAIL_BODY
        assert called[0]['attach_format'] == OC_MAIL_ATTACH_FORMAT
        assert called[0]['send'] is False
        assert called[0]['recipient'].to == ('buyer@watergates.de',)

    def test_draft_모드는_묻지_않는다(self, order, df_customer, tmp_path, monkeypatch):
        def boom(_):
            raise AssertionError('묻지 않아야 한다')
        monkeypatch.setattr('builtins.input', boom)
        monkeypatch.setattr(
            create_oc, 'create_document_mail',
            lambda **kw: mailer.MailResult(success=True, sent=False, recipient=kw['recipient']),
        )

        opts = self._opts(df_customer, MailMode.DRAFT)
        assert _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', opts) is True

    def test_off_모드는_아무것도_안_한다(self, order, tmp_path, monkeypatch):
        monkeypatch.setattr(
            create_oc, 'create_document_mail',
            lambda **kw: pytest.fail('OFF 모드에서 메일이 나가면 안 된다'),
        )
        assert _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', MailOptions.disabled()) is True

    def test_미등록_수신자는_묻지도_않는다(self, df_customer, tmp_path, monkeypatch):
        def boom(_):
            raise AssertionError('묻지 않아야 한다')
        monkeypatch.setattr('builtins.input', boom)

        order = pd.Series({
            'Business registration number': 'C-9999',
            'Customer name': '미등록',
        })
        opts = self._opts(df_customer, MailMode.ASK)
        assert _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', opts) is False

    def test_다중_아이템의_발주번호가_모두_실린다(self, order, df_customer, tmp_path, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'y')
        called = []

        def fake_mail(**kw):
            called.append(kw)
            return mailer.MailResult(success=True, sent=False, recipient=kw['recipient'])

        monkeypatch.setattr(create_oc, 'create_document_mail', fake_mail)

        items = pd.DataFrame({'Customer PO': ['88826', '88827']})
        opts = self._opts(df_customer, MailMode.ASK)
        _mail_oc(order, tmp_path / 'x.xlsx', 'SOO-1', opts, items_df=items)
        assert called[0]['customer_po'] == '88826, 88827'


# === CLI 옵션 — OC 고유 배선만 (모드 판정 규칙 자체는 test_mailer.py가 소유) ===

class TestPrepareMailOptions:
    def test_해외_마스터를_읽도록_배선된다(self, monkeypatch):
        """국내(Customer_국내)를 읽으면 해외 고객이 전부 미등록으로 뜬다"""
        monkeypatch.setattr('sys.stdin', SimpleNamespace(isatty=lambda: True))
        args = SimpleNamespace(mail=True, send=False, no_mail=False)
        opts = prepare_mail_options(args)
        assert opts.sheet_label == CUSTOMER_EXPORT_SHEET
        assert opts.loader is load_customer_overseas
        assert opts.loader is not mail_cli.load_customer_domestic
