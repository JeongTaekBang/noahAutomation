"""
거래명세표 메일 발송 테스트
============================

Outlook COM은 mock으로 대체하여 실제 메일이 나가지 않게 합니다.
"""

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

import create_ts
from create_ts import _mail_ts
from po_generator import mail_cli, mailer
from po_generator.mail_cli import (
    MailMode,
    MailOptions,
    collect_customer_po as _collect_customer_po,
    confirm as _confirm,
    format_mail_date as _format_mail_date,
    prepare_mail_options,
    resolve_mail_mode,
)
from po_generator.mailer import (
    MailBackend,
    MailConfigError,
    Recipient,
    create_ts_mail,
    find_recipient,
    render_template,
    resolve_backend,
    split_emails,
)
from po_generator.utils import normalize_biz_no, resolve_column

# 메일 백엔드 격리(user_settings 무시)는 conftest.isolate_mail_settings가 전 테스트에 적용


# === Fixtures ===

@pytest.fixture
def df_customer():
    """Customer_국내 마스터 (load_customer_domestic 출력 형태)"""
    df = pd.DataFrame({
        '사업자번호': ['220-81-21175', '123-45-67890', '999-99-99999'],
        '거래처명': ['가나밸브', '다라산업', '메일없는곳'],
        'Customer Name ENG': ['GANA VALVE', None, 'NO MAIL CO'],
        '이메일': ['buyer@gana.co.kr', 'a@dara.kr; b@dara.kr', None],
        '참조메일': [None, 'cc@dara.kr', None],
    })
    df['_사업자번호_정규화'] = df['사업자번호'].map(normalize_biz_no)
    return df


@pytest.fixture
def recipient():
    return Recipient(
        customer_key='220-81-21175',
        customer_name='가나밸브',
        to=('buyer@gana.co.kr',),
        cc=('fixed@rotork.com',),
    )


# === 사업자번호 정규화 ===

class TestNormalizeBizNo:
    def test_hyphen_removed(self):
        assert normalize_biz_no('220-81-21175') == '2208121175'

    def test_spaces_and_mixed_format(self):
        assert normalize_biz_no(' 220 81 21175 ') == '2208121175'

    def test_numeric_cell_with_float_tail(self):
        """Excel이 숫자로 읽어 2208121175.0이 된 경우"""
        assert normalize_biz_no('2208121175.0') == '2208121175'
        assert normalize_biz_no(2208121175) == '2208121175'

    def test_empty_values(self):
        assert normalize_biz_no(None) == ''
        assert normalize_biz_no('') == ''
        assert normalize_biz_no('nan') == ''
        assert normalize_biz_no(float('nan')) == ''


# === 메일 주소 파싱 ===

class TestSplitEmails:
    def test_single(self):
        assert split_emails('a@b.com') == ('a@b.com',)

    def test_multiple_separators(self):
        assert split_emails('a@b.com; c@d.com, e@f.com\ng@h.com') == (
            'a@b.com', 'c@d.com', 'e@f.com', 'g@h.com',
        )

    def test_duplicates_removed_order_kept(self):
        assert split_emails('a@b.com; c@d.com; a@b.com') == ('a@b.com', 'c@d.com')

    def test_invalid_dropped(self):
        """오타 주소는 버리고 유효한 것만 남긴다"""
        assert split_emails('notanemail; ok@b.com; @nope.com') == ('ok@b.com',)

    def test_empty(self):
        assert split_emails(None) == ()
        assert split_emails('') == ()
        assert split_emails(float('nan')) == ()


# === 수신자 조회 ===

class TestFindRecipient:
    def test_match_by_biz_no(self, df_customer):
        r = find_recipient('220-81-21175', df_customer)
        assert r is not None
        assert r.customer_name == '가나밸브'
        assert r.to == ('buyer@gana.co.kr',)

    def test_match_ignores_hyphen_format(self, df_customer):
        """DN에 하이픈 없이 들어와도 조인된다"""
        r = find_recipient('2208121175', df_customer)
        assert r is not None
        assert r.to == ('buyer@gana.co.kr',)

    def test_multiple_to_and_cc(self, df_customer, monkeypatch):
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ('fixed@rotork.com',))
        r = find_recipient('123-45-67890', df_customer)
        assert r.to == ('a@dara.kr', 'b@dara.kr')
        assert r.cc == ('fixed@rotork.com', 'cc@dara.kr')

    def test_fixed_cc_always_added(self, df_customer, monkeypatch):
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ('fixed@rotork.com',))
        r = find_recipient('220-81-21175', df_customer)
        assert r.cc == ('fixed@rotork.com',)

    def test_cc_excludes_address_already_in_to(self, df_customer, monkeypatch):
        """To와 CC에 같은 주소가 겹치면 CC에서 뺀다"""
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ('BUYER@gana.co.kr', 'other@rotork.com'))
        r = find_recipient('220-81-21175', df_customer)
        assert r.cc == ('other@rotork.com',)

    def test_unregistered_biz_no_returns_none(self, df_customer):
        assert find_recipient('000-00-00000', df_customer) is None

    def test_registered_but_no_email_returns_none(self, df_customer):
        """마스터에 있지만 메일이 비어 있으면 None (조용히 성공하지 않음)"""
        assert find_recipient('999-99-99999', df_customer) is None

    def test_empty_biz_no_returns_none(self, df_customer):
        assert find_recipient('', df_customer) is None
        assert find_recipient(None, df_customer) is None

    def test_missing_email_column_raises(self, df_customer):
        """이메일 컬럼 자체가 없으면 명확히 실패해야 한다"""
        df = df_customer.drop(columns=['이메일', '참조메일'])
        with pytest.raises(MailConfigError, match='이메일 컬럼'):
            find_recipient('220-81-21175', df)

    def test_alias_column_name(self, df_customer):
        """헤더가 'Email'이어도 인식된다"""
        df = df_customer.rename(columns={'이메일': 'Email'})
        r = find_recipient('220-81-21175', df)
        assert r.to == ('buyer@gana.co.kr',)

    @pytest.mark.parametrize('to_header,cc_header', [
        ('수신자 이메일', '참조 이메일'),   # 실제 Customer_국내 헤더
        ('수신자이메일', '참조이메일'),
        ('이메일', '참조메일'),
        ('Email', 'CC'),
    ])
    def test_to_and_cc_headers_never_collide(self, df_customer, to_header, cc_header, monkeypatch):
        """'참조 이메일'도 '이메일'을 포함한다 — To 컬럼으로 잘못 잡히면 안 된다"""
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ())
        df = df_customer.rename(columns={'이메일': to_header, '참조메일': cc_header})

        assert resolve_column(df.columns, 'customer_email') == to_header
        assert resolve_column(df.columns, 'customer_email_cc') == cc_header

        r = find_recipient('123-45-67890', df)
        assert r.to == ('a@dara.kr', 'b@dara.kr')   # 수신자 컬럼에서
        assert r.cc == ('cc@dara.kr',)              # 참조 컬럼에서

    def test_empty_email_column_is_not_a_crash(self, df_customer):
        """컬럼만 만들고 주소는 아직 안 채운 상태 — 조용히 None"""
        df = df_customer.assign(이메일=None)
        assert find_recipient('220-81-21175', df) is None

    def test_raw_sheet_without_normalized_column(self, df_customer):
        """정규화 컬럼 없는 원본 시트를 넘겨도 KeyError 없이 동작"""
        df = df_customer.drop(columns=['_사업자번호_정규화'])
        r = find_recipient('220-81-21175', df)
        assert r is not None
        assert r.to == ('buyer@gana.co.kr',)

    def test_duplicate_cc_deduped_case_insensitively(self, df_customer, monkeypatch):
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ('CC@dara.kr', 'cc@dara.kr'))
        r = find_recipient('123-45-67890', df_customer)
        # 고정 CC 2개(대소문자만 다름) + 시트의 cc@dara.kr → 최종 1개
        assert r.cc == ('CC@dara.kr',)


# === 발주번호 수집 ===

class TestCollectCustomerPo:
    def test_single_po(self):
        items = pd.DataFrame({'Customer PO': ['VT-C35-250820-01-R1']})
        assert _collect_customer_po(pd.Series(dtype=object), items) == 'VT-C35-250820-01-R1'

    def test_merged_lists_every_distinct_po(self):
        """월합은 DN마다 발주번호가 다르다 — 첫 건만 쓰면 거래처에 틀린 번호가 간다"""
        items = pd.DataFrame({'Customer PO': [
            'VT-C35-250820-01-R1', 'VT-C35-250820-01-R1', 'VT-C35-251107-01-R1',
        ]})
        out = _collect_customer_po(pd.Series(dtype=object), items)
        assert out == 'VT-C35-250820-01-R1, VT-C35-251107-01-R1'

    def test_empty_becomes_na(self):
        """발주번호 없는 DN이 1330건 중 95건 — 빈칸이 아니라 N/A로"""
        items = pd.DataFrame({'Customer PO': [None, float('nan')]})
        assert _collect_customer_po(pd.Series(dtype=object), items) == 'N/A'

    def test_no_column_at_all_becomes_na(self):
        items = pd.DataFrame({'Item': ['수동 핸들']})
        assert _collect_customer_po(pd.Series(dtype=object), items) == 'N/A'

    def test_partial_blanks_are_skipped(self):
        items = pd.DataFrame({'Customer PO': ['P26-0317-02-J', None, '  ']})
        assert _collect_customer_po(pd.Series(dtype=object), items) == 'P26-0317-02-J'

    def test_falls_back_to_order_data(self):
        """items_df가 없으면 주문 행에서 꺼낸다"""
        order = pd.Series({'Customer PO': 'P26-0317-02-J'})
        assert _collect_customer_po(order, None) == 'P26-0317-02-J'

    def test_falls_back_to_na_when_nothing(self):
        assert _collect_customer_po(pd.Series(dtype=object), None) == 'N/A'


class TestCustomerPoInTemplate:
    def test_rendered_into_body(self, recipient):
        out = render_template('발주번호: {customer_po}', recipient, 'DN-1', '2026-07-27',
                              'VT-C35-250820-01-R1')
        assert out == '발주번호: VT-C35-250820-01-R1'

    def test_empty_string_becomes_na(self, recipient):
        out = render_template('발주번호: {customer_po}', recipient, 'DN-1', '2026-07-27', '')
        assert out == '발주번호: N/A'

    def test_default_is_na(self, recipient):
        out = render_template('발주번호: {customer_po}', recipient, 'DN-1', '2026-07-27')
        assert out == '발주번호: N/A'


# === 메일 날짜 ===

class TestFormatMailDate:
    def test_normal_date(self):
        assert _format_mail_date(pd.Timestamp('2026-01-09')) == '2026-01-09'

    def test_string_date(self):
        assert _format_mail_date('2026-01-09') == '2026-01-09'

    def test_none_falls_back_to_today(self):
        """선수금 거래명세표는 SO_국내 기반이라 출고일이 아예 없다 — 죽으면 안 된다"""
        out = _format_mail_date(None)
        assert out == datetime.now().strftime('%Y-%m-%d')

    def test_nat_and_nan_fall_back_to_today(self):
        today = datetime.now().strftime('%Y-%m-%d')
        assert _format_mail_date(pd.NaT) == today
        assert _format_mail_date(float('nan')) == today

    def test_garbage_falls_back_to_today(self):
        assert _format_mail_date('출고예정') == datetime.now().strftime('%Y-%m-%d')


# === 메일 모드 결정 ===

def _args(mail=False, send=False, no_mail=False):
    return SimpleNamespace(mail=mail, send=send, no_mail=no_mail)


class TestResolveMailMode:
    def test_default_is_ask_on_terminal(self):
        assert resolve_mail_mode(_args(), is_tty=True) is MailMode.ASK

    def test_default_is_off_when_not_a_terminal(self):
        """배치/파이프 실행에서 input()으로 멈추면 안 된다"""
        assert resolve_mail_mode(_args(), is_tty=False) is MailMode.OFF

    def test_explicit_flags_survive_non_tty(self):
        """명시적 --mail/--send는 비대화형에서도 존중"""
        assert resolve_mail_mode(_args(mail=True), is_tty=False) is MailMode.DRAFT
        assert resolve_mail_mode(_args(send=True), is_tty=False) is MailMode.SEND

    def test_no_mail_wins_over_everything(self):
        assert resolve_mail_mode(_args(mail=True, send=True, no_mail=True), is_tty=True) is MailMode.OFF

    def test_send_wins_over_mail(self):
        assert resolve_mail_mode(_args(mail=True, send=True), is_tty=True) is MailMode.SEND


class TestPrepareMailOptions:
    """공용 옵션 구성 헬퍼 — 세 CLI가 상수만 넘겨 쓰므로 규칙 검증은 여기 한 곳"""

    def test_no_mail은_OFF(self, monkeypatch):
        monkeypatch.setattr('sys.stdin', SimpleNamespace(isatty=lambda: True))
        opts = prepare_mail_options(_args(no_mail=True), attach_format='pdf')
        assert opts.mode is MailMode.OFF

    def test_비대화형은_묻지_않는다(self, monkeypatch):
        """배치 실행이 input()에서 멈추면 안 된다"""
        monkeypatch.setattr('sys.stdin', SimpleNamespace(isatty=lambda: False))
        assert prepare_mail_options(_args(), attach_format='pdf').mode is MailMode.OFF

    def test_대화형_기본은_건별_확인_및_배너(self, monkeypatch, capsys):
        monkeypatch.setattr('sys.stdin', SimpleNamespace(isatty=lambda: True))
        opts = prepare_mail_options(_args(), attach_format='pdf')
        assert opts.mode is MailMode.ASK
        assert 'PDF' in capsys.readouterr().out

    def test_기본_마스터는_국내(self, monkeypatch):
        monkeypatch.setattr('sys.stdin', SimpleNamespace(isatty=lambda: True))
        opts = prepare_mail_options(_args(mail=True), attach_format='pdf')
        assert opts.sheet_label == mail_cli.CUSTOMER_DOMESTIC_SHEET
        assert opts.loader is None  # None = 호출 시점에 국내 로더로 해석 (monkeypatch 가능)


class TestConfirm:
    @pytest.mark.parametrize('answer', ['y', 'Y', 'yes', 'YES', ' y ', 'ㅛ', '네', 'ㅇ'])
    def test_accepts_yes_forms(self, answer, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: answer)
        assert _confirm('보낼까요? ') is True

    @pytest.mark.parametrize('answer', ['', 'n', 'N', 'no', 'ㅜ', 'abc'])
    def test_defaults_to_no(self, answer, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: answer)
        assert _confirm('보낼까요? ') is False

    def test_eof_is_no(self, monkeypatch):
        """stdin이 닫혀 있으면 발송하지 않는다"""
        def raise_eof(_):
            raise EOFError
        monkeypatch.setattr('builtins.input', raise_eof)
        assert _confirm('보낼까요? ') is False

    def test_ctrl_c_is_no(self, monkeypatch):
        def raise_int(_):
            raise KeyboardInterrupt
        monkeypatch.setattr('builtins.input', raise_int)
        assert _confirm('보낼까요? ') is False


# === 확인 프롬프트 배선 (_mail_ts) ===

class TestMailTsPrompt:
    @pytest.fixture
    def order(self):
        return pd.Series({
            'Business registration number': '220-81-21175',
            'Customer name': '가나밸브',
            '출고일': pd.Timestamp('2026-01-09'),
        })

    def test_ask_no_does_not_open_outlook(self, order, df_customer, tmp_path, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'n')
        called = []
        monkeypatch.setattr(create_ts, 'create_ts_mail', lambda **kw: called.append(kw))

        opts = MailOptions(mode=MailMode.ASK, df_customer=df_customer)
        assert _mail_ts(order, tmp_path / 'x.xlsx', 'DN-1', opts) is False
        assert called == []

    def test_ask_yes_opens_outlook(self, order, df_customer, tmp_path, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'y')
        called = []

        def fake_mail(**kw):
            called.append(kw)
            return mailer.MailResult(success=True, sent=False, recipient=kw['recipient'])

        monkeypatch.setattr(create_ts, 'create_ts_mail', fake_mail)

        opts = MailOptions(mode=MailMode.ASK, df_customer=df_customer)
        assert _mail_ts(order, tmp_path / 'x.xlsx', 'DN-1', opts) is True
        assert called and called[0]['send'] is False

    def test_draft_mode_skips_prompt(self, order, df_customer, tmp_path, monkeypatch):
        """--mail은 묻지 않는다 — input이 호출되면 실패"""
        def boom(_):
            raise AssertionError('묻지 않아야 한다')
        monkeypatch.setattr('builtins.input', boom)
        monkeypatch.setattr(
            create_ts, 'create_ts_mail',
            lambda **kw: mailer.MailResult(success=True, sent=False, recipient=kw['recipient']),
        )

        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)
        assert _mail_ts(order, tmp_path / 'x.xlsx', 'DN-1', opts) is True

    def test_off_mode_is_noop(self, order, tmp_path, monkeypatch):
        monkeypatch.setattr(
            create_ts, 'create_ts_mail',
            lambda **kw: pytest.fail('OFF 모드에서 메일이 나가면 안 된다'),
        )
        assert _mail_ts(order, tmp_path / 'x.xlsx', 'DN-1', MailOptions.disabled()) is True

    def test_unregistered_recipient_never_prompts(self, df_customer, tmp_path, monkeypatch):
        """수신자가 없으면 물어볼 것도 없다"""
        def boom(_):
            raise AssertionError('묻지 않아야 한다')
        monkeypatch.setattr('builtins.input', boom)

        order = pd.Series({
            'Business registration number': '000-00-00000',
            'Customer name': '미등록거래처',
        })
        opts = MailOptions(mode=MailMode.ASK, df_customer=df_customer)
        assert _mail_ts(order, tmp_path / 'x.xlsx', 'DN-1', opts) is False


# === 마스터 지연 로딩 ===

class TestCustomerMasterLazyLoad:
    """MailOptions는 po_generator/mail_cli.py에 있다 (create_ts / delivery_status 공유).

    지연 로딩을 가로채려면 그 모듈의 이름을 패치해야 한다.
    """

    def test_loads_once_and_caches(self, df_customer, monkeypatch):
        calls = []

        def fake_load():
            calls.append(1)
            return df_customer

        monkeypatch.setattr(mail_cli, 'load_customer_domestic', fake_load)
        opts = MailOptions(mode=MailMode.ASK)
        assert opts.customer_master() is df_customer
        assert opts.customer_master() is df_customer
        assert len(calls) == 1

    def test_missing_email_column_reported_once_then_silent(self, df_customer, monkeypatch, capsys):
        """이메일 컬럼이 없으면 한 번만 안내하고 이후 실행 내내 조용히 건너뛴다"""
        no_email = df_customer.drop(columns=['이메일', '참조메일'])
        monkeypatch.setattr(mail_cli, 'load_customer_domestic', lambda: no_email)

        opts = MailOptions(mode=MailMode.ASK)
        assert opts.customer_master() is None
        first = capsys.readouterr().out
        assert '이메일 컬럼이 없습니다' in first

        assert opts.customer_master() is None
        assert capsys.readouterr().out == ''

    def test_load_failure_reported_once(self, monkeypatch, capsys):
        def boom():
            raise FileNotFoundError('소스 파일 없음')
        monkeypatch.setattr(mail_cli, 'load_customer_domestic', boom)

        opts = MailOptions(mode=MailMode.ASK)
        assert opts.customer_master() is None
        assert '소스 파일 없음' in capsys.readouterr().out
        assert opts.customer_master() is None
        assert capsys.readouterr().out == ''


# === 템플릿 치환 ===

class TestRenderTemplate:
    def test_substitution(self, recipient):
        out = render_template('{customer} / {doc_id} / {date}', recipient, 'DN-1', '2026-07-27')
        assert out == '가나밸브 / DN-1 / 2026-07-27'

    def test_unknown_placeholder_keeps_original(self, recipient):
        """알 수 없는 치환자가 있어도 예외 없이 원본 유지"""
        tpl = '{customer} {없는키}'
        assert render_template(tpl, recipient, 'DN-1', '2026-07-27') == tpl

    def test_customer_en_uses_english_name(self):
        r = Recipient(customer_key='1', customer_name='가나밸브', to=('a@b.com',), cc=(),
                      customer_name_en='GANA VALVE')
        assert render_template('Dear {customer_en},', r, 'DN-1', '2026-07-27') == 'Dear GANA VALVE,'

    def test_customer_en_falls_back_to_korean(self):
        """영문명이 비어 있는 6% 거래처 — 빈칸 대신 한글명으로"""
        r = Recipient(customer_key='1', customer_name='주식회사 진테크', to=('a@b.com',), cc=())
        assert render_template('Dear {customer_en},', r, 'DN-1', '2026-07-27') == 'Dear 주식회사 진테크,'


class TestEnglishNameLookup:
    def test_english_name_loaded(self, df_customer):
        r = find_recipient('220-81-21175', df_customer)
        assert r.customer_name_en == 'GANA VALVE'
        assert r.display_name_en == 'GANA VALVE'

    def test_missing_english_name_falls_back(self, df_customer):
        """다라산업은 Customer Name ENG가 비어 있다"""
        r = find_recipient('123-45-67890', df_customer)
        assert r.customer_name_en == ''
        assert r.display_name_en == '다라산업'

    def test_no_english_column_at_all(self, df_customer):
        df = df_customer.drop(columns=['Customer Name ENG'])
        r = find_recipient('220-81-21175', df)
        assert r.display_name_en == '가나밸브'


# === 백엔드 결정 ===

class TestResolveBackend:
    def test_explicit_settings_win(self):
        assert resolve_backend('outlook') is MailBackend.OUTLOOK
        assert resolve_backend('eml') is MailBackend.EML

    def test_reads_setting_at_call_time(self, monkeypatch):
        """설정을 기본 인자로 박아두면 import 시점에 고정돼 변경이 안 먹는다"""
        monkeypatch.setattr(mailer, 'TS_MAIL_BACKEND', 'eml')
        assert resolve_backend() is MailBackend.EML
        monkeypatch.setattr(mailer, 'TS_MAIL_BACKEND', 'outlook')
        assert resolve_backend() is MailBackend.OUTLOOK

    def test_auto_draft_is_eml_even_with_com(self, monkeypatch):
        """auto 초안은 COM이 있어도 .eml — 사용자의 기본 메일 앱에서 열려야 한다

        COM 초안은 항상 클래식 Outlook 창을 띄운다. 클래식이 설치되는 순간
        평소 새 Outlook을 쓰는 PC에서 초안이 낯선 옛 창으로 뜨는 회귀가 있었다
        (2026-07-30 배포판 테스트 실측) — 이 테스트가 그 회귀 방지다.
        """
        monkeypatch.setattr(mailer, 'outlook_com_available', lambda: True)
        assert resolve_backend('auto') is MailBackend.EML
        assert resolve_backend('auto', send=False) is MailBackend.EML

    def test_auto_send_uses_com_when_available(self, monkeypatch):
        """즉시 발송은 창 없이 나가야 하므로 COM이 있으면 COM"""
        monkeypatch.setattr(mailer, 'outlook_com_available', lambda: True)
        assert resolve_backend('auto', send=True) is MailBackend.OUTLOOK

    def test_auto_send_falls_back_to_eml_without_com(self, monkeypatch):
        """COM이 없으면 발송 요청도 .eml 초안으로 강등 (호출부가 안내)"""
        monkeypatch.setattr(mailer, 'outlook_com_available', lambda: False)
        assert resolve_backend('auto', send=True) is MailBackend.EML

    def test_explicit_outlook_wins_for_draft(self, monkeypatch):
        """명시 설정('outlook')은 초안도 COM으로 — auto 규칙보다 우선"""
        monkeypatch.setattr(mailer, 'outlook_com_available', lambda: True)
        assert resolve_backend('outlook', send=False) is MailBackend.OUTLOOK

    def test_com_probe_runs_once(self, monkeypatch):
        """COM 실패는 실행당 1회만 조사한다 (건마다 재시도하면 느려짐)"""
        monkeypatch.setattr(mailer, '_com_available', None)
        calls = []

        def boom():
            calls.append(1)
            raise MailConfigError('서버 실행이 실패했습니다')

        monkeypatch.setattr(mailer, '_get_outlook', boom)
        assert mailer.outlook_com_available() is False
        assert mailer.outlook_com_available() is False
        assert len(calls) == 1


# === .eml 초안 ===

class TestBuildEml:
    @pytest.fixture
    def pdf(self, tmp_path):
        p = tmp_path / "거래명세표.pdf"
        p.write_bytes(b"%PDF-1.4\ntest\n%%EOF\n")
        return p

    def _parse(self, path):
        import email
        import email.policy
        # policy=default라야 EmailMessage(get_body/iter_attachments)로 파싱된다
        return email.message_from_bytes(Path(path).read_bytes(), policy=email.policy.default)

    def test_x_unsent_header_present(self, recipient, pdf, tmp_path):
        """X-Unsent가 없으면 '받은 메일'로 열려 [보내기]가 없다 — 이 기능의 핵심"""
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf,), tmp_path)
        assert self._parse(eml)['X-Unsent'] == '1'

    def test_no_date_header(self, recipient, pdf, tmp_path):
        """Date가 있으면 수신 메일로 취급될 수 있다"""
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf,), tmp_path)
        assert self._parse(eml)['Date'] is None

    def test_to_and_cc_written(self, recipient, pdf, tmp_path):
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf,), tmp_path)
        msg = self._parse(eml)
        assert 'buyer@gana.co.kr' in msg['To']
        assert 'fixed@rotork.com' in msg['Cc']

    def test_korean_subject_and_body_roundtrip(self, recipient, pdf, tmp_path):
        """한글 제목/본문이 깨지지 않아야 한다"""
        subject = '[로토크] 거래명세표 송부 - 가나밸브'
        body = '가나밸브 담당자님,\n\n안녕하세요.\n'
        eml = mailer.build_eml(recipient, subject, body, (pdf,), tmp_path)
        msg = self._parse(eml)

        from email.header import decode_header, make_header
        assert str(make_header(decode_header(msg['Subject']))) == subject
        assert '가나밸브 담당자님' in msg.get_body(('plain',)).get_content()

    def test_pdf_attached_with_original_name(self, recipient, pdf, tmp_path):
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf,), tmp_path)
        parts = [p for p in self._parse(eml).iter_attachments()]
        assert len(parts) == 1
        assert parts[0].get_filename() == '거래명세표.pdf'
        assert parts[0].get_content_type() == 'application/pdf'
        assert parts[0].get_payload(decode=True) == pdf.read_bytes()

    def test_multiple_attachments(self, recipient, pdf, tmp_path):
        xlsx = tmp_path / "거래명세표.xlsx"
        xlsx.write_bytes(b"PK\x03\x04dummy")
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf, xlsx), tmp_path)
        names = [p.get_filename() for p in self._parse(eml).iter_attachments()]
        assert names == ['거래명세표.pdf', '거래명세표.xlsx']

    def test_no_cc_omits_header(self, pdf, tmp_path):
        r = Recipient(customer_key='1', customer_name='A', to=('a@b.com',), cc=())
        eml = mailer.build_eml(r, '제목', '본문', (pdf,), tmp_path)
        assert self._parse(eml)['Cc'] is None

    def test_html_alternative_present(self, recipient, pdf, tmp_path):
        """평문만 있으면 Outlook이 서명을 본문 '위'에 넣는다 — HTML 대체본이 있어야 제어된다"""
        eml = mailer.build_eml(recipient, '제목', '가나밸브 귀중\n\n본문\n', (pdf,), tmp_path)
        msg = self._parse(eml)

        plain = msg.get_body(('plain',))
        html = msg.get_body(('html',))
        assert plain is not None and html is not None
        assert '가나밸브 귀중' in plain.get_content()
        assert '가나밸브 귀중' in html.get_content()
        assert '<br>' in html.get_content()

    def test_wrap_body_html_is_single_top_level_block(self):
        """공용 래퍼는 최상위 블록이 하나여야 한다 (Outlook 자동 서명 억제 구조)

        인라인 `<br>` 본문이면 Outlook이 자동 서명을 본문에 끼워 넣는다.
        전체를 `<table><tr><td>` 한 칸에 담은 구조에서는 서명이 붙지 않는 것이
        실측(2026-07-29, 납기현황 메일)으로 확인됐다 — 이 구조로 회귀하면 안 된다.

        래퍼 구조는 여기 한 곳에서만 검사한다. 소비자 쪽 테스트(아래, 납기현황의
        TestHtmlBody)는 "래퍼를 그대로 쓴다"는 위임만 확인한다.
        """
        html = mailer.wrap_body_html('<p>본문</p>')
        inner = html[html.index('<body>') + len('<body>'):html.index('</body>')]
        assert inner.startswith('<table role="presentation"')
        assert inner.endswith('</table>')
        assert inner.count('<td ') == 1

    def test_eml_html_part_is_body_to_html_verbatim(self, recipient, pdf, tmp_path):
        """.eml의 HTML 대체본은 body_to_html 결과 그대로여야 한다

        아래 위임 테스트와 함께 build_eml → body_to_html → wrap_body_html 사슬이 닫힌다.
        """
        body = '가나밸브 귀중\n\n본문\n'
        eml = mailer.build_eml(recipient, '제목', body, (pdf,), tmp_path)
        html = self._parse(eml).get_body(('html',)).get_content()
        assert html.rstrip('\n') == mailer.body_to_html(body)

    def test_body_to_html_matches_wrapper(self):
        """body_to_html은 wrap_body_html 래퍼를 그대로 써야 한다 (TS·납기현황 구조 공유)"""
        html = mailer.body_to_html('첫줄\n둘째줄')
        assert html == mailer.wrap_body_html('첫줄<br>\n둘째줄')

    def test_html_and_attachment_coexist(self, recipient, pdf, tmp_path):
        """HTML 대체본을 넣어도 첨부가 살아있어야 한다 (multipart 구조 회귀)"""
        eml = mailer.build_eml(recipient, '제목', '본문', (pdf,), tmp_path)
        msg = self._parse(eml)

        assert msg.get_content_type() == 'multipart/mixed'
        atts = list(msg.iter_attachments())
        assert len(atts) == 1
        assert atts[0].get_filename() == '거래명세표.pdf'
        assert atts[0].get_payload(decode=True) == pdf.read_bytes()

    def test_html_escapes_special_characters(self, pdf, tmp_path):
        """'S&T중공업' 같은 거래처명이 HTML에서 깨지면 안 된다"""
        r = Recipient(customer_key='1', customer_name='S&T중공업', to=('a@b.com',), cc=())
        eml = mailer.build_eml(r, '제목', 'S&T중공업 <귀중>\n', (pdf,), tmp_path)
        raw_html = self._parse(eml).get_body(('html',)).get_content()

        assert 'S&amp;T중공업' in raw_html
        assert '&lt;귀중&gt;' in raw_html
        assert '<귀중>' not in raw_html   # 태그로 해석될 여지 없음


class TestBuildAttachments:
    """첨부는 여러 장일 수 있다 — 같은 날 같은 거래처로 나간 DN이 여러 건인 경우"""

    @pytest.fixture
    def docs(self, tmp_path):
        paths = []
        for name in ('TS_A', 'TS_B', 'TS_C'):
            p = tmp_path / f"{name}.xlsx"
            p.write_bytes(b'dummy')
            paths.append(p)
        return paths

    def test_single_path_still_works(self, docs, monkeypatch):
        monkeypatch.setattr(mailer, 'export_pdfs', _fake_export_pdfs)
        assert mailer.build_attachments(docs[0], 'pdf') == (docs[0].with_suffix('.pdf'),)

    def test_many_paths_keep_order(self, docs, monkeypatch):
        monkeypatch.setattr(mailer, 'export_pdfs', _fake_export_pdfs)
        assert mailer.build_attachments(docs, 'pdf') == tuple(
            p.with_suffix('.pdf') for p in docs)

    def test_converts_in_one_excel_launch(self, docs, monkeypatch):
        """장마다 Excel을 새로 띄우면 8장짜리 하루치가 하염없이 느려진다"""
        calls = []

        def spy(paths, targets=None):
            calls.append(list(paths))
            return _fake_export_pdfs(paths)

        monkeypatch.setattr(mailer, 'export_pdfs', spy)
        mailer.build_attachments(docs, 'pdf')

        assert len(calls) == 1
        assert len(calls[0]) == 3

    def test_xlsx_format_skips_excel(self, docs, monkeypatch):
        monkeypatch.setattr(
            mailer, 'export_pdfs',
            lambda paths, targets=None: pytest.fail('xlsx 첨부는 변환이 필요 없다'),
        )
        assert mailer.build_attachments(docs, 'xlsx') == tuple(docs)

    def test_both_lists_pdfs_then_xlsx(self, docs, monkeypatch):
        monkeypatch.setattr(mailer, 'export_pdfs', _fake_export_pdfs)
        got = mailer.build_attachments(docs, 'both')
        assert got == tuple(p.with_suffix('.pdf') for p in docs) + tuple(docs)

    def test_as_paths_accepts_str_path_and_sequence(self, docs):
        assert mailer.as_paths(str(docs[0])) == (docs[0],)
        assert mailer.as_paths(docs[0]) == (docs[0],)
        assert mailer.as_paths(docs) == tuple(docs)


class TestCreateViaEml:
    @pytest.fixture
    def xlsx(self, tmp_path):
        p = tmp_path / "TS.xlsx"
        p.write_bytes(b'dummy')
        return p

    def _run(self, xlsx, recipient, send, monkeypatch, tmp_path):
        opened = []
        monkeypatch.setattr(mailer, 'open_eml', lambda p: opened.append(p))
        monkeypatch.setattr(
            mailer, 'export_pdfs',
            lambda paths, targets=None: tuple(
                _write_pdf(tmp_path / (Path(p).stem + '.pdf')) for p in paths
            ),
        )
        result = create_ts_mail(
            xlsx_path=xlsx, recipient=recipient, doc_id='DN-1',
            date_str='2026-07-27', send=send, backend=MailBackend.EML,
        )
        return result, opened

    def test_opens_draft(self, xlsx, recipient, monkeypatch, tmp_path):
        result, opened = self._run(xlsx, recipient, False, monkeypatch, tmp_path)
        assert result.success is True
        assert result.backend is MailBackend.EML
        assert len(opened) == 1
        assert Path(opened[0]).suffix == '.eml'

    def test_send_degrades_to_draft_and_says_so(self, xlsx, recipient, monkeypatch, tmp_path):
        """.eml로는 자동 발송이 불가능 — 조용히 성공한 척하면 안 된다"""
        result, opened = self._run(xlsx, recipient, True, monkeypatch, tmp_path)
        assert result.success is True
        assert result.sent is False
        assert '자동 발송' in result.message
        assert len(opened) == 1

    def test_open_failure_is_reported(self, xlsx, recipient, monkeypatch, tmp_path):
        monkeypatch.setattr(
            mailer, 'export_pdfs',
            lambda paths, targets=None: (_write_pdf(tmp_path / 'x.pdf'),),
        )
        monkeypatch.setattr(
            mailer, 'open_eml',
            lambda p: (_ for _ in ()).throw(MailConfigError('연결된 앱 없음')),
        )
        with pytest.raises(MailConfigError):
            create_ts_mail(
                xlsx_path=xlsx, recipient=recipient, doc_id='DN-1',
                backend=MailBackend.EML,
            )


def _write_pdf(path: Path) -> Path:
    path.write_bytes(b"%PDF-1.4\ntest\n%%EOF\n")
    return path


def _fake_export_pdfs(paths, targets=None) -> tuple[Path, ...]:
    """Excel 없이 PDF 변환 흉내 — 경로만 .pdf로 바꿔 돌려준다"""
    return tuple(Path(p).with_suffix('.pdf') for p in paths)


# === Outlook 메일 생성 (COM mock) ===

class TestCreateTsMail:
    @pytest.fixture
    def xlsx(self, tmp_path):
        p = tmp_path / "TS_DN-2026-0001_가나밸브.xlsx"
        p.write_bytes(b'dummy')
        return p

    def _run(self, xlsx, recipient, send, mock_mail):
        mock_outlook = MagicMock()
        mock_outlook.CreateItem.return_value = mock_mail
        with patch.object(mailer, '_get_outlook', return_value=mock_outlook), \
             patch.object(mailer, 'export_pdfs', side_effect=_fake_export_pdfs):
            return create_ts_mail(
                xlsx_path=xlsx, recipient=recipient,
                doc_id='DN-2026-0001', date_str='2026-07-27', send=send,
                backend=MailBackend.OUTLOOK,
            )

    def test_draft_displays_not_sends(self, xlsx, recipient):
        mail = MagicMock()
        result = self._run(xlsx, recipient, send=False, mock_mail=mail)

        assert result.success is True
        assert result.sent is False
        mail.Display.assert_called_once()
        mail.Send.assert_not_called()

    def test_send_sends_not_displays(self, xlsx, recipient):
        mail = MagicMock()
        result = self._run(xlsx, recipient, send=True, mock_mail=mail)

        assert result.success is True
        assert result.sent is True
        mail.Send.assert_called_once()
        mail.Display.assert_not_called()

    def test_to_cc_and_attachment_set(self, xlsx, recipient):
        mail = MagicMock()
        self._run(xlsx, recipient, send=False, mock_mail=mail)

        assert mail.To == 'buyer@gana.co.kr'
        assert mail.CC == 'fixed@rotork.com'
        assert '가나밸브' in mail.Subject
        mail.Attachments.Add.assert_called_once()
        assert mail.Attachments.Add.call_args[0][0].endswith('.pdf')

    def test_many_documents_go_in_one_mail(self, xlsx, recipient, tmp_path):
        """DN이 여러 건 나간 날 — 메일은 한 통, 첨부만 여러 개"""
        second = tmp_path / "TS_DN-2026-0002_가나밸브.xlsx"
        second.write_bytes(b'dummy')

        mail = MagicMock()
        mock_outlook = MagicMock()
        mock_outlook.CreateItem.return_value = mail
        with patch.object(mailer, '_get_outlook', return_value=mock_outlook), \
             patch.object(mailer, 'export_pdfs', side_effect=_fake_export_pdfs):
            result = create_ts_mail(
                xlsx_path=[xlsx, second], recipient=recipient,
                doc_id='DN-2026-0001 외 1건', date_str='2026-08-06',
                send=False, backend=MailBackend.OUTLOOK,
            )

        assert mock_outlook.CreateItem.call_count == 1
        assert mail.Attachments.Add.call_count == 2
        assert len(result.attachments) == 2
        mail.Display.assert_called_once()

    def test_pdf_conversion_failure_reported_not_raised(self, xlsx, recipient):
        """PDF 변환이 깨져도 예외가 CLI로 튀지 않고 실패 결과로 돌아온다"""
        with patch.object(mailer, 'export_pdfs', side_effect=RuntimeError('Excel 없음')):
            result = create_ts_mail(
                xlsx_path=xlsx, recipient=recipient,
                doc_id='DN-2026-0001', send=False,
                backend=MailBackend.OUTLOOK,
            )
        assert result.success is False
        assert '첨부 파일 준비 실패' in result.message

    def test_outlook_failure_reported_not_raised(self, xlsx, recipient):
        mock_outlook = MagicMock()
        mock_outlook.CreateItem.side_effect = RuntimeError('COM 오류')
        with patch.object(mailer, '_get_outlook', return_value=mock_outlook), \
             patch.object(mailer, 'export_pdfs', side_effect=_fake_export_pdfs):
            result = create_ts_mail(
                xlsx_path=xlsx, recipient=recipient,
                doc_id='DN-2026-0001', send=False,
                backend=MailBackend.OUTLOOK,
            )
        assert result.success is False
        assert 'Outlook' in result.message


# === 국내 / 해외 경계 ===

class TestDomesticOverseasBoundary:
    """국내(사업자번호)와 해외(고객코드)는 `_build_recipient` 하나를 공유한다.

    공유 뒤에도 두 경로가 섞이지 않는지를 지킨다 — 섞이면 고객에게 남의 메일이 간다.
    """

    def test_국내_조회는_해외_마스터를_읽지_않는다(self, df_customer):
        """해외 마스터를 국내 함수에 넘기면 사업자번호 컬럼이 없어 설정 오류"""
        overseas = pd.DataFrame({
            'C-code by 해외': ['C-0095'],
            '고객명': ['WATERGATES GMBH'],
            '이메일': ['buyer@watergates.de'],
        })
        with pytest.raises(MailConfigError, match='사업자번호'):
            find_recipient('220-81-21175', overseas)

    def test_해외_코드를_국내_경로로_찾으면_안_잡힌다(self, df_customer):
        """'C-0095'를 국내 정규화에 넣으면 '0095'가 되어 엉뚱한 매칭 위험"""
        assert find_recipient('C-0095', df_customer) is None

    def test_국내_기본_CC는_TS_MAIL_CC로_유지된다(self, df_customer, monkeypatch):
        """리팩터로 기본 CC가 바뀌면 거래명세표 참조자가 조용히 사라진다"""
        monkeypatch.setattr(mailer, 'TS_MAIL_CC', ('fixed@rotork.com',))
        r = find_recipient('220-81-21175', df_customer)
        assert 'fixed@rotork.com' in r.cc

    def test_해외_기본_CC는_OC_MAIL_CC를_쓴다(self):
        overseas = pd.DataFrame({
            'C-code by 해외': ['C-0095'],
            '고객명': ['WATERGATES GMBH'],
            '수신자 이메일': ['buyer@watergates.de'],
        })
        r = mailer.find_recipient_overseas('C-0095', overseas, fixed_cc=('oc@rotork.com',))
        assert r.cc == ('oc@rotork.com',)
