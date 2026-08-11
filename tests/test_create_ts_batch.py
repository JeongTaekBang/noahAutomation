"""
거래명세표 묶음 발송 테스트 (--date / --one-mail)
================================================

문서는 DN별 1장 그대로, 메일만 **거래처별 한 통**(첨부 여러 개)으로 묶는 경로다.
지켜야 할 것은 하나다 — **한 통에 남의 거래처가 섞이지 않는다.**

Excel COM을 타지 않도록 생성 단계(`_build_ts_from_dn`)와 메일 발송(`create_ts_mail`)은
가로챈다.
"""

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

import create_ts
from po_generator import mailer
from po_generator.mail_cli import MailMode, MailOptions
from po_generator.services.finder_service import OrderData
from po_generator.utils import normalize_biz_no

# 2026-08-06 씨앤케이엔지니어링 실측: DN 8건, DN마다 고객 발주번호가 다르다
CK_NAME = '씨앤케이엔지니어링(C&K ENG)'
CK_BIZ = '678-37-00167'
AUTO_NAME = '오토밸브'
AUTO_BIZ = '117-81-76942'


# === Fixtures ===

@pytest.fixture
def df_dn():
    """DN_국내 축소판 (한 DN이 여러 행 = 다중 아이템)"""
    rows = []
    for seq, po in enumerate(['26071402R0', '26071403R0', '26071404R0'], start=742):
        for line in (1, 2):
            rows.append({
                'DN_ID': f'DND-2026-0{seq}', 'Customer name': CK_NAME,
                'Business registration number': CK_BIZ, 'Customer PO': po,
                'Line item': line, '출고일': pd.Timestamp('2026-08-06'),
            })
    rows.append({
        'DN_ID': 'DND-2026-0749', 'Customer name': AUTO_NAME,
        'Business registration number': AUTO_BIZ, 'Customer PO': '10051',
        'Line item': 1, '출고일': pd.Timestamp('2026-08-06'),
    })
    rows.append({
        'DN_ID': 'DND-2026-0700', 'Customer name': CK_NAME,
        'Business registration number': CK_BIZ, 'Customer PO': '26060101R0',
        'Line item': 1, '출고일': pd.Timestamp('2026-07-02'),
    })
    return pd.DataFrame(rows)


@pytest.fixture
def df_customer():
    """Customer_국내 마스터 (load_customer_domestic 출력 형태)"""
    df = pd.DataFrame({
        '사업자번호': [CK_BIZ, AUTO_BIZ],
        '거래처명': [CK_NAME, AUTO_NAME],
        '이메일': ['ck@example.co.kr', 'auto@example.co.kr'],
    })
    df['_사업자번호_정규화'] = df['사업자번호'].map(normalize_biz_no)
    return df


def make_built(doc_id, customer, biz_no, po, tmp_path, date='2026-08-06'):
    """생성이 끝난 거래명세표 한 장 (Excel 없이)

    `date=None`은 선수금 거래명세표 — SO_국내 기반이라 출고일 컬럼 자체가 없다.
    """
    fields = {
        'Business registration number': biz_no,
        'Customer name': customer,
        'Customer PO': po,
    }
    if date:
        fields['출고일'] = pd.Timestamp(date)
    order = pd.Series(fields)
    path = tmp_path / f"TS_{doc_id}.xlsx"
    path.write_bytes(b'dummy')
    return create_ts.BuiltTS(
        doc_id=doc_id, output_file=path,
        order_data=order, items_df=pd.DataFrame([order]),
    )


def _single_item_order(customer, biz_no, po, date='2026-08-10'):
    """DN 한 줄 = 단일 아이템 (실제 조회 결과는 DataFrame이 아니라 Series다)"""
    return pd.Series({
        'Business registration number': biz_no,
        'Customer name': customer,
        'Customer PO': po,
        '출고일': pd.Timestamp(date),
    })


class _FakeService:
    """DocumentService 대역 — 조회는 주어진 행, 생성은 빈 파일 (Excel COM 없이)"""

    def __init__(self, rows: dict, out_dir: Path):
        self.rows = rows
        self.out_dir = out_dir
        self.finder = SimpleNamespace(
            find_dn=lambda dn_id: OrderData.from_result(self.rows[dn_id]))

    def generate_ts(self, doc_id, doc_type='DN'):
        path = self.out_dir / f'TS_{doc_id}.xlsx'
        path.write_bytes(b'dummy')
        return SimpleNamespace(success=True, output_file=path)


@pytest.fixture
def sent(monkeypatch):
    """create_ts_mail 호출 인자 수집 (메일은 나가지 않는다)"""
    calls = []

    def fake_mail(**kw):
        calls.append(kw)
        return mailer.MailResult(success=True, sent=False, recipient=kw['recipient'])

    monkeypatch.setattr(create_ts, 'create_ts_mail', fake_mail)
    return calls


@pytest.fixture
def ck_group(tmp_path):
    """씨앤케이 8건 (실측 그대로)"""
    pos = ['26071402R0', '26071403R0', '26071404R0', '26071405R0',
           '26071406R0', '26071407R0', '26071408R0', '26072302R0']
    ids = [f'DND-2026-07{n}' for n in (42, 43, 44, 45, 46, 47, 48, 50)]
    return [make_built(doc_id, CK_NAME, CK_BIZ, po, tmp_path)
            for doc_id, po in zip(ids, pos)]


# === 날짜 조회어 파싱 ===

class TestParseDispatchDate:
    @pytest.mark.parametrize('text', ['2026-08-06', '2026/08/06', '2026/8/6', '2026.08.06', '20260806'])
    def test_full_date_forms(self, text):
        assert create_ts.parse_dispatch_date(text) == pd.Timestamp('2026-08-06')

    @pytest.mark.parametrize('text', ['08-06', '8/6', '08.06'])
    def test_year_omitted_means_this_year(self, text):
        expected = pd.Timestamp(datetime(datetime.now().year, 8, 6))
        assert create_ts.parse_dispatch_date(text) == expected

    @pytest.mark.parametrize('text', ['', '   ', '어제', '2026-13-45', 'DND-2026-0742'])
    def test_unknown_forms_are_none(self, text):
        """DN 번호를 날짜 칸에 넣는 실수가 조용히 통과하면 안 된다"""
        assert create_ts.parse_dispatch_date(text) is None


# === 대상 선택 ===

class TestSelectDnIds:
    def test_picks_that_day_only(self, df_dn):
        ids = create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-06'))
        assert ids == ['DND-2026-0742', 'DND-2026-0743', 'DND-2026-0744', 'DND-2026-0749']

    def test_multi_item_dn_appears_once(self, df_dn):
        """한 DN이 여러 행이어도 문서는 한 장 — ID가 중복되면 같은 문서를 두 번 만든다"""
        ids = create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-06'))
        assert len(ids) == len(set(ids))

    def test_customer_filter_by_name(self, df_dn):
        ids = create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-06'), '씨앤케이')
        assert ids == ['DND-2026-0742', 'DND-2026-0743', 'DND-2026-0744']

    def test_customer_filter_by_biz_no(self, df_dn):
        ids = create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-06'), CK_BIZ)
        assert ids == ['DND-2026-0742', 'DND-2026-0743', 'DND-2026-0744']

    def test_other_days_excluded(self, df_dn):
        """같은 거래처라도 그날 나간 것만 — 지난달 출고분이 딸려가면 안 된다"""
        ids = create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-06'), '씨앤케이')
        assert 'DND-2026-0700' not in ids

    def test_no_match_returns_empty(self, df_dn, capsys):
        assert create_ts.select_dn_ids(df_dn, pd.Timestamp('2026-08-05')) == []
        assert '최근 출고일' in capsys.readouterr().out


# === 거래처 묶기 ===

class TestGroupByCustomer:
    def test_same_biz_no_is_one_group(self, ck_group):
        groups = create_ts.group_by_customer(ck_group)
        assert len(groups) == 1
        assert len(groups[0]) == 8

    def test_name_spelling_does_not_split(self, tmp_path):
        """'(주)' 유무로 갈리면 같은 거래처에 메일이 두 통 간다"""
        built = [
            make_built('DND-1', '씨앤케이엔지니어링', CK_BIZ, 'PO-1', tmp_path),
            make_built('DND-2', '(주)씨앤케이엔지니어링 ', '678-3700167', 'PO-2', tmp_path),
        ]
        assert len(create_ts.group_by_customer(built)) == 1

    def test_different_customers_never_merge(self, tmp_path):
        built = [
            make_built('DND-1', CK_NAME, CK_BIZ, 'PO-1', tmp_path),
            make_built('DND-2', AUTO_NAME, AUTO_BIZ, '10051', tmp_path),
        ]
        groups = create_ts.group_by_customer(built)
        assert [len(g) for g in groups] == [1, 1]

    def test_missing_biz_no_never_merges(self, tmp_path):
        """사업자번호가 비었다고 서로 묶으면 남의 명세표가 첨부된다"""
        built = [
            make_built('DND-1', '알수없음A', '', 'PO-1', tmp_path),
            make_built('DND-2', '알수없음B', None, 'PO-2', tmp_path),
        ]
        assert len(create_ts.group_by_customer(built)) == 2

    def test_first_seen_order_preserved(self, tmp_path):
        built = [
            make_built('DND-1', AUTO_NAME, AUTO_BIZ, '10051', tmp_path),
            make_built('DND-2', CK_NAME, CK_BIZ, 'PO-1', tmp_path),
            make_built('DND-3', AUTO_NAME, AUTO_BIZ, '10086', tmp_path),
        ]
        groups = create_ts.group_by_customer(built)
        assert [g[0].doc_id for g in groups] == ['DND-1', 'DND-2']
        assert [one.doc_id for one in groups[0]] == ['DND-1', 'DND-3']


class TestGroupLabels:
    def test_single_keeps_own_id(self):
        assert create_ts._group_doc_label(['DND-2026-0742']) == 'DND-2026-0742'

    def test_many_shows_count(self):
        label = create_ts._group_doc_label(['DND-2026-0742', 'DND-2026-0743'])
        assert label == 'DND-2026-0742 외 1건'

    def test_group_date_is_latest(self, tmp_path):
        built = [
            make_built('DND-1', CK_NAME, CK_BIZ, 'PO-1', tmp_path, date='2026-08-05'),
            make_built('DND-2', CK_NAME, CK_BIZ, 'PO-2', tmp_path, date='2026-08-06'),
        ]
        assert create_ts._group_date_str(built) == '2026-08-06'

    def test_group_date_none_when_no_dispatch(self, tmp_path):
        """선수금 거래명세표는 출고일 자체가 없다 — 호출부가 오늘로 폴백한다"""
        built = [make_built('ADV-1', CK_NAME, CK_BIZ, 'PO-1', tmp_path, date=None)]
        assert create_ts._group_date_str(built) is None


# === 묶음 실행 (생성 N장 → 메일 1통) ===

class TestGenerateTsBatch:
    def _patch_builder(self, monkeypatch, built):
        registry = {one.doc_id: one for one in built}
        monkeypatch.setattr(
            create_ts, '_build_ts_from_dn',
            lambda doc_id, service: registry.get(doc_id),
        )
        return registry

    def test_one_mail_with_all_attachments(self, ck_group, df_customer, sent, monkeypatch):
        self._patch_builder(monkeypatch, ck_group)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        ok = create_ts.generate_ts_batch([one.doc_id for one in ck_group], opts)

        assert ok is True
        assert len(sent) == 1                          # DN 8건 → 메일 한 통
        assert len(sent[0]['xlsx_path']) == 8          # 첨부는 8장 전부
        assert sent[0]['doc_id'] == 'DND-2026-0742 외 7건'
        assert sent[0]['date_str'] == '2026-08-06'

    def test_all_customer_pos_in_one_mail(self, ck_group, df_customer, sent, monkeypatch):
        """발주번호가 첫 건만 실리면 고객은 나머지 7건을 못 찾는다"""
        self._patch_builder(monkeypatch, ck_group)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        create_ts.generate_ts_batch([one.doc_id for one in ck_group], opts)

        customer_po = sent[0]['customer_po']
        assert '26071402R0' in customer_po
        assert '26072302R0' in customer_po
        assert customer_po.count(',') == 7

    def test_customers_never_mixed(self, ck_group, df_customer, sent, monkeypatch, tmp_path):
        """한 통에 남의 거래처가 섞이면 명세표가 유출된다"""
        auto = make_built('DND-2026-0749', AUTO_NAME, AUTO_BIZ, '10051', tmp_path)
        self._patch_builder(monkeypatch, [*ck_group, auto])
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        create_ts.generate_ts_batch(
            [one.doc_id for one in ck_group] + ['DND-2026-0749'], opts)

        assert len(sent) == 2
        ck_mail, auto_mail = sent
        assert ck_mail['recipient'].to == ('ck@example.co.kr',)
        assert auto_mail['recipient'].to == ('auto@example.co.kr',)
        assert len(auto_mail['xlsx_path']) == 1
        assert '10051' not in ck_mail['customer_po']
        assert '26071402R0' not in auto_mail['customer_po']

    def test_asks_once_per_customer(self, ck_group, df_customer, sent, monkeypatch, tmp_path):
        """확인은 문서마다가 아니라 메일마다 — 8번 묻던 것이 1번이 된다"""
        auto = make_built('DND-2026-0749', AUTO_NAME, AUTO_BIZ, '10051', tmp_path)
        self._patch_builder(monkeypatch, [*ck_group, auto])
        asked = []
        monkeypatch.setattr('builtins.input', lambda prompt: asked.append(prompt) or 'y')
        opts = MailOptions(mode=MailMode.ASK, df_customer=df_customer)

        create_ts.generate_ts_batch(
            [one.doc_id for one in ck_group] + ['DND-2026-0749'], opts)

        assert len(asked) == 2
        assert len(sent) == 2

    def test_no_mail_mode_still_generates(self, ck_group, monkeypatch):
        self._patch_builder(monkeypatch, ck_group)
        monkeypatch.setattr(
            create_ts, 'create_ts_mail',
            lambda **kw: pytest.fail('--no-mail에서 메일이 나가면 안 된다'),
        )
        assert create_ts.generate_ts_batch(
            [one.doc_id for one in ck_group], MailOptions.disabled()) is True

    def test_failed_doc_does_not_block_the_rest(self, ck_group, df_customer, sent, monkeypatch):
        """한 건이 실패해도 나머지는 나간다 (반환값만 False)"""
        self._patch_builder(monkeypatch, ck_group)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        ok = create_ts.generate_ts_batch(
            [one.doc_id for one in ck_group] + ['DND-9999-9999'], opts)

        assert ok is False
        assert len(sent) == 1
        assert len(sent[0]['xlsx_path']) == 8

    def test_all_failed_sends_nothing(self, df_customer, sent, monkeypatch):
        monkeypatch.setattr(create_ts, '_build_ts_from_dn', lambda doc_id, service: None)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        assert create_ts.generate_ts_batch(['DND-9999-9999'], opts) is False
        assert sent == []

    def test_mixed_customer_doc_excluded_but_rest_sent(
            self, ck_group, df_customer, sent, monkeypatch, tmp_path):
        """한 DN에 두 거래처가 들어간 문서는 첨부에서 뺀다 (실측: DND-2026-0748)

        DN 번호를 재사용한 시트 오류인데, 그대로 붙이면 남의 품목·단가가 고객에게 나간다.
        나머지 7장은 그대로 나가야 한다 — 한 건 때문에 하루치를 못 보내면 안 되니까.
        """
        mixed = make_built('DND-2026-0748', CK_NAME, CK_BIZ, '26071408R0', tmp_path)
        foreign_row = mixed.items_df.iloc[0].copy()
        foreign_row['Business registration number'] = AUTO_BIZ
        foreign_row['Customer PO'] = '10042'
        mixed = create_ts.BuiltTS(
            doc_id=mixed.doc_id, output_file=mixed.output_file,
            order_data=mixed.order_data,
            items_df=pd.concat([mixed.items_df, pd.DataFrame([foreign_row])], ignore_index=True),
        )
        clean = [one for one in ck_group if one.doc_id != mixed.doc_id]
        self._patch_builder(monkeypatch, [*clean, mixed])
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        create_ts.generate_ts_batch(
            [one.doc_id for one in clean] + [mixed.doc_id], opts)

        assert len(sent) == 1
        assert len(sent[0]['xlsx_path']) == len(clean) == 7
        assert '10042' not in sent[0]['customer_po']
        assert '26071408R0' not in sent[0]['customer_po']   # 섞인 문서의 발주번호도 빠진다

    def test_single_item_docs_still_mail(self, df_customer, sent, monkeypatch, tmp_path):
        """단일 아이템 DN만 모여도 메일이 나간다 (2026-08-10 실측 크래시)

        `OrderData.items_df`는 **단일 아이템이면 None**이다. 그것을 그대로 BuiltTS에
        담았더니 묶음의 `pd.concat`이 'All objects passed were None'으로 죽었다.
        여기 fixture들이 늘 1행짜리 DataFrame을 넘겨서 못 잡던 자리다 — 그래서 이
        테스트만은 **진짜 `_build_ts_from_dn`을 지나간다** (조회·생성만 대역).
        """
        rows = {
            'DND-2026-0756': _single_item_order(AUTO_NAME, AUTO_BIZ, '10051'),
            'DND-2026-0757': _single_item_order(AUTO_NAME, AUTO_BIZ, '10086'),
        }
        monkeypatch.setattr(
            create_ts, 'DocumentService', lambda: _FakeService(rows, tmp_path))
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        ok = create_ts.generate_ts_batch(list(rows), opts)

        assert ok is True
        assert len(sent) == 1
        assert len(sent[0]['xlsx_path']) == 2
        # 발주번호도 둘 다 실린다 (None을 흘려보내면 concat이 조용히 건너뛴다)
        assert sent[0]['customer_po'] == '10051, 10086'

    def test_single_item_builder_fills_items_df(self, monkeypatch, tmp_path):
        """BuiltTS.items_df는 단일 아이템이어도 None이 아니다 (묶음의 전제)"""
        rows = {'DND-2026-0756': _single_item_order(AUTO_NAME, AUTO_BIZ, '10051')}
        built = create_ts._build_ts_from_dn(
            'DND-2026-0756', _FakeService(rows, tmp_path))

        assert built is not None
        assert len(built.items_df) == 1

    def test_advance_id_uses_advance_builder(self, df_customer, sent, monkeypatch, tmp_path):
        """선수금도 같은 거래처면 한 통에 담긴다 (조용히 빠지지 않는다)"""
        dn = make_built('DND-2026-0742', CK_NAME, CK_BIZ, 'PO-1', tmp_path)
        adv = make_built('ADV_2026-0001', CK_NAME, CK_BIZ, 'PO-2', tmp_path, date=None)
        monkeypatch.setattr(create_ts, '_build_ts_from_dn', lambda doc_id, service: dn)
        monkeypatch.setattr(create_ts, '_build_ts_from_adv', lambda doc_id, service: adv)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        create_ts.generate_ts_batch(['DND-2026-0742', 'ADV_2026-0001'], opts)

        assert len(sent) == 1
        assert len(sent[0]['xlsx_path']) == 2


# === 단건 경로 (묶음과 같은 생성 코드를 쓰는지) ===

class TestSingleDocPath:
    """`--one-mail` 없이 ID 하나를 주는 기존 사용법은 그대로여야 한다"""

    def test_one_doc_one_mail(self, df_customer, sent, monkeypatch, tmp_path):
        built = make_built('DND-2026-0742', CK_NAME, CK_BIZ, '26071402R0', tmp_path)
        monkeypatch.setattr(create_ts, '_build_ts_from_dn', lambda doc_id, service: built)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        assert create_ts.generate_ts_from_dn('DND-2026-0742', pd.DataFrame(), opts) is True
        assert len(sent) == 1
        assert len(sent[0]['xlsx_path']) == 1          # 단건은 첨부도 한 장
        assert sent[0]['doc_id'] == 'DND-2026-0742'    # '외 N건'이 붙지 않는다

    def test_generation_failure_reports_false(self, df_customer, sent, monkeypatch):
        monkeypatch.setattr(create_ts, '_build_ts_from_dn', lambda doc_id, service: None)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        assert create_ts.generate_ts_from_dn('DND-9999-9999', pd.DataFrame(), opts) is False
        assert sent == []

    def test_advance_path_uses_advance_builder(self, df_customer, sent, monkeypatch, tmp_path):
        built = make_built('ADV_2026-0001', CK_NAME, CK_BIZ, 'PO-1', tmp_path, date=None)
        monkeypatch.setattr(create_ts, '_build_ts_from_adv', lambda doc_id, service: built)
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        assert create_ts.generate_ts_from_adv('ADV_2026-0001', opts) is True
        assert len(sent) == 1


# === 거래처 섞임 관문 (단건·월합·묶음 공통) ===

class TestForeignCustomerGuard:
    """한 문서에 두 거래처가 실리면 메일로 나가지 않는다

    실측(2026-08-06 `DND-2026-0748`, 2026-04-14 `DND-2026-0328`): DN 번호를 재사용해
    한 DN에 두 SO가 들어간 시트 오류가 있다. 시트를 고치는 건 사람 몫이지만,
    고치기 전에 남의 품목·단가가 첨부돼 나가는 것은 코드가 막는다.
    """

    def _order(self, biz_no=CK_BIZ):
        return pd.Series({
            'Business registration number': biz_no,
            'Customer name': CK_NAME,
            'Customer PO': '26071408R0',
        })

    def _items(self, *biz_numbers):
        return pd.DataFrame([
            {'Business registration number': biz, 'Customer PO': f'PO-{i}'}
            for i, biz in enumerate(biz_numbers)
        ])

    def test_clean_document_has_none(self):
        assert create_ts.foreign_biz_numbers(
            self._order(), self._items(CK_BIZ, CK_BIZ)) == []

    def test_mixed_document_lists_the_intruder(self):
        found = create_ts.foreign_biz_numbers(self._order(), self._items(CK_BIZ, AUTO_BIZ))
        assert found == [normalize_biz_no(AUTO_BIZ)]

    def test_blank_biz_no_is_not_an_intruder(self):
        """빈 칸은 '다른 거래처'가 아니라 '미입력'이다 — 여기서 막으면 정상 건이 안 나간다"""
        assert create_ts.foreign_biz_numbers(self._order(), self._items(CK_BIZ, '')) == []

    def test_missing_column_does_not_block(self):
        items = pd.DataFrame([{'Customer PO': 'PO-1'}])
        assert create_ts.foreign_biz_numbers(self._order(), items) == []

    def test_warning_names_the_intruder(self):
        """번호만 찍으면 시트에서 어느 줄인지 못 찾는다"""
        items = pd.DataFrame([
            {'Business registration number': CK_BIZ, 'Customer name': CK_NAME},
            {'Business registration number': AUTO_BIZ, 'Customer name': AUTO_NAME},
        ])
        found = create_ts.foreign_biz_numbers(self._order(), items)
        assert create_ts.describe_foreign(items, found) == \
            f'{AUTO_NAME}({normalize_biz_no(AUTO_BIZ)})'

    def test_single_doc_mail_is_blocked(self, df_customer, monkeypatch, tmp_path):
        """단건 경로도 같은 관문을 지난다 (묶음에만 있으면 예전 경로로 새 나간다)"""
        monkeypatch.setattr(
            create_ts, 'create_ts_mail',
            lambda **kw: pytest.fail('거래처가 섞인 문서는 나가면 안 된다'),
        )
        opts = MailOptions(mode=MailMode.DRAFT, df_customer=df_customer)

        blocked = create_ts._mail_ts(
            self._order(), tmp_path / 'TS.xlsx', 'DND-2026-0748', opts,
            items_df=self._items(CK_BIZ, AUTO_BIZ),
        )
        assert blocked is False


# === 인자 조합 검사 ===

class TestValidateSelectionArgs:
    def _args(self, **kw):
        import argparse
        base = dict(merge=False, one_mail=False, date=None, customer=None, doc_ids=[])
        base.update(kw)
        return argparse.Namespace(**base)

    def test_plain_ids_are_fine(self):
        assert create_ts.validate_selection_args(self._args(doc_ids=['DND-1'])) is None

    def test_date_alone_is_fine(self):
        assert create_ts.validate_selection_args(self._args(date='2026-08-06')) is None

    def test_date_with_customer_is_fine(self):
        args = self._args(date='2026-08-06', customer='씨앤케이')
        assert create_ts.validate_selection_args(args) is None

    def test_merge_and_one_mail_conflict(self):
        """'문서를 합침'과 '메일만 합침'은 뜻이 정반대다"""
        args = self._args(merge=True, one_mail=True, doc_ids=['DND-1', 'DND-2'])
        assert '--merge' in (create_ts.validate_selection_args(args) or '')

    def test_merge_and_date_conflict(self):
        args = self._args(merge=True, date='2026-08-06')
        assert create_ts.validate_selection_args(args) is not None

    def test_customer_without_date_blocked(self):
        """날짜 없이 거래처만 주면 과거 DN 전부가 대상이 된다"""
        args = self._args(customer='씨앤케이')
        assert '--customer' in (create_ts.validate_selection_args(args) or '')

    def test_date_with_ids_blocked(self):
        args = self._args(date='2026-08-06', doc_ids=['DND-1'])
        assert create_ts.validate_selection_args(args) is not None


# === CLI 배선 ===

class TestCliWiring:
    def test_one_mail_routes_to_batch(self, monkeypatch, tmp_path):
        called = []
        monkeypatch.setattr(create_ts, 'load_dn_data', lambda: pd.DataFrame({'DN_ID': ['DND-1']}))
        monkeypatch.setattr(create_ts, 'load_pmt_data', lambda: pd.DataFrame({'선수금_ID': []}))
        monkeypatch.setattr(
            create_ts, 'generate_ts_batch',
            lambda doc_ids, opts: called.append(list(doc_ids)) or True,
        )
        monkeypatch.setattr(
            'sys.argv', ['create_ts.py', 'DND-1', 'DND-2', '--one-mail', '--no-mail'])

        assert create_ts.main() == 0
        assert called == [['DND-1', 'DND-2']]

    def test_duplicate_ids_removed(self, monkeypatch):
        """같은 DN을 두 번 주면 같은 문서가 두 번 첨부된다"""
        called = []
        monkeypatch.setattr(create_ts, 'load_dn_data', lambda: pd.DataFrame({'DN_ID': ['DND-1']}))
        monkeypatch.setattr(create_ts, 'load_pmt_data', lambda: pd.DataFrame({'선수금_ID': []}))
        monkeypatch.setattr(
            create_ts, 'generate_ts_batch',
            lambda doc_ids, opts: called.append(list(doc_ids)) or True,
        )
        monkeypatch.setattr(
            'sys.argv', ['create_ts.py', 'DND-1', 'DND-1', '--one-mail', '--no-mail'])

        create_ts.main()
        assert called == [['DND-1']]

    def test_date_selects_and_batches(self, df_dn, monkeypatch):
        called = []
        monkeypatch.setattr(create_ts, 'load_dn_data', lambda: df_dn)
        monkeypatch.setattr(create_ts, 'load_pmt_data', lambda: pd.DataFrame({'선수금_ID': []}))
        monkeypatch.setattr(
            create_ts, 'generate_ts_batch',
            lambda doc_ids, opts: called.append(list(doc_ids)) or True,
        )
        monkeypatch.setattr(
            'sys.argv',
            ['create_ts.py', '--date', '2026-08-06', '--customer', '씨앤케이', '--no-mail'])

        assert create_ts.main() == 0
        assert called == [['DND-2026-0742', 'DND-2026-0743', 'DND-2026-0744']]

    def test_bad_date_stops(self, df_dn, monkeypatch, capsys):
        monkeypatch.setattr(create_ts, 'load_dn_data', lambda: df_dn)
        monkeypatch.setattr(create_ts, 'load_pmt_data', lambda: pd.DataFrame({'선수금_ID': []}))
        monkeypatch.setattr(
            create_ts, 'generate_ts_batch',
            lambda doc_ids, opts: pytest.fail('날짜를 못 읽으면 생성하면 안 된다'),
        )
        monkeypatch.setattr('sys.argv', ['create_ts.py', '--date', '어제', '--no-mail'])

        assert create_ts.main() == 1
        assert '날짜 형식' in capsys.readouterr().out

    def test_conflicting_flags_stop_before_loading(self, monkeypatch, capsys):
        monkeypatch.setattr(
            create_ts, 'load_dn_data',
            lambda: pytest.fail('인자가 틀렸으면 파일도 읽지 않는다'),
        )
        monkeypatch.setattr(
            'sys.argv', ['create_ts.py', 'DND-1', 'DND-2', '--merge', '--one-mail'])

        assert create_ts.main() == 1
        assert '--merge' in capsys.readouterr().out
