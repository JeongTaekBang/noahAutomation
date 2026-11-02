"""
AX 오더북 — IR + Item transaction 엔진 (po_generator/ax_orderbook.py, build_ax_orderbook.py) 테스트

한 프로젝트가 규칙 하나를 맡는 가짜 회계 오더북(P01 Start)·Item requirement·Item transaction으로 2026 P01~P03을
만들어 본다. 실측 근거는 tasks/todo.md 'AX 오더북' 절.
"""

from __future__ import annotations

from datetime import date

import openpyxl
import pandas as pd
import pytest

import build_ax_orderbook as cli
from po_generator import ax_orderbook as ax


def official_row(project: str, cc: str, dept: str, period: str = 'P01', qty: float = 0, value: float = 0,
                 account: str = 'C00001') -> dict:
    """회계 Raw 한 행 — 기초(Start)만 있고 그 달 흐름은 없다(Start = End)"""
    row = {c: None for c in ax.RAW_COLUMNS}
    row.update({
        'Project': project, 'Cost center': cc, 'Prod/Dept (Revenue)': dept, 'Period': period,
        'Account number': account, 'Customer Name': f'고객{account}', 'Market': 'CP&I', 'Industry': '70',
        'Order Book Code': 'IQ3', 'Account Manager': '10523', 'Name': f'{project} name', 'Location': '4080',
        'Project creation date': pd.Timestamp('2025-10-01'), 'Contract Engineer': '12594',
        'Contract Engineer Name': 'CE Kim',
    })
    for c in ax.NUM_COLUMNS:
        row[c] = 0.0
    row.update({'Qty Start': qty, 'Value Start': value, 'Qty End': qty, 'Value End': value})
    return row


def official_df(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df['Key'] = df['Project'] + df['Cost center'] + df['Prod/Dept (Revenue)']
    df['pn'] = df['Period'].str[1:].astype(int)
    return df


def ir_line(project: str, line, material: str, status: str, qty: float, unit_price: float, created_utc: str,
            ship: str | None = None, remainder: float | None = None, customer: str = 'C00001',
            division: str = '10', vat: str = '01') -> dict:
    net = 0.0 if status == 'Cancelled' else qty * unit_price
    return {
        'Project': project, 'Line number': str(line), 'Created date': pd.Timestamp(created_utc),
        'Customer': customer, 'Customer name': f'고객{customer}', 'Project group': 'ProdSale',
        'Item number': 'ITEM', 'Item name': 'item', 'Quantity': qty, 'Unit price': unit_price,
        'Net amount': net, 'Line status': status, 'Actual ship date': pd.Timestamp(ship) if ship else pd.NaT,
        'Deliver remainder': remainder, 'Division': division, 'Prod/Dept (Material)': material,
        'Contract engineer': 'CE Kim', 'VAT group': vat,
    }


def it_row(project: str, line, when: str, qty: float, net: float, status: str = 'Invoiced',
           currency: str = 'KRW', prop: str = 'Invoiceable') -> dict:
    return {'Project date': pd.Timestamp(when), 'Project': project, 'Line number': float(line),
            'Item number': 'ITEM', 'Item name': 'item', 'Quantity': qty, 'Net amount': net,
            'Sales currency': currency, 'Line property': prop, 'Transaction status': status}


OFFICIAL = [
    official_row('82-00001.001', '10', '0112', qty=2, value=2_000_000),
    official_row('82-00002.001', '50', '0849', qty=6, value=600_000),
    official_row('82-00003.001', '10', '0112', qty=1, value=500_000),
    official_row('82-00004.001', '10', '0134', qty=1, value=1_400_000, account='C00033'),
    official_row('82-00007.001', '10', '0108', qty=2, value=1_000_000),
    official_row('82-00010.001', '10', '0112', qty=1, value=50_000),
    official_row('82-00011.001', '10', '0112', qty=1, value=2_250_000, account='C00014'),
]
IR_LINES = [
    # A: 12월 수주·2월 청구
    ir_line('82-00001.001', 1, '1012', 'Invoiced', 2, 1_000_000, '2025-12-10', ship='2026-02-10'),
    # B: 분할출고 — 1월 2개 + 2월 4개 (IR엔 마지막 출고일만 남는다)
    ir_line('82-00002.001', 1, '1849', 'Invoiced', 6, 100_000, '2025-11-01', ship='2026-02-15'),
    # C: 나중에 취소 — IR엔 금액 0, 회계 기초엔 있다
    ir_line('82-00003.001', 1, '1012', 'Cancelled', 1, 500_000, '2025-12-01'),
    # D: USD — 수주 12월(1,400) · 청구 2월(1,450)
    ir_line('82-00004.001', 1, '1034', 'Invoiced', 1, 1_000, '2025-12-15', ship='2026-02-05', customer='C00033',
            vat='ZR'),
    # E: 1월 수주(UTC 1/9 16:00 = KST 1/10)·1월 출고, 청구 전(Posted)
    ir_line('82-00005.001', 1, '1007', 'Invoiced', 1, 300_000, '2026-01-09 16:00', ship='2026-01-28'),
    # F: 2025년 출고 후 크레딧 노트 — Posted +2 / Non-Invoiceable Posted −2 → 상쇄
    ir_line('82-00006.001', 1, '1012', 'Invoiced', 2, 500_000, '2025-10-01', ship='2025-11-01'),
    # G: 크레인(3060, 배운 짝·규칙 없음) + 인건비(1008 → 0108) — 회계는 둘 다 0108
    ir_line('82-00007.001', 1, '3060', 'Open order', 1, 700_000, '2025-12-01', remainder=1),
    ir_line('82-00007.001', 2, '1008', 'Open order', 1, 300_000, '2025-12-01', remainder=1),
    # H: 라인번호 재사용 — 1번이 1월(1012, 10만)과 2월(1849, 25만) 두 라인
    ir_line('82-00008.001', 1, '1012', 'Invoiced', 1, 100_000, '2026-01-05', ship='2026-03-03'),
    ir_line('82-00008.001', 1, '1849', 'Invoiced', 1, 250_000, '2026-02-01', ship='2026-02-20'),
    # J: EUR — 수주월(2025-11) 환율이 없다
    ir_line('82-00011.001', 1, '1012', 'Open order', 1, 1_500, '2025-11-15', remainder=1, customer='C00014',
            vat='ZR'),
    # K: 수주 100만인데 90만 청구
    ir_line('82-00012.001', 1, '1012', 'Invoiced', 1, 1_000_000, '2026-01-03', ship='2026-01-15'),
    # L: UTC 1/31 16:00 = KST 2/1 → 2월 수주
    ir_line('82-00013.001', 1, '1012', 'Open order', 1, 50_000, '2026-01-31 16:00', remainder=1),
]
IT_ROWS = [
    it_row('82-00001.001', 1, '2026-02-10', 2, 2_000_000),
    it_row('82-00002.001', 1, '2026-01-20', 2, 200_000),
    it_row('82-00002.001', 1, '2026-02-15', 4, 400_000),
    it_row('82-00004.001', 1, '2026-02-05', 1, 1_000, currency='USD'),
    it_row('82-00005.001', 1, '2026-01-28', 1, 300_000, status='Posted'),
    it_row('82-00006.001', 1, '2025-11-01', 2, 1_000_000, status='Posted'),
    it_row('82-00006.001', 1, '2025-12-01', -2, -1_000_000, status='Posted', prop='Non-Invoiceable'),
    it_row('82-00008.001', 1, '2026-02-20', 1, 250_000),
    it_row('82-00008.001', 1, '2026-03-03', 1, 100_000),
    it_row('82-00012.001', 1, '2026-01-15', 1, 900_000),
    it_row('82-00012.001', 1, '2026-01-20', 1, 900_000, status='Adjusted'),
    it_row('82-00012.001', 1, '2026-01-20', -1, -900_000, status='Adjusted'),
    it_row('82-00010.001', 0, '2026-01-10', 1, 50_000),          # IR 라인 없는 출고
]


def make_fx() -> ax.FxBook:
    fx = ax.FxBook()
    for month, rate in (('2025-12', 1_400), ('2026-01', 1_420), ('2026-02', 1_450), ('2026-03', 1_460)):
        fx.put('USD', month, rate, 'test')
    fx.put('EUR', '2026-01', 1_500, 'test')
    return fx


@pytest.fixture
def scenario():
    official = official_df(OFFICIAL)
    ir = ax.prepare_ir(pd.DataFrame(IR_LINES))
    it = ax.prepare_it(pd.DataFrame(IT_ROWS))
    res = ax.build_orderbook(official, ir, it, 2026, 3, make_fx())
    return official, ir, it, res


def k(res: ax.OrderbookResult, key: str, month: int) -> pd.Series:
    t = res.keys[(res.keys['Key'] == key) & (res.keys['pn'] == month)]
    assert len(t) == 1, f"{key} P{month:02d} 없음 — 있는 키: {sorted(set(res.keys['Key']))}"
    return t.iloc[0]


class TestRollForward:
    def test_identity_holds_for_every_key_and_month(self, scenario):
        t = scenario[3].keys
        for q in ('q', 'v'):
            resid = t[f'S{q}'] + t[f'I{q}'] - t[f'O{q}'] + t[f'V{q}'] - t[f'E{q}']
            assert resid.abs().max() < 1e-6

    def test_start_is_official_p01_then_our_previous_end(self, scenario):
        official, _, _, res = scenario
        t = res.keys
        assert t[t['pn'] == 1]['Sv'].sum() == pytest.approx(official['Value Start'].sum())
        for m in (2, 3):
            prev = t[t['pn'] == m - 1].set_index('Key')['Ev']
            cur = t[t['pn'] == m].set_index('Key')['Sv']
            idx = prev.index.union(cur.index)
            assert (prev.reindex(idx).fillna(0) - cur.reindex(idx).fillna(0)).abs().max() < 1e-6

    def test_billed_shipment_is_output_in_ship_month(self, scenario):
        res = scenario[3]
        jan, feb = k(res, '82-00001.001100112', 1), k(res, '82-00001.001100112', 2)
        assert (jan['Ov'], jan['Ev']) == (0, 2_000_000)
        assert (feb['Oq'], feb['Ov'], feb['Ev'], feb['Vv']) == (2, 2_000_000, 0, 0)

    def test_split_shipment_is_exact_per_month(self, scenario):
        """IR엔 마지막 출고일(2/15)만 있지만 IT는 출고마다 한 행 — 1월 2개, 2월 4개"""
        res = scenario[3]
        jan, feb = k(res, '82-00002.001500849', 1), k(res, '82-00002.001500849', 2)
        assert (jan['Oq'], jan['Ov'], jan['Eq'], jan['Ev'], jan['Vv']) == (2, 200_000, 4, 400_000, 0)
        assert (feb['Oq'], feb['Ov'], feb['Ev'], feb['Vv']) == (4, 400_000, 0, 0)

    def test_unbilled_shipment_stays_in_end(self, scenario):
        """출고됐지만 청구 전(Posted) = Backlog 잔류 — Output 아님"""
        res = scenario[3]
        for m in (1, 2, 3):
            row = k(res, '82-00005.001100107', m)
            assert (row['Ov'], row['Ev']) == (0, 300_000)
        assert k(res, '82-00005.001100107', 1)['Iv'] == 300_000

    def test_credit_note_pair_cancels_out(self, scenario):
        """+2 Posted와 −2 Non-Invoiceable Posted는 라인 합계로 0 — 청구 대기가 아니다"""
        res = scenario[3]
        assert '82-00006.001' not in set(res.keys['Project'])
        assert '82-00006.001' not in set(res.pending['Project'])

    def test_cancelled_line_is_excluded_from_start(self, scenario):
        """회계 기초엔 있지만 IR에선 취소 → P01 Variance로 빠진다(기초 차이)"""
        row = k(scenario[3], '82-00003.001100112', 1)
        assert (row['Sv'], row['Ev'], row['Vv']) == (500_000, 0, -500_000)

    def test_foreign_output_at_bill_month_rate_end_at_order_month_rate(self, scenario):
        res = scenario[3]
        jan, feb = k(res, '82-00004.001100134', 1), k(res, '82-00004.001100134', 2)
        assert jan['Ev'] == pytest.approx(1_000 * 1_400)
        assert feb['Ov'] == pytest.approx(1_000 * 1_450)
        assert feb['Vv'] == pytest.approx(50_000)             # 환율 재평가

    def test_billed_amount_differs_from_order_goes_to_variance(self, scenario):
        row = k(scenario[3], '82-00012.001100112', 1)
        assert (row['Iv'], row['Ov'], row['Ev'], row['Vv']) == (1_000_000, 900_000, 0, -100_000)

    def test_shipment_without_ir_line_goes_to_project_key(self, scenario):
        row = k(scenario[3], '82-00010.001100112', 1)
        assert (row['Sv'], row['Ov'], row['Ev']) == (50_000, 50_000, 0)


class TestLinesAndKeys:
    def test_kst_decides_the_month(self, scenario):
        res = scenario[3]
        assert k(res, '82-00005.001100107', 1)['Iv'] == 300_000                 # UTC 1/9 16:00 = KST 1/10
        assert '82-00013.001100112' not in set(res.keys.loc[res.keys['pn'] == 1, 'Key'])
        assert k(res, '82-00013.001100112', 2)['Iv'] == 50_000                  # UTC 1/31 16:00 = KST 2/1

    def test_line_with_unmapped_dept_is_kept_as_key_split(self, scenario):
        """매핑이 다른 라인도 빠지지 않는다 — 프로젝트 합계는 회계와 같고 키만 갈린다"""
        res = scenario[3]
        p = res.keys[(res.keys['Project'] == '82-00007.001') & (res.keys['pn'] == 1)]
        assert p['Ev'].sum() == pytest.approx(1_000_000)
        assert k(res, '82-00007.001100108', 1)['Ev'] == 300_000
        assert k(res, '82-00007.00110????', 1)['Ev'] == 700_000

    def test_reused_line_number_links_by_unit_price(self, scenario):
        """1번 라인 두 개 — 3/3 출고(10만)는 단가가 맞는 1월 라인(0112)에 붙어야 한다"""
        res = scenario[3]
        mar = k(res, '82-00008.001100112', 3)
        assert (mar['Oq'], mar['Ov'], mar['Ev']) == (1, 100_000, 0)
        feb = k(res, '82-00008.001100849', 2)
        assert (feb['Iv'], feb['Ov'], feb['Ev']) == (250_000, 250_000, 0)
        assert res.keys[(res.keys['Key'] == '82-00008.001100849') & (res.keys['pn'] == 3)].empty

    def test_adjusted_rows_are_dropped(self, scenario):
        it = scenario[2]
        assert not it['Transaction status'].eq('Adjusted').any()

    def test_line_numbers_are_normalized(self):
        assert [ax.normalize_line_number(v) for v in ('1.001', 1.001, '14', 14.0, None)] == ['1.001', '1.001', '14', '14', None]


class TestOutputs:
    def test_pending_list(self, scenario):
        p = scenario[3].pending
        assert list(p['Project']) == ['82-00005.001']
        assert (p['금액(원)'].iloc[0], p['경과 개월'].iloc[0]) == (300_000, 2)

    def test_raw_rows_have_all_months_and_attributes(self, scenario):
        raw = scenario[3].raw_rows
        assert set(raw['Period']) == {'P01', 'P02', 'P03'}
        row = raw[(raw['_key'] == '82-00001.001100112') & (raw['pn'] == 1)].iloc[0]
        assert (row['Market'], row['Order Book Code'], row['Report Start']) == ('CP&I', 'IQ3', pd.Timestamp('2026-01-01'))
        for c in ax.FORMULA_COLUMNS:
            assert raw[c].isna().all(), c

    def test_review(self, scenario):
        rv = scenario[3].review
        kinds = set(rv['구분'])
        assert {'신규 키 — 속성 입력 필요', '부서 매핑 없음', '부서 규칙으로 추정', '수주·청구 금액 차이',
                'IR 라인 없는 IT 출고', '환율 없음 — FX 시트에 추가'} <= kinds
        diff = rv[rv['구분'] == '수주·청구 금액 차이']
        assert (diff['Project'].iloc[0], diff['금액'].iloc[0]) == ('82-00012.001', -100_000)
        assert 'EUR 2025-11' in rv.loc[rv['구분'] == '환율 없음 — FX 시트에 추가', '내용'].iloc[0]
        assert '통화 확인' not in kinds

    def test_same_input_same_result(self, scenario):
        official, ir, it, res = scenario
        again = ax.build_orderbook(official, ir, it, 2026, 3, make_fx())
        pd.testing.assert_frame_equal(res.raw_rows, again.raw_rows)


class TestCompareAndClassify:
    @staticmethod
    def as_official(res: ax.OrderbookResult) -> pd.DataFrame:
        off = res.raw_rows.copy()
        off['Key'] = off['_key']
        return off

    def test_matching_official_has_no_differences(self, scenario):
        _, ir, _, res = scenario
        off = self.as_official(res)
        chk = ax.compare_with_official(res.keys[res.keys['pn'] == 2], off[off['pn'] == 2])
        assert chk['totals']['차이'].abs().max() < 1e-6
        assert chk['project_match']['Output'] == 100
        assert ax.classify_differences(res, off, ir).empty

    def test_causes(self, scenario):
        _, ir, _, res = scenario
        off = self.as_official(res)
        a = off['_key'] == '82-00001.001100112'
        # 청구 월: 회계는 A를 3월에 청구 — 2월 Output 0·End 200만, 3월 Output 200만
        off.loc[a & (off['pn'] == 2), ['Value Output', 'Value End']] = [0.0, 2_000_000.0]
        extra = off[a & (off['pn'] == 2)].copy()
        extra[['pn', 'Period', 'Value Start', 'Value Output', 'Value End']] = [3, 'P03', 2_000_000.0, 2_000_000.0, 0.0]
        off = pd.concat([off, extra], ignore_index=True)
        # 회계 조정: B의 회계 2월 Start가 회계 1월 End와 다르다
        b = (off['_key'] == '82-00002.001500849') & (off['pn'] == 2)
        off.loc[b, 'Value Start'] = 450_000.0
        d = ax.classify_differences(res, off, ir)
        cause = d.set_index(['Period', 'Key'])['원인']
        assert cause[('P02', '82-00001.001100112')] == '청구 월'
        assert cause[('P03', '82-00001.001100112')] == '청구 월'
        assert cause[('P02', '82-00002.001500849')] == '회계 조정'


class TestDeptMapping:
    def test_learned_pairs_and_rule(self, scenario):
        official, ir, _, _ = scenario
        m = ax.learn_dept_map(official, ir)
        assert m['1012'] == '0112' and m['1849'] == '0849' and m['1034'] == '0134'
        assert '3060' not in m                                  # 82-00007.001은 자재 부서가 둘이라 짝을 못 배운다

    @pytest.mark.parametrize('material,expected', [('1007', '0107'), ('1228', '0328'), ('1600', '0700'),
                                                   ('1844', '0844'), ('3000', None), ('abc', None)])
    def test_rule(self, material, expected):
        assert ax.dept_by_rule(material) == expected


class TestFx:
    def _closing_file(self, folder):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '102026'
        ws.append(['Closing mid rates as at 22/10/2026'])
        ws.append([None])
        ws.append(['CURRENCY', None, 'Rate v GBP'])
        ws.append(['KOREAN WON', 'KRW', 1776.7799])
        ws.append(['US DOLLARS', 'USD', 1.3401])
        ws.append(['JAPANESE YEN', 'JPY', 210.083])
        path = folder / 'FX Rates - Closing rates - P10 2026.xlsx'
        wb.save(path)
        return path

    def test_closing_file_applies_to_next_month(self, tmp_path):
        applies, rates = ax.parse_closing_rates(self._closing_file(tmp_path))
        assert applies == '2026-11'
        assert rates['USD'] == pytest.approx(1776.7799 / 1.3401)
        assert rates['JPY'] == pytest.approx(1776.7799 / 210.083)
        assert rates['GBP'] == pytest.approx(1776.7799)

    def test_hub_rates_are_kept_in_history(self, tmp_path):
        hub = tmp_path / 'hub'
        hub.mkdir()
        self._closing_file(hub)
        history = tmp_path / 'fx_history.csv'
        book = ax.load_fx_book(None, history, hub)
        assert book.get('USD', '2026-11') == pytest.approx(1325.8562, abs=1e-3)
        assert history.exists()
        n = len(pd.read_csv(history))
        ax.load_fx_book(None, history, hub)              # 같은 달은 다시 쌓지 않는다
        assert len(pd.read_csv(history)) == n
        (hub / 'FX Rates - Closing rates - P10 2026.xlsx').unlink()
        assert ax.load_fx_book(None, history, hub).get('USD', '2026-11') == pytest.approx(1325.8562, abs=1e-3)

    def test_override_beats_everything(self):
        book = ax.load_fx_book(None, None, None, {('JPY', '2026-10'): 8.6157})
        assert book.get('JPY', '2026-10') == 8.6157
        assert book.get('JPY', '2026-09') == 8.6968           # 역산 시드

    def test_missing_month_falls_back_to_nearest_and_is_recorded(self):
        book = ax.FxBook()
        book.put('USD', '2026-08', 1_452.5, 'test')
        assert book.get('USD', '2026-10') == 1_452.5
        assert book.fallbacks[('USD', '2026-10')] == '2026-08'
        assert book.get('KRW', '2026-10') == 1.0
        assert list(book.rate_series(pd.Series(['KRW', 'USD']), pd.Series(['2026-01', '2026-08']))) == [1.0, 1_452.5]


class TestCurrency:
    def test_project_override_beats_customer(self):
        cust = pd.Series(['C01132', 'C01132', 'C00033'])
        proj = pd.Series(['82-17677.001', '82-15160.001', '82-1.001'])
        cur = ax.assign_currency(cust, proj, {'82-17677.001': 'EUR'})
        assert list(cur) == ['EUR', 'KRW', 'USD']

    def test_override_file(self, tmp_path):
        p = tmp_path / '통화지정.csv'
        p.write_text('코드,통화\n82-17677.001,eur\nC09999,USD\n', encoding='utf-8-sig')
        assert ax.read_currency_overrides(p) == {'82-17677.001': 'EUR', 'C09999': 'USD'}
        assert ax.read_currency_overrides(tmp_path / 'none.csv') == {}


class TestCli:
    def test_default_year(self):
        assert cli.default_year(12, date(2027, 1, 5)) == 2026
        assert cli.default_year(10, date(2026, 11, 2)) == 2026

    def test_fx_overrides(self):
        assert cli.parse_fx_overrides(['JPY=8.6157', 'eur=2026-09=1614.1'], '2026-10') == {
            ('JPY', '2026-10'): 8.6157, ('EUR', '2026-09'): 1614.1}
        with pytest.raises(ValueError):
            cli.parse_fx_overrides(['JPY'], '2026-10')

    def test_read_ir_and_it_require_columns(self, tmp_path):
        p = tmp_path / 'x.xlsx'
        pd.DataFrame([{'Project': '82-1.001'}]).to_excel(p, index=False)
        with pytest.raises(ValueError, match='없는 열'):
            ax.read_ir([p])
        with pytest.raises(ValueError, match='없는 열'):
            ax.read_it([p])

    def test_period(self):
        assert ax.parse_period('P10') == 10 and ax.parse_period('3') == 3
        with pytest.raises(ValueError):
            ax.parse_period('P13')

    def test_latest_extract_by_file_date(self, tmp_path):
        for name in ('Item transaction_20261001.xlsx', 'Item transaction_20261102.xlsx', 'Item requirement_20261102.xlsx'):
            (tmp_path / name).write_bytes(b'')
        assert [p.name for p in cli.find_it_files(tmp_path)] == ['Item transaction_20261102.xlsx']
        assert [p.name for p in cli.find_ir_files(tmp_path)] == ['Item requirement_20261102.xlsx']

    def test_official_needs_p01(self, tmp_path):
        p = tmp_path / 'Sales order book_02_2026.xlsx'
        official_df([official_row('82-1.001', '10', '0112', period='P02')])[list(ax.RAW_COLUMNS)].to_excel(
            p, sheet_name='Raw', index=False)
        with pytest.raises(FileNotFoundError, match='P01'):
            cli.load_official([p], 2026, 3)

    def test_main_writes_workbook(self, tmp_path, monkeypatch):
        ir_p, it_p = tmp_path / 'Item requirement_20260401.xlsx', tmp_path / 'Item transaction_20260401.xlsx'
        ob_p = tmp_path / 'Sales order book_01_2026.xlsx'
        pd.DataFrame(IR_LINES).to_excel(ir_p, index=False)
        pd.DataFrame(IT_ROWS).to_excel(it_p, index=False)
        official_df(OFFICIAL)[list(ax.RAW_COLUMNS)].to_excel(ob_p, sheet_name='Raw', index=False)
        monkeypatch.setattr(cli, 'load_fx_book', lambda *a, **kw: make_fx())
        monkeypatch.setattr(cli, 'CURRENCY_OVERRIDE_FILE', tmp_path / 'none.csv')
        rc = cli.main(['P03', '--year', '2026', '--ir', str(ir_p), '--it', str(it_p), '--orderbook', str(ob_p),
                       '--out', str(tmp_path / 'out')])
        assert rc == 0
        out = tmp_path / 'out' / 'AX_오더북_2026_P03.xlsx'
        with pd.ExcelFile(out) as x:
            assert {'요약', 'Raw', '차이_월별', '출고후미청구', '확인목록', '환율'} <= set(x.sheet_names)
            assert len(pd.read_excel(x, sheet_name='Raw')) > 0
