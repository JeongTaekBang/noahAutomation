"""
AX 오더북 — IR + Item transaction 엔진·저장소 (po_generator/ax_orderbook.py, ax_orderbook_store.py,
build_ax_orderbook.py) 테스트

한 프로젝트가 규칙 하나를 맡는 가짜 회계 오더북(P01 Start)·Item requirement·Item transaction으로 2026 P01~P03을
만들어 본다(기준선). 그 결과를 '저장된 달'로 두고, 나중 추출본으로 다음 달을 이어 계산한다(AX Sales Order Book처럼).
실측 근거는 tasks/todo.md 'AX 오더북' 절.
"""

from __future__ import annotations

import json
from datetime import date

import openpyxl
import pandas as pd
import pytest

import build_ax_orderbook as cli
from po_generator import ax_orderbook as ax
from po_generator import ax_orderbook_report as rep
from po_generator.ax_orderbook_store import BASELINE, CLOSED, MONTHLY, OPEN, OrderbookStore, PeriodError, Plan


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
            division: str = '10', vat: str = '01', item: str = 'ITEM') -> dict:
    net = 0.0 if status == 'Cancelled' else qty * unit_price
    return {
        'Project': project, 'Line number': str(line), 'Created date': pd.Timestamp(created_utc),
        'Customer': customer, 'Customer name': f'고객{customer}', 'Project group': 'ProdSale',
        'Item number': item, 'Item name': 'item', 'Quantity': qty, 'Unit price': unit_price,
        'Net amount': net, 'Line status': status, 'Actual ship date': pd.Timestamp(ship) if ship else pd.NaT,
        'Deliver remainder': remainder, 'Division': division, 'Prod/Dept (Material)': material,
        'Contract engineer': 'CE Kim', 'VAT group': vat,
    }


def it_row(project: str, line, when: str, qty: float, net: float, status: str = 'Invoiced',
           currency: str = 'KRW', prop: str = 'Invoiceable', item: str = 'ITEM') -> dict:
    return {'Project date': pd.Timestamp(when), 'Project': project, 'Line number': float(line),
            'Item number': item, 'Item name': 'item', 'Quantity': qty, 'Net amount': net,
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
    ir_line('82-00008.001', 1, '1012', 'Invoiced', 1, 100_000, '2026-01-05', ship='2026-03-03', item='GAUGE'),
    ir_line('82-00008.001', 1, '1849', 'Invoiced', 1, 250_000, '2026-02-01', ship='2026-02-20', item='POSITIONER'),
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
    it_row('82-00008.001', 1, '2026-02-20', 1, 250_000, item='POSITIONER'),
    it_row('82-00008.001', 1, '2026-03-03', 1, 100_000, item='GAUGE'),
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


def write_inputs(folder, monkeypatch, stamp: str = '20260401', ir_lines=None, it_rows=None) -> list[str]:
    """CLI 입력 파일(IR·IT·회계 오더북)을 쓰고 main()에 넘길 인자를 돌려준다 — 저장소·결과 폴더는 folder 안"""
    ir_p, it_p = folder / f'Item requirement_{stamp}.xlsx', folder / f'Item transaction_{stamp}.xlsx'
    ob_p = folder / 'Sales order book_01_2026.xlsx'
    pd.DataFrame(ir_lines or IR_LINES).to_excel(ir_p, index=False)
    pd.DataFrame(it_rows or IT_ROWS).to_excel(it_p, index=False)
    if not ob_p.exists():
        official_df(OFFICIAL)[list(ax.RAW_COLUMNS)].to_excel(ob_p, sheet_name='Raw', index=False)
    monkeypatch.setattr(cli, 'load_fx_book', lambda *a, **kw: make_fx())
    monkeypatch.setattr(cli, 'CURRENCY_OVERRIDE_FILE', folder / 'none.csv')
    return ['--year', '2026', '--ir', str(ir_p), '--it', str(it_p), '--orderbook', str(ob_p),
            '--store', str(folder / 'store.db'), '--out', str(folder / 'out')]


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

    def test_manual_rate_is_remembered(self, tmp_path):
        """--fx 환율도 이력에 남는다 — 저장한 달의 End를 매긴 환율이 다음 달 계산에서 빠지면 가짜 Variance가 생긴다"""
        history = tmp_path / 'fx_history.csv'
        ax.load_fx_book(None, history, None, {('USD', '2026-11'): 1_400.0})
        book = ax.load_fx_book(None, history, None)
        assert book.get('USD', '2026-11') == 1_400.0
        hub = tmp_path / 'hub'
        hub.mkdir()
        self._closing_file(hub)                          # P10 파일 = 2026-11 — 직접 입력이 있어도 쌓인다
        book = ax.load_fx_book(None, history, hub)
        assert book.get('USD', '2026-11') == 1_400.0     # 직접 입력이 이긴다
        assert book.get('JPY', '2026-11') == pytest.approx(1776.7799 / 210.083)
        ax.load_fx_book(None, history, None, {('USD', '2026-11'): 1_410.0})   # 다시 넣으면 바꿔 넣는다
        assert ax.load_fx_book(None, history, None).get('USD', '2026-11') == 1_410.0
        assert (pd.read_csv(history)['source'] == ax.MANUAL_FX_SOURCE).sum() == 1

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
        args = write_inputs(tmp_path, monkeypatch)
        rc = cli.main(['P03', *args])
        assert rc == 0
        out = tmp_path / 'out' / 'AX_오더북_2026_P03.xlsx'
        with pd.ExcelFile(out) as x:
            assert {'요약', '기간', 'Raw', '차이_월별', '출고후미청구', '확인목록', '환율'} <= set(x.sheet_names)
            assert len(pd.read_excel(x, sheet_name='Raw')) > 0
        assert out.with_suffix('.html').stat().st_size > 1_000_000         # 엑셀 옆에 HTML (Plotly 엔진 포함)

    def test_reference_orderbook_falls_back_to_last_year(self, tmp_path):
        """1월 계산엔 그 해 회계 파일이 아직 없다 — 지난 해 파일로 매핑·속성만, 대조는 하지 않는다"""
        rows = official_df([official_row('82-1.001', '10', '0112', period='P10')])
        rows['Report Start'] = pd.Timestamp('2026-10-01')
        p = tmp_path / 'Sales order book_10_2026.xlsx'
        rows[list(ax.RAW_COLUMNS)].to_excel(p, sheet_name='Raw', index=False)
        ref, path = cli.load_reference([p], 2027, 1)
        assert path == p and list(ref['year']) == [2026]
        assert cli.official_of_year(ref, 2027).empty
        assert cli.load_reference([], 2027, 1)[1] is None


VALUES = ['Sq', 'Iq', 'Oq', 'Eq', 'Vq', 'Sv', 'Iv', 'Ov', 'Ev', 'Vv']


def as_stored(res: ax.OrderbookResult, month: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """기준선 결과의 한 달을 '저장된 달'처럼 — (Raw 행, 청구 누계)"""
    rows = res.raw_rows[res.raw_rows['pn'] == month].assign(year=2026)
    billed = res.billed[res.billed['pn'] == month].set_index('bgroup')[list(ax.BILLED_COLUMNS)]
    return rows, billed


def month_values(res: ax.OrderbookResult, month: int | None = None) -> pd.DataFrame:
    t = res.keys if month is None else res.keys[res.keys['pn'] == month]
    return t.set_index('Key')[VALUES].sort_index()


class TestMonthly:
    """다음 달 계산 — 저장된 직전 달(End·청구 누계)에서 시작한다. 지난 달은 그대로, 그 뒤의 일은 이번 달로"""

    def test_same_extract_continues_the_baseline(self, scenario):
        """기준선 P02를 저장해 두고 같은 추출본으로 P03을 계산하면 기준선 P03과 같다 — 계산 경로가 하나"""
        official, ir, it, res = scenario
        rows, billed = as_stored(res, 2)
        history = res.raw_rows[res.raw_rows['pn'] <= 2].assign(year=2026)
        nxt = ax.build_month(official, ir, it, 2026, 3, make_fx(), rows, billed, history)
        pd.testing.assert_frame_equal(month_values(nxt), month_values(res, 3), atol=1e-6)
        order = ['Project', 'Cost center', 'Prod/Dept (Revenue)']
        a = nxt.raw_rows[list(ax.RAW_COLUMNS)].sort_values(order).reset_index(drop=True)
        b = res.raw_rows[res.raw_rows['pn'] == 3][list(ax.RAW_COLUMNS)].sort_values(order).reset_index(drop=True)
        pd.testing.assert_frame_equal(a, b, check_dtype=False, atol=1e-6)

    def test_billing_after_close_is_output_of_the_billing_month(self, scenario):
        """지난 마감 때 청구 전(Posted)이던 1월 출고가 나중 추출본에선 청구됨 → 이번 달 Output (AX처럼 청구한 달)"""
        official, ir, _, res = scenario
        rows, billed = as_stored(res, 2)
        later = [dict(r, **{'Transaction status': 'Invoiced'}) if r['Project'] == '82-00005.001' else r
                 for r in IT_ROWS]
        nxt = ax.build_month(official, ir, ax.prepare_it(pd.DataFrame(later)), 2026, 3, make_fx(), rows, billed)
        row = k(nxt, '82-00005.001100107', 3)
        assert (row['Sv'], row['Oq'], row['Ov'], row['Ev'], row['Vv']) == (300_000, 1, 300_000, 0, 0)
        assert k(res, '82-00005.001100107', 2)['Ev'] == 300_000            # 저장된 달은 그대로

    def test_cancellation_after_close_is_variance_of_this_month(self, scenario):
        """저장한 뒤 취소된 라인 — Start는 저장된 End 그대로, 이번 달 Variance로 빠진다"""
        official, _, it, res = scenario
        rows, billed = as_stored(res, 2)
        history = res.raw_rows[res.raw_rows['pn'] <= 2].assign(year=2026)
        later = [dict(r, **{'Line status': 'Cancelled', 'Net amount': 0.0}) if r['Project'] == '82-00013.001'
                 else r for r in IR_LINES]
        nxt = ax.build_month(official, ax.prepare_ir(pd.DataFrame(later)), it, 2026, 3, make_fx(), rows, billed,
                             history)
        row = k(nxt, '82-00013.001100112', 3)
        assert (row['Project'], row['Sv'], row['Iv'], row['Ov'], row['Ev'], row['Vv']) == (
            '82-00013.001', 50_000, 0, 0, 0, -50_000)

    def test_credit_after_billing_is_negative_output(self, scenario):
        official, ir, _, res = scenario
        rows, billed = as_stored(res, 2)
        later = IT_ROWS + [it_row('82-00001.001', 1, '2026-03-05', -2, -2_000_000)]
        nxt = ax.build_month(official, ir, ax.prepare_it(pd.DataFrame(later)), 2026, 3, make_fx(), rows, billed)
        row = k(nxt, '82-00001.001100112', 3)
        assert (row['Oq'], row['Ov'], row['Ev'], row['Vv']) == (-2, -2_000_000, 0, -2_000_000)

    def test_attribute_history_orders_by_year_then_month(self):
        """작년 회계 파일(P10)과 저장된 작년 12월 행 — 연도까지 보고 12월 행이 나중이다"""
        off = official_df([official_row('82-1.001', '10', '0112', period='P10')]).assign(year=2026)
        stored = off.assign(_key=off['Key'], pn=12, Market='STORED')
        h = ax.attribute_history(off, stored, 2027)
        assert list(h['pn']) == [202610, 202612]
        assert h.iloc[-1]['Market'] == 'STORED'


def norm(df: pd.DataFrame) -> pd.DataFrame:
    """저장소 왕복 비교용 — 빈 값은 None, 행 순서는 달·키 순"""
    out = df.astype(object).where(df.notna(), None)
    return out.sort_values(['pn', '_key']).reset_index(drop=True)


class TestStore:
    def test_empty_store_is_baseline_and_reading_creates_nothing(self, tmp_path):
        store = OrderbookStore(tmp_path / 'ob.db')
        assert store.plan(2026, 10) == Plan('baseline', None)
        assert store.rows().empty and store.periods().empty
        assert not store.path.exists()

    def test_round_trip(self, tmp_path, scenario):
        res = scenario[3]
        store = OrderbookStore(tmp_path / 'ob.db')
        store.save(2026, [1, 2, 3], res.raw_rows, res.billed, BASELINE, 'ir.xlsx', 'it.xlsx', '2026-04-01')
        cols = ['pn', '_key', *ax.RAW_COLUMNS]
        pd.testing.assert_frame_equal(norm(store.rows(2026)[cols]), norm(res.raw_rows[cols]), check_dtype=False)
        b = res.billed[res.billed['pn'] == 3].set_index('bgroup')[list(ax.BILLED_COLUMNS)]
        pd.testing.assert_frame_equal(store.billed(2026, 3).sort_index(), b.sort_index(), check_dtype=False)
        p = store.periods(2026)
        assert list(p['status']) == [CLOSED, CLOSED, OPEN]
        assert set(p['basis']) == {BASELINE} and set(p['snapshot']) == {'2026-04-01'}
        assert list(p['row_count']) == list(res.raw_rows['pn'].value_counts().sort_index())

    def test_period_rules(self, tmp_path, scenario):
        res = scenario[3]
        store = OrderbookStore(tmp_path / 'ob.db')
        store.save(2026, [1, 2, 3], res.raw_rows, res.billed, BASELINE)
        assert store.plan(2026, 3) == Plan('recompute', (2026, 2))         # 마지막 달(Open)은 다시 계산
        assert store.plan(2026, 4) == Plan('next', (2026, 3))
        with pytest.raises(PeriodError, match='닫혔'):
            store.plan(2026, 2)
        with pytest.raises(PeriodError, match='건너뛸'):
            store.plan(2026, 5)
        with pytest.raises(PeriodError, match='남아 있지 않습니다'):
            store.billed(2026, 1)                    # 청구 누계는 마지막 달과 그 앞 달 것만 남는다
        with pytest.raises(PeriodError, match='닫힌 달'):
            store.save(2026, [2, 3], res.raw_rows[res.raw_rows['pn'] >= 2], res.billed, MONTHLY)   # 저장소도 지킨다

        p4 = res.raw_rows[res.raw_rows['pn'] == 3].assign(pn=4)
        b4 = res.billed[res.billed['pn'] == 3].assign(pn=4)
        store.save(2026, [4], p4, b4, MONTHLY)
        assert list(store.periods(2026)['status']) == [CLOSED, CLOSED, CLOSED, OPEN]
        with pytest.raises(PeriodError, match='닫혔'):
            store.plan(2026, 3)
        with pytest.raises(PeriodError):
            store.billed(2026, 2)
        assert len(store.billed(2026, 3)) == len(b4)
        store.save(2026, [4], p4, b4, MONTHLY)        # 마지막 달 다시 저장 = 바꿔 넣기
        assert len(store.rows(2026, 4)) == len(p4) and len(store.periods()) == 4

    def test_year_boundary(self, tmp_path, scenario):
        res = scenario[3]
        store = OrderbookStore(tmp_path / 'ob.db')
        rows, billed = res.raw_rows[res.raw_rows['pn'] == 3], res.billed[res.billed['pn'] == 3]
        store.save(2026, [12], rows.assign(pn=12), billed.assign(pn=12), BASELINE)
        assert store.plan(2027, 1) == Plan('next', (2026, 12))
        store.save(2027, [1], rows.assign(pn=1), billed.assign(pn=1), MONTHLY)
        p = store.periods()
        assert list(zip(p['year'], p['pn'], p['status'])) == [(2026, 12, CLOSED), (2027, 1, OPEN)]
        assert len(store.billed(2026, 12)) == len(billed)         # 앞 달(작년 12월) 누계는 남는다


class TestLifecycle:
    """CLI — 기준선 → 다음 달 → 다시 계산, 추출 시점 점검"""

    def test_baseline_then_next_month_equals_one_baseline(self, tmp_path, monkeypatch, scenario):
        args = write_inputs(tmp_path, monkeypatch)
        assert cli.main(['P02', *args]) == 0                       # 기준선 P01~P02
        assert cli.main(['P03', *args]) == 0                       # 저장된 P02에서 P03
        store = OrderbookStore(tmp_path / 'store.db')
        p = store.periods(2026)
        assert list(zip(p['pn'], p['status'], p['basis'])) == [(1, CLOSED, BASELINE), (2, CLOSED, BASELINE),
                                                               (3, OPEN, MONTHLY)]
        got = ax.keys_from_raw(store.rows(2026, 3)).set_index('Key')[VALUES].sort_index()
        pd.testing.assert_frame_equal(got, month_values(scenario[3], 3), check_dtype=False, atol=1e-6)
        with pd.ExcelFile(tmp_path / 'out' / 'AX_오더북_2026_P03.xlsx') as x:
            assert set(pd.read_excel(x, sheet_name='Raw')['Period']) == {'P01', 'P02', 'P03'}
            assert list(pd.read_excel(x, sheet_name='기간')['상태']) == [CLOSED, CLOSED, OPEN]

    def test_closed_month_is_refused_and_open_month_is_recomputed(self, tmp_path, monkeypatch, capsys):
        args = write_inputs(tmp_path, monkeypatch)
        assert cli.main(['P02', *args]) == 0
        store = OrderbookStore(tmp_path / 'store.db')
        before = store.rows(2026, 2)
        assert cli.main(['P01', *args]) == 1
        assert '닫혔' in capsys.readouterr().out
        assert cli.main(['P02', *args]) == 0                       # 같은 추출본으로 다시 계산 → 같은 숫자
        p = store.periods(2026)
        assert list(zip(p['status'], p['basis'])) == [(CLOSED, BASELINE), (OPEN, MONTHLY)]
        cols = ['pn', '_key', *ax.RAW_COLUMNS]
        pd.testing.assert_frame_equal(norm(store.rows(2026, 2)[cols]), norm(before[cols]), check_dtype=False)

    def test_pre_close_extract_is_not_saved_but_can_be_previewed(self, tmp_path, monkeypatch):
        args = write_inputs(tmp_path, monkeypatch, stamp='20260320')
        assert cli.main(['P03', *args]) == 1                       # 3월이 끝나기 전 추출본 — 저장 거부
        assert cli.main(['P03', *args, '--dry-run']) == 0           # 미리보기는 된다
        assert (tmp_path / 'out' / 'AX_오더북_2026_P03_미리보기.xlsx').exists()
        assert (tmp_path / 'out' / 'AX_오더북_2026_P03_미리보기.html').exists()
        assert not (tmp_path / 'store.db').exists()

    def test_snapshot_checks(self, tmp_path):
        def pair(ir_day: str, it_day: str):
            return [tmp_path / f'Item requirement_{ir_day}.xlsx'], [tmp_path / f'Item transaction_{it_day}.xlsx']
        assert cli.check_snapshot(*pair('20261102', '20261102'), 2026, 10) == (date(2026, 11, 2), [], [])
        assert cli.check_snapshot(*pair('20261107', '20261107'), 2026, 10)[2] == []
        _, errors, _ = cli.check_snapshot(*pair('20261102', '20261103'), 2026, 10)
        assert any('추출일이 다릅니다' in e for e in errors)
        _, errors, _ = cli.check_snapshot(*pair('20261030', '20261030'), 2026, 10)
        assert any('끝나기 전' in e for e in errors)
        _, errors, warns = cli.check_snapshot(*pair('20261108', '20261108'), 2026, 10)
        assert not errors and any('8일째' in w for w in warns)
        _, errors, warns = cli.check_snapshot([tmp_path / 'ir.xlsx'], [tmp_path / 'it.xlsx'], 2026, 10)
        assert not errors and warns


def unpack(table: dict, strings: list[str]) -> pd.DataFrame:
    """report_data의 열 단위 표 → DataFrame (사전 열은 문자열로 되돌린다 — 템플릿 JS의 rows()와 같은 일)"""
    df = pd.DataFrame(table['cols'])
    for c in table['str']:                      # 빈 값이 섞인 열은 float가 되므로 int로
        df[c] = df[c].map(lambda i: None if pd.isna(i) else strings[int(i)])
    return df


class TestReport:
    """결과 HTML — 데이터(report_data)와 파일 쓰기(write_report). 화면 로직(JS)은 실제 브라우저로 확인한다"""

    @pytest.fixture
    def report(self, scenario):
        _, ir, _, res = scenario
        off = res.raw_rows.copy()
        off['Key'] = off['_key']
        a = (off['_key'] == '82-00001.001100112') & (off['pn'] == 2)
        off.loc[a, ['Value Output', 'Value End']] = [0.0, 2_000_000.0]          # 회계는 A를 3월에 청구했다
        diff = ax.classify_differences(res, off, ir)
        periods = pd.DataFrame([{'Period': 'P01', '상태': 'Closed', 'Raw 행': 10}])
        data = rep.report_data(res, diff, off, periods, {'Item requirement': 'ir.xlsx'}, '방식', preview=True)
        return res, off, diff, data

    def test_totals_match_the_engine(self, report):
        res, off, _, data = report
        ours = unpack(data['ours'], data['strings']).groupby('pn')[['S', 'I', 'O', 'V', 'E']].sum()
        want = res.keys.groupby('pn')[['Sv', 'Iv', 'Ov', 'Vv', 'Ev']].sum()
        assert abs(ours.to_numpy() - want.to_numpy()).max() < 0.01
        acct = unpack(data['acct'], data['strings']).groupby('pn')['E'].sum()
        assert acct.to_dict() == pytest.approx(off.groupby('pn')['Value End'].sum().to_dict())

    def test_keys_carry_filter_attributes(self, report):
        res, _, _, data = report
        keys = unpack(data['keys'], data['strings'])
        ours = unpack(data['ours'], data['strings'])
        assert set(keys['key']) == set(res.keys['Key'])
        a = keys.set_index('key').loc['82-00001.001100112']
        assert (a['project'], a['market'], a['obc'], a['cc']) == ('82-00001.001', 'CP&I', 'IQ3', '10')
        assert ours['k'].map(keys['key']).tolist() == res.keys['Key'].tolist()      # 번호가 같은 키를 가리킨다

    def test_diff_rows_point_to_keys_and_causes(self, report):
        _, _, diff, data = report
        d = unpack(data['diff'], data['strings'])
        keys = unpack(data['keys'], data['strings'])
        assert len(d) == len(diff) > 0
        for row, (_, want) in zip(d.itertuples(), diff.iterrows()):
            assert keys['key'][int(row.k)] == want['Key']
            assert data['causes'][int(row.c)]['name'] == want['원인']
        assert set(ax.DIFF_CAUSE_NOTES) == set(ax.DIFF_CAUSES)

    def test_periods_and_meta(self, report):
        data = report[3]
        assert [p['label'] for p in data['periods']] == ['P01', 'P02', 'P03']
        assert all(p['official'] for p in data['periods'])
        assert (data['meta']['title'], data['meta']['preview']) == ('AX 오더북 2026 P03', True)
        json.dumps(data, allow_nan=False)                                     # NaN이 없어야 JS가 읽는다

    def test_write_report_embeds_data_and_engine(self, report, tmp_path):
        from plotly.offline import get_plotlyjs
        data = report[3]
        data['strings'].append('</script><b>x<!--')        # 태그를 닫으려는 문자열도 데이터 안에 갇혀야 한다
        html = rep.write_report(data, tmp_path / 'r.html').read_text(encoding='utf-8')
        assert rep.PLACEHOLDER_DATA not in html and rep.PLACEHOLDER_PLOTLY not in html
        assert '</script><b>x' not in html
        payload = html.split('<script type="application/json" id="ax-data">')[1].split('</script>')[0]
        assert json.loads(payload) == data
        assert get_plotlyjs()[:200] in html

    def test_template_needs_each_placeholder_once(self, tmp_path):
        bad = tmp_path / 't.html'
        bad.write_text('<html>__AX_DATA__</html>', encoding='utf-8')
        with pytest.raises(ValueError, match='자리표시자'):
            rep.write_report({'a': 1}, tmp_path / 'x.html', template=bad)
