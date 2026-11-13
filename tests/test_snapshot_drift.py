"""
마감 후 변동 — 스냅샷이 굳은 뒤 원 데이터가 바뀌어 다음 마감에 Variance로 잡힐 것
================================================================================

`snapshot.post_close_drift()`는 대시보드 '정합성 → 마감 후 변동' 탭의 원본이다.
여기서 지키고 싶은 것은 둘이다:

1. **다음 마감이 적을 값과 같다** — 같은 `_ORDER_BOOK_BASE_SQL`·같은 키·같은 문턱으로 구하지만
   `take_snapshot`과 코드를 나눠 쓰지는 않는다(마감 경로는 건드리지 않았다). 그래서 실제로
   다음 달을 마감해 `ob_snapshot.variance_amount`와 원 단위로 맞춰 본다.
2. **정상 업무가 '변동'으로 보이지 않는다** — 납기만 바뀐 주문은 키가 옮겨갈 뿐 금액은 그대로다.

저장소의 첫 스냅샷 테스트이기도 하다 — DB는 `test_order_book_fx._mk_db`(뷰 포함)로 만든다.
"""

from __future__ import annotations

import sqlite3
from contextlib import closing

import pandas as pd
import pytest

import dashboard
from po_generator.snapshot import SnapshotEngine, post_close_drift
from tests.test_order_book_fx import _mk_db

JAN, FEB, MAR = '2026-01', '2026-02', '2026-03'


def so(so_id, line, amount, period=JAN, **kw):
    """국내 SO 라인 — 기본은 1월 수주, 납기 3/1"""
    row = {'SO_ID': so_id, 'Line item': str(line), 'Customer name': '고객A', 'Customer PO': 'CPO-1',
           'Item name': 'IQ3', 'OS name': 'IQ3', 'Item qty': '10', 'Sales amount': str(amount),
           'Period': period, 'Expected delivery date': '2026-03-01', 'Status': '미출고'}
    row.update(kw)
    return row


def dn(dn_id, so_id, line, qty, amount, invoice=None, **kw):
    """국내 DN 라인 — `invoice`(세금계산서 발행일)가 없으면 미인식 = Backlog 잔류"""
    row = {'DN_ID': dn_id, 'SO_ID': so_id, 'Line item': str(line), 'Qty': str(qty),
           'Total Sales': str(amount), '출고일': '2026-01-20', '세금계산서 발행일': invoice,
           'Customer name': '고객A', 'Item': 'IQ3'}
    row.update(kw)
    return row


class Ledger:
    """임시 DB 한 벌 — 마감(엔진), 수정(SQL), 조회(post_close_drift)를 각자 새 연결로"""

    def __init__(self, tmp_path, **tables):
        _mk_db(tmp_path, **tables).close()
        self.path = tmp_path / 'ob.db'
        self.engine = SnapshotEngine(self.path)

    def close_period(self, period):
        result = self.engine.take_snapshot(period)
        assert result.success, result.message

    def run(self, sql, *params):
        with closing(sqlite3.connect(str(self.path))) as conn:
            conn.execute(sql, params)
            conn.commit()

    def query(self, sql, *params):
        with closing(sqlite3.connect(str(self.path))) as conn:
            return conn.execute(sql, params).fetchall()

    def drift(self):
        with closing(sqlite3.connect(str(self.path))) as conn:
            return post_close_drift(conn)

    def rows(self):
        return {r.so_id: r for r in self.drift().rows}


# ═══════════════════════════════════════════════════════════════
# 마감이 없거나, 마감 뒤 아무것도 안 바뀐 경우
# ═══════════════════════════════════════════════════════════════
class TestNothingToShow:

    def test_no_close_yet(self, tmp_path):
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        assert ledger.drift() is None, '스냅샷 테이블이 아예 없는 DB — 예외가 아니라 None'

    def test_undone_close_does_not_count(self, tmp_path):
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        assert ledger.engine.undo_snapshot(JAN).success
        assert ledger.drift() is None

    def test_right_after_close(self, tmp_path):
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)],
                        dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400, invoice='2026-01-25')])
        ledger.close_period(JAN)
        drift = ledger.drift()
        assert (drift.period, drift.rows, drift.moved_only) == (JAN, (), 0)
        assert drift.closed_at == ledger.query('SELECT closed_at FROM ob_snapshot_meta')[0][0]


# ═══════════════════════════════════════════════════════════════
# 무엇이 바뀌었나 — 마감월 수주 / 마감월 매출 / 그 밖
# ═══════════════════════════════════════════════════════════════
class TestWhatChanged:

    def test_tax_invoice_entered_after_close(self, tmp_path):
        """실측(2026-10): 10/30에 마감하고 10/31자 세금계산서를 11/5에 입력 — 53행 34.9M"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)],
                        dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400)])  # 출고는 했고 계산서는 아직
        ledger.close_period(JAN)
        ledger.run("UPDATE dn_domestic SET [세금계산서 발행일] = '2026-01-31' WHERE DN_ID = 'DND-1'")

        row = ledger.rows()['SOD-1']
        assert (row.d_input, row.d_output, row.rest) == (0, 400, 0)
        assert (row.snap_ending, row.live_ending, row.variance) == (1000, 600, -400)
        assert (row.customer, row.market) == ('고객A', '국내')

    def test_order_registered_after_close(self, tmp_path):
        """마감월 Period로 뒤늦게 등록한 주문 — 스냅샷엔 없던 수주"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.run("INSERT INTO so_domestic (SO_ID, [Line item], [Customer name], [OS name], "
                   "[Item qty], [Sales amount], Period, [Expected delivery date], Status) "
                   "VALUES ('SOD-2', '1', '고객B', 'IQ3', '2', '500', ?, '', '미출고')", JAN)

        rows = ledger.rows()
        assert set(rows) == {'SOD-2'}
        row = rows['SOD-2']
        assert (row.d_input, row.d_output, row.rest, row.variance) == (500, 0, 0, 500)
        assert (row.snap_ending, row.live_ending, row.customer) == (0, 500, '고객B')

    def test_price_cut_on_an_earlier_month(self, tmp_path):
        """실측: 3월 수주 `SOD-2026-0333`의 단가를 10월 마감 뒤에 내렸다 — 마감월 숫자는 그대로, 잔고만 준다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.close_period(FEB)
        ledger.run("UPDATE so_domestic SET [Sales amount] = '900' WHERE SO_ID = 'SOD-1'")

        drift = ledger.drift()
        assert drift.period == FEB
        row = drift.rows[0]
        assert (row.d_input, row.d_output, row.rest, row.variance) == (0, 0, -100, -100)

    @pytest.mark.parametrize('sql', [
        "DELETE FROM so_domestic WHERE SO_ID = 'SOD-1'",
        "UPDATE so_domestic SET Status = 'Cancelled' WHERE SO_ID = 'SOD-1'",
    ])
    def test_order_removed_after_close(self, tmp_path, sql):
        """지우든 취소하든 Order Book에서는 빠진다 — 고객명은 스냅샷에 남은 것으로 보여 준다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000), so('SOD-2', 1, 700)])
        ledger.close_period(JAN)
        ledger.run(sql)

        rows = ledger.rows()
        assert set(rows) == {'SOD-1'}
        row = rows['SOD-1']
        assert (row.d_input, row.variance, row.live_ending) == (-1000, -1000, 0)
        assert (row.customer, row.market) == ('고객A', '국내')

    def test_booked_and_billed_after_close_nets_to_zero_but_is_listed(self, tmp_path):
        """수주와 매출이 같이 들어오면 잔고 차이는 0이지만 마감월 숫자는 둘 다 스냅샷과 다르다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.run("INSERT INTO so_domestic (SO_ID, [Line item], [Customer name], [OS name], "
                   "[Item qty], [Sales amount], Period, [Expected delivery date], Status) "
                   "VALUES ('SOD-2', '1', '고객B', 'IQ3', '3', '300', ?, '', '출고 완료')", JAN)
        ledger.run("INSERT INTO dn_domestic (DN_ID, SO_ID, [Line item], Qty, [Total Sales], "
                   "[출고일], [세금계산서 발행일]) "
                   "VALUES ('DND-9', 'SOD-2', '1', '3', '300', '2026-01-30', '2026-01-30')")

        row = ledger.rows()['SOD-2']
        assert (row.d_input, row.d_output, row.rest, row.variance) == (300, 300, 0, 0)

    def test_export_shipped_after_close_puts_revaluation_in_rest(self, tmp_path):
        """해외는 선적월 환율로 다시 매긴다 — 매출은 재환산액, 잔고는 수주시점 금액만큼 줄고, 차액이 '그 밖'"""
        ledger = Ledger(
            tmp_path,
            so_exp=[{'SO_ID': 'SOO-1', 'Line item': '1', 'Customer name': 'ACME', 'OS name': 'IQ3',
                     'Item qty': '1', 'Sales amount KRW': '1500', 'Period': JAN,
                     'Expected delivery date': '', 'Status': '미출고'}],
            dn_exp=[{'DN_ID': 'DNO-1', 'SO_ID': 'SOO-1', 'Line item': '1', 'Qty': '1',
                     'Currency': 'EUR', 'Total Sales': '1', 'Total Sales KRW': '1500',
                     '출고일': '2026-01-15'}],
            fx=[('EUR', JAN, 1600)])
        ledger.close_period(JAN)
        ledger.run("UPDATE dn_export SET [선적일] = '2026-01-20' WHERE DN_ID = 'DNO-1'")

        row = ledger.rows()['SOO-1']
        assert (row.d_input, row.d_output, row.rest, row.variance) == (0, 1600, 100, -1500)
        assert row.market == '해외'

    def test_fx_rate_corrected_after_close_moves_revenue_but_not_backlog(self, tmp_path):
        """마감월 환율을 FX 시트에서 고치면 매출 재환산액만 바뀐다 — 잔고는 환율과 무관해 Variance 0

        다음 마감에 아무것도 적히지 않지만 스냅샷의 그 달 Output은 낡았으므로 나열한다.
        """
        ledger = Ledger(
            tmp_path,
            so_exp=[{'SO_ID': 'SOO-1', 'Line item': '1', 'Customer name': 'ACME', 'OS name': 'IQ3',
                     'Item qty': '1', 'Sales amount KRW': '1500', 'Period': JAN,
                     'Expected delivery date': '', 'Status': '출고 완료'}],
            dn_exp=[{'DN_ID': 'DNO-1', 'SO_ID': 'SOO-1', 'Line item': '1', 'Qty': '1',
                     'Currency': 'EUR', 'Total Sales': '1', 'Total Sales KRW': '1500',
                     '출고일': '2026-01-15', '선적일': '2026-01-20'}],
            fx=[('EUR', JAN, 1600)])
        ledger.close_period(JAN)
        ledger.run("UPDATE fx SET rate = 1700 WHERE currency = 'EUR' AND ym = ?", JAN)

        row = ledger.rows()['SOO-1']
        assert (row.d_input, row.d_output, row.rest, row.variance) == (0, 100, 100, 0)

    def test_quantity_only_change_is_listed_not_counted_as_a_date_move(self, tmp_path):
        """0원 라인(FOC) 추가·수량 정정은 금액이 그대로여도 엔진이 수량 Variance로 적는다

        '납기만 바뀜'으로 세면 안 된다 — 금액 표에서 사라진 채 다음 마감에 variance_qty가 생긴다.
        """
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.run("INSERT INTO so_domestic (SO_ID, [Line item], [Customer name], [OS name], "
                   "[Item qty], [Sales amount], Period, [Expected delivery date], Status) "
                   "VALUES ('SOD-1', '2', '고객A', 'IQ3', '2', '0', ?, '2026-03-01', '미출고')", JAN)

        drift = ledger.drift()
        assert drift.moved_only == 0
        row = drift.rows[0]
        assert (row.so_id, row.variance, row.variance_qty) == ('SOD-1', 0, 2)
        assert (row.d_input, row.d_output, row.rest) == (0, 0, 0)

    def test_broken_snapshot_tables_raise_instead_of_reading_as_no_close(self, tmp_path):
        """'마감 없음'(None)은 테이블이 아예 없을 때뿐 — 조회 실패를 None으로 가리면 '달라진 것 없음'이 된다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.run('DROP TABLE ob_snapshot')
        with pytest.raises(sqlite3.OperationalError):
            ledger.drift()


# ═══════════════════════════════════════════════════════════════
# 변동이 아닌 것
# ═══════════════════════════════════════════════════════════════
class TestNotDrift:

    def test_delivery_date_change_only_moves_the_key(self, tmp_path):
        """납기는 수식 열이라 수시로 바뀐다 — 키(SO, OS name, 납기)만 옮겨갈 뿐 금액은 그대로

        실측(2026-10 마감 이후): 키 단위로는 112행·절대값 7.6억이 움직였는데 그중 29주문은 순액 0이었다.
        """
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000), so('SOD-2', 1, 700)])
        ledger.close_period(JAN)
        ledger.run("UPDATE so_domestic SET [Expected delivery date] = '2026-04-15' WHERE SO_ID = 'SOD-1'")

        drift = ledger.drift()
        assert drift.rows == ()
        assert drift.moved_only == 1

    def test_changes_in_the_open_month_are_not_drift(self, tmp_path):
        """마감 다음 달의 수주·매출은 그 달의 Input/Output이다 — 지난 스냅샷과 비교할 것이 아니다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)],
                        dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400)])
        ledger.close_period(JAN)
        ledger.run("UPDATE dn_domestic SET [세금계산서 발행일] = '2026-02-10' WHERE DN_ID = 'DND-1'")
        ledger.run("INSERT INTO so_domestic (SO_ID, [Line item], [Customer name], [OS name], "
                   "[Item qty], [Sales amount], Period, [Expected delivery date], Status) "
                   "VALUES ('SOD-2', '1', '고객B', 'IQ3', '2', '500', ?, '', '미출고')", FEB)

        drift = ledger.drift()
        assert (drift.rows, drift.moved_only) == ((), 0)


# ═══════════════════════════════════════════════════════════════
# 불변식 — 다음 마감이 적는 Variance와 같다
# ═══════════════════════════════════════════════════════════════
class TestMatchesNextClose:

    def test_same_as_what_take_snapshot_books(self, tmp_path):
        """마감 → 이것저것 수정 → (우리 계산) == (실제로 다음 달을 마감했을 때 엔진이 적은 값)

        국내(KRW)만 쓴다 — 엔진은 `variance_amount`에 그 달 환율 재평가분을 합쳐 적으므로
        해외 건이 섞이면 소급 변경분만 떼어 비교할 수 없다.
        """
        ledger = Ledger(
            tmp_path,
            so_dom=[so('SOD-1', 1, 1000), so('SOD-2', 1, 700), so('SOD-3', 1, 500),
                    so('SOD-4', 1, 300), so('SOD-5', 1, 900, period=FEB)],
            dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400),                        # 계산서 아직
                    dn('DND-5', 'SOD-5', 1, 5, 450, invoice='2026-02-10')])
        ledger.close_period(JAN)
        ledger.close_period(FEB)

        ledger.run("UPDATE dn_domestic SET [세금계산서 발행일] = '2026-02-28' WHERE DN_ID = 'DND-1'")  # 마감월 매출
        ledger.run("UPDATE so_domestic SET [Sales amount] = '650' WHERE SO_ID = 'SOD-2'")            # 이전 달 단가
        ledger.run("UPDATE so_domestic SET [Expected delivery date] = '2026-05-01' "
                   "WHERE SO_ID = 'SOD-3'")                                                          # 납기만
        ledger.run("DELETE FROM so_domestic WHERE SO_ID = 'SOD-4'")                                  # 삭제
        ledger.run("UPDATE so_domestic SET [Sales amount] = '950' WHERE SO_ID = 'SOD-5'")            # 마감월 수주
        ledger.run("INSERT INTO so_domestic (SO_ID, [Line item], [Customer name], [OS name], "
                   "[Item qty], [Sales amount], Period, [Expected delivery date], Status) "
                   "VALUES ('SOD-6', '1', '고객B', 'IQ3', '2', '200', ?, '', '미출고')", FEB)        # 뒤늦은 등록

        drift = ledger.drift()
        ours = {r.so_id: r.variance for r in drift.rows}
        assert ours == {'SOD-1': -400, 'SOD-2': -50, 'SOD-4': -300, 'SOD-5': 50, 'SOD-6': 200}
        assert drift.moved_only == 1  # SOD-3
        for r in drift.rows:
            assert r.variance == pytest.approx(r.d_input - r.d_output + r.rest)

        ledger.close_period(MAR)
        booked = dict(ledger.query(
            "SELECT SO_ID, SUM(variance_amount) FROM ob_snapshot WHERE snapshot_period = ? "
            "GROUP BY SO_ID HAVING ABS(SUM(variance_amount)) > 0.5", MAR))
        assert booked == ours
        booked_qty = dict(ledger.query(
            "SELECT SO_ID, SUM(variance_qty) FROM ob_snapshot WHERE snapshot_period = ? "
            "GROUP BY SO_ID HAVING ABS(SUM(variance_qty)) > 0.001", MAR))
        assert booked_qty == {r.so_id: r.variance_qty for r in drift.rows
                              if abs(r.variance_qty) > 0.001}

        after = ledger.drift()
        assert (after.period, after.rows) == (MAR, ()), '마감하면 변동은 Variance로 넘어가고 비워진다'


# ═══════════════════════════════════════════════════════════════
# 대시보드 로더
# ═══════════════════════════════════════════════════════════════
class TestDashboardLoader:

    @pytest.fixture
    def load(self, monkeypatch):
        def _load(path):
            monkeypatch.setattr(dashboard, '_conn', lambda: sqlite3.connect(str(path)))
            dashboard.load_post_close_drift.clear()
            try:
                return dashboard.load_post_close_drift()
            finally:
                dashboard.load_post_close_drift.clear()
        return _load

    def test_rows_come_from_the_engine_module(self, tmp_path, load):
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)],
                        dn_dom=[dn('DND-1', 'SOD-1', 1, 4, 400)])
        ledger.close_period(JAN)
        ledger.run("UPDATE dn_domestic SET [세금계산서 발행일] = '2026-01-31' WHERE DN_ID = 'DND-1'")

        got = load(ledger.path)
        assert (got['period'], got['moved_only'], got.get('error')) == (JAN, 0, None)
        assert got['rows'].to_dict('records') == [{
            'so_id': 'SOD-1', 'market': '국내', 'customer': '고객A',
            'snap_ending': 1000, 'live_ending': 600, 'variance': -400,
            'd_input': 0, 'd_output': 400, 'rest': 0, 'variance_qty': -4}]

    def test_no_drift_keeps_the_columns(self, tmp_path, load):
        """0행이어도 열은 있어야 화면 코드가 `rows['variance']`에서 죽지 않는다"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        got = load(ledger.path)
        assert got['rows'].empty and 'variance' in got['rows'].columns

    def test_no_close_is_none(self, tmp_path, load):
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        assert load(ledger.path) is None

    def test_query_failure_is_not_shown_as_no_drift(self, tmp_path, load):
        """조회가 죽었는데 '달라진 것 없음'으로 보이면 안 된다 (검사 실패 ≠ 0건과 같은 이유)"""
        ledger = Ledger(tmp_path, so_dom=[so('SOD-1', 1, 1000)])
        ledger.close_period(JAN)
        ledger.run('DROP VIEW v_dn_by_month')
        got = load(ledger.path)
        assert got['error'] and 'rows' not in got


def test_events_after_close_keeps_only_this_orders_ledger_changes():
    """드릴다운 — 마감 시각 이후, SO/DN 시트, PK 어느 자리에든 그 SO_ID가 있는 이벤트만 (오래된 것부터)"""
    log = pd.DataFrame([
        # id, 시각, 시트, pk_json
        (1, '2026-10-30 09:00:00', 'SO_국내', '["SOD-1","1"]'),              # 마감 전
        (5, '2026-11-05 09:12:36', 'DN_국내', '["DND-9","SOD-1","1","1"]'),  # DN은 SO_ID가 둘째 자리
        (4, '2026-10-30 17:22:29', 'SO_국내', '["SOD-1","2"]'),
        (6, '2026-11-05 09:12:36', 'PO_국내', '["ND-1","SOD-1","1"]'),       # PO는 잔고와 무관
        (7, '2026-11-05 09:12:36', 'SO_국내', '["SOD-10","1"]'),             # 다른 주문 (접두어만 같다)
        (8, '2026-11-06 10:00:00', 'DN_국내', ''),                           # 깨진 PK
    ], columns=['id', 'sync_time', 'sheet_name', 'pk_json'])

    got = dashboard._events_after_close(log, '2026-10-30T13:53:20.836816', 'SOD-1')
    assert got['id'].tolist() == [4, 5]
    assert dashboard._events_after_close(log.iloc[0:0], '2026-10-30T13:53:20', 'SOD-1').empty
