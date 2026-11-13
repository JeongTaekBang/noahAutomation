"""
대시보드 이력 화면 — 미확인 변경 카드 · '주요 변경' 보기 · 주문 타임라인
=====================================================================

담당자가 이력을 볼 때 쓰는 세 화면의 판정 로직. 지키고 싶은 것:

1. **미확인 단가/수량 변동**은 주문 단위로 묶이고, 그 변경이 얼마나 무거운지(미출고 → 출고 →
   매출 인식 → 마감된 달)와 금액이 같이 보인다. 확인은 **원본 DB**에 사유와 함께 남는다.
2. **주요 변경**은 사람이 값을 고친 것만 남긴다 — 빈칸 채움·수식 열·진행 Status는 빠지되,
   취소·보류처럼 잔고를 움직이는 Status 전환은 남는다.
3. **주문 타임라인**은 고객이 부르는 번호(Customer PO)로도 찾아지고, 삭제가 문서째인지 행만인지 구분된다.

Streamlit 렌더링은 테스트하지 않는다 (`tests/test_dashboard.py`와 같은 방침).
"""

from __future__ import annotations

import json
import sqlite3

import pandas as pd
import pytest

import dashboard
from po_generator.db_schema import (
    create_order_book_views, create_snapshot_tables, ensure_sync_log_tables,
)

SO_COLS = ['SO_ID', 'Line item', 'Customer name', 'Customer PO', 'AX Project number',
           'Item name', 'Item qty', 'Sales Unit Price', 'Currency', 'Status']
PO_COLS = ['PO_ID', 'SO_ID', 'Line item', '_row_seq', 'Item name', 'Item qty', 'ICO Unit', 'Status']
DN_DOM_COLS = ['DN_ID', 'SO_ID', 'Line item', '_row_seq', 'Qty', 'Total Sales', '출고일',
               '세금계산서 발행일', '선수금 세금계산서 발행일', 'Customer name', 'Item',
               'Customer PO', 'Business registration number']
DN_EXP_COLS = ['DN_ID', 'SO_ID', 'Line item', '_row_seq', 'Qty', 'Currency', 'Total Sales',
               'Total Sales KRW', '출고일', '선적일', 'Customer name', 'Item', 'Customer PO']
TABLES = (('so_domestic', SO_COLS), ('so_export', SO_COLS),
          ('po_domestic', PO_COLS), ('po_export', PO_COLS),
          ('dn_domestic', DN_DOM_COLS), ('dn_export', DN_EXP_COLS))


def so(so_id, line, qty=1, price=100, **kw):
    row = {'SO_ID': so_id, 'Line item': str(line), 'Customer name': '고객A', 'Customer PO': 'CPO-1',
           'Item name': 'IQ3', 'Item qty': str(qty), 'Sales Unit Price': str(price),
           'Currency': 'KRW', 'Status': '미출고'}
    row.update(kw)
    return row


def po(po_id, so_id, line, **kw):
    row = {'PO_ID': po_id, 'SO_ID': so_id, 'Line item': str(line), '_row_seq': '1',
           'Item name': 'IQ3', 'Item qty': '1', 'ICO Unit': '50', 'Status': 'Confirmed'}
    row.update(kw)
    return row


def dn(dn_id, so_id, line, qty, amount, shipped='2026-11-09', invoice=None, **kw):
    """국내 DN — `invoice`(세금계산서 발행일)가 없으면 출고만 됐고 매출은 미인식"""
    row = {'DN_ID': dn_id, 'SO_ID': so_id, 'Line item': str(line), '_row_seq': '1',
           'Qty': str(qty), 'Total Sales': str(amount), '출고일': shipped,
           '세금계산서 발행일': invoice, 'Customer name': '고객A', 'Item': 'IQ3'}
    row.update(kw)
    return row


def edit(log_id, sync_id, sheet, pk, changes):
    """수정 로그 — changes = {컬럼: (이전, 이후)}"""
    return (log_id, sync_id, sheet, '수정', pk,
            {c: {'old': o, 'new': n} for c, (o, n) in changes.items()}, None)


def added(log_id, sync_id, sheet, pk, values):
    return (log_id, sync_id, sheet, '신규', pk, values, None)


def deleted(log_id, sync_id, sheet, pk, snapshot):
    return (log_id, sync_id, sheet, '삭제', pk, None, snapshot)


def make_db(tmp_path, runs=None, log=(), closed=(), views=True, **tables):
    """대시보드가 읽는 모양의 임시 DB (값은 실DB처럼 TEXT). 경로를 돌려준다.

    runs = {sync_id: 시작 시각}, log = edit()/added()/deleted() 목록, closed = 마감된 달 목록.
    """
    path = tmp_path / 'noah_data.db'
    conn = sqlite3.connect(str(path))
    for table, cols in TABLES:
        conn.execute(f"CREATE TABLE {table} ({', '.join(f'[{c}] TEXT' for c in cols)})")
        for row in tables.get(table, ()):
            unknown = set(row) - set(cols)
            assert not unknown, f'{table}에 없는 컬럼: {unknown}'
            conn.execute(f"INSERT INTO {table} ({', '.join(f'[{c}]' for c in cols)}) "
                         f"VALUES ({', '.join('?' * len(cols))})", [row.get(c) for c in cols])
    if views:
        create_order_book_views(conn)
    create_snapshot_tables(conn)
    for period in closed:
        conn.execute("INSERT INTO ob_snapshot_meta (period, closed_at, is_active) VALUES (?, ?, 1)",
                     (period, f'{period}-28T18:00:00'))
    ensure_sync_log_tables(conn)
    for sync_id, started in (runs or {}).items():
        conn.execute("INSERT INTO _sync_runs (sync_id, started_at, actor, dry_run) VALUES (?, ?, '담당자', 0)",
                     (sync_id, started))
    for log_id, sync_id, sheet, ctype, pk, changes, snapshot in log:
        conn.execute(
            "INSERT INTO _sync_log (id, sync_id, sheet_name, change_type, pk_json, pk_display, "
            "changes_json, row_snapshot_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (log_id, sync_id, sheet, ctype, json.dumps(pk, ensure_ascii=False), ' | '.join(pk),
             json.dumps(changes, ensure_ascii=False) if changes is not None else None,
             json.dumps(snapshot, ensure_ascii=False) if snapshot is not None else None))
    conn.commit()
    conn.close()
    return path


CACHED = ('load_so_unauth_changes', 'load_so_ack_history', 'load_sync_log',
          'load_live_doc_ids', 'resolve_related_ids')


@pytest.fixture
def use_db(monkeypatch):
    """대시보드의 읽기(`_conn`)와 쓰기(`DB_FILE`)를 임시 DB로 돌리고, 캐시를 앞뒤로 비운다"""
    def _clear():
        for name in CACHED:
            getattr(dashboard, name).clear()

    def _use(path):
        monkeypatch.setattr(dashboard, '_conn', lambda: sqlite3.connect(str(path)))
        monkeypatch.setattr(dashboard, 'DB_FILE', path)
        _clear()
        return path
    yield _use
    _clear()


def log_frame(rows):
    """`load_sync_log`가 돌려주는 모양의 DataFrame (순수 함수 테스트용)"""
    records = []
    for log_id, sync_id, sheet, ctype, pk, changes, snapshot in rows:
        records.append({
            'id': log_id, 'sync_id': sync_id, 'sync_time': '2026-11-05 09:00:00', 'actor': '담당자',
            'sheet_name': sheet, 'change_type': ctype, 'pk': ' | '.join(pk),
            'pk_json': json.dumps(pk, ensure_ascii=False),
            'changes_json': json.dumps(changes, ensure_ascii=False) if changes is not None else None,
            'row_snapshot_json': json.dumps(snapshot, ensure_ascii=False) if snapshot is not None else None,
        })
    return pd.DataFrame(records)


# ═══════════════════════════════════════════════════════════════
# 3. 미확인 단가/수량 변동 — 판정
# ═══════════════════════════════════════════════════════════════
class TestWatchedChanges:

    def test_price_and_quantity_value_changes(self):
        got = dashboard._watched_changes({
            'Sales Unit Price': {'old': '7410000', 'new': '7300000'},
            'Item qty': {'old': '4', 'new': '3'},
            'Sales amount': {'old': '1', 'new': '2'},   # 파생 — 감시 대상 아님
        })
        assert got == [('Item qty', '4', '3'), ('Sales Unit Price', '7410000', '7300000')]

    def test_customer_po_change_authorizes_it(self):
        assert dashboard._watched_changes({
            'Customer PO': {'old': 'A', 'new': 'A R1'},
            'Item qty': {'old': '4', 'new': '3'}}) == []

    @pytest.mark.parametrize('old,new', [(None, '3'), ('', '3'), ('3', None), ('3', '3')])
    def test_first_entry_and_no_change_are_ignored(self, old, new):
        assert dashboard._watched_changes({'Item qty': {'old': old, 'new': new}}) == []


class TestAmountImpact:

    def test_price_only_uses_current_quantity(self):
        watched = [('Sales Unit Price', '7410000', '7300000')]
        assert dashboard._amount_impact(watched, so_qty='2', so_price='7300000') == -220000

    def test_quantity_only_uses_current_price(self):
        assert dashboard._amount_impact([('Item qty', '4', '3')], so_qty='3', so_price='200000') == -200000

    def test_both_changed(self):
        watched = [('Item qty', '2', '3'), ('Sales Unit Price', '100', '90')]
        assert dashboard._amount_impact(watched, so_qty='999', so_price='999') == 3 * 90 - 2 * 100

    @pytest.mark.parametrize('watched,qty,price', [
        ([('Sales Unit Price', '100', '90')], None, '100'),    # 단가를 바꿨는데 수량을 모른다
        ([('Sales Unit Price', '100', '90')], '두 개', '100'),
        ([('Item qty', '4', '3')], '3', None),                 # 수량을 바꿨는데 단가를 모른다
    ])
    def test_unknown_stays_unknown(self, watched, qty, price):
        """모르는 값을 0으로 메워 금액을 내지 않는다 — SO 라인이 그 뒤 지워졌을 수 있다"""
        assert dashboard._amount_impact(watched, qty, price) is None

    def test_the_changed_side_needs_no_current_value(self):
        assert dashboard._amount_impact([('Sales Unit Price', '100', '90')], '2', None) == -20


class TestShipmentStage:
    """단계는 **변경(동기화) 시점 기준**이다 — 지금 출고돼 있다고 다 무거운 게 아니다.

    주문을 고친 뒤 그 값으로 출고한 것은 정상 업무다. 무거운 것은 이미 나간 뒤에 고친 것,
    그중에서도 마감이 굳은 뒤에 고친 것이다.
    """
    CHANGED = '2026-11-12 13:03:14'
    CLOSED = {'2026-10': '2026-10-30T13:53:20.836816'}
    SHIPPED = {'dn_id': 'DND-1', 'shipped': '2026-11-09', 'month': None}
    BILLED = {'dn_id': 'DND-2', 'shipped': '2026-11-02', 'month': '2026-11'}
    IN_CLOSED_MONTH = {'dn_id': 'DND-3', 'shipped': '2026-10-20', 'month': '2026-10'}

    def _stage(self, rows, changed=CHANGED, closed=CLOSED):
        return dashboard._shipment_stage(rows, changed, closed)

    def test_four_stages_get_heavier(self):
        ranks = [self._stage(rows)[0] for rows in
                 ([], [self.SHIPPED], [self.BILLED], [self.BILLED, self.IN_CLOSED_MONTH])]
        assert ranks == sorted(ranks) and len(set(ranks)) == 4

    def test_labels_say_what_happened(self):
        assert self._stage([])[1] == '미출고'
        label = self._stage([self.SHIPPED])[1]
        assert '출고 뒤 변경' in label and '2026-11-09' in label and '매출 미인식' in label
        assert '매출 인식(2026-11)' in self._stage([self.BILLED])[1]
        label = self._stage([self.IN_CLOSED_MONTH])[1]
        assert '마감 뒤 변경' in label and '마감된 달(2026-10)' in label

    def test_shipped_after_the_change_is_ordinary_business(self):
        """실측 `SOD-2026-1102`: 10/30에 수량을 고치고 11/11에 그 수량으로 출고했다"""
        rows = [{'dn_id': 'DND-1053', 'shipped': '2026-11-11', 'month': '2026-11'}]
        rank, label = self._stage(rows, changed='2026-10-30 17:22:29')
        assert rank == self._stage([])[0] and '변경 뒤 출고' in label and '2026-11-11' in label

    def test_changed_before_the_close_is_already_in_the_snapshot(self):
        rank, label = self._stage([self.IN_CLOSED_MONTH], changed='2026-10-25 10:00:00')
        assert rank == self._stage([self.BILLED])[0] and '마감 뒤 변경' not in label

    def test_close_time_is_compared_as_time_not_text(self):
        """마감 시각은 'T' 구분, 동기화 시각은 공백 구분 — 문자열로 비교하면 같은 날 17시가 13시보다 앞선다"""
        rank, _ = self._stage([self.IN_CLOSED_MONTH], changed='2026-10-30 17:22:29')
        assert rank == self._stage([self.IN_CLOSED_MONTH])[0], '10/30 13:53 마감 뒤인 17:22 변경'

    def test_no_close_yet_means_recognized_at_most(self):
        assert '매출 인식' in self._stage([self.IN_CLOSED_MONTH], closed={})[1]

    def test_shipment_without_a_date_counts_as_shipped(self):
        rank, label = self._stage([{'dn_id': 'DND-1', 'shipped': None, 'month': None}])
        assert rank == self._stage([self.SHIPPED])[0] and '출고일 미입력' in label

    def test_malformed_month_is_not_a_recognition(self):
        """뷰의 해외 매출월은 `SUBSTR(선적일, 1, 7)` — 선적일이 빈 문자열이면 ''이 온다"""
        row = {'dn_id': 'DNO-1', 'shipped': '2026-11-09', 'month': ''}
        assert self._stage([row]) == self._stage([{**row, 'month': None}])

    def test_lookup_failure_is_not_unshipped(self):
        """DN을 못 읽었는데 '미출고'로 보이면 가장 가벼운 단계로 정렬된다 — 모른다고 한다"""
        rank, label = self._stage(None)
        assert rank < 0 and '미출고' not in label


# ═══════════════════════════════════════════════════════════════
# 3. 미확인 단가/수량 변동 — 로더 · 묶기 · 확인
# ═══════════════════════════════════════════════════════════════
class TestUnackedChanges:

    @pytest.fixture
    def db(self, tmp_path, use_db):
        """실측을 줄인 모양: 한 주문 3라인 단가 인하(출고 뒤) + 다른 주문들"""
        return use_db(make_db(
            tmp_path,
            so_domestic=[
                so('SOD-0333', 1, qty=1, price=10170000, **{'Customer name': '(주)삼신', 'Customer PO': 'SS-1'}),
                so('SOD-0333', 2, qty=1, price=33040000, **{'Customer name': '(주)삼신', 'Customer PO': 'SS-1'}),
                so('SOD-1068', 1, qty=3, price=200000),
                so('SOD-0900', 1, qty=1, price=500000),
                so('SOD-1116', 4, qty=1, price=2100000)],
            so_export=[so('SOO-0001', 1, qty=2, price=90, Currency='EUR', **{'Customer name': 'ACME'})],
            dn_domestic=[
                dn('DND-1032', 'SOD-0333', 1, 1, 10170000, shipped='2026-11-09'),
                dn('DND-1032', 'SOD-0333', 2, 1, 33040000, shipped='2026-11-09'),
                dn('DND-1002', 'SOD-1068', 1, 3, 600000, shipped='2026-11-02', invoice='2026-11-02'),
                dn('DND-0950', 'SOD-0900', 1, 1, 500000, shipped='2026-10-20', invoice='2026-10-31')],
            runs={239: '2026-10-30 17:22:29', 240: '2026-11-05 09:12:36', 242: '2026-11-12 13:03:14'},
            log=[
                edit(1, 239, 'SO_국내', ['SOD-1116', '4'], {'Item qty': ('2', '1')}),
                edit(2, 240, 'SO_국내', ['SOD-1068', '1'], {'Item qty': ('4', '3')}),
                edit(3, 242, 'SO_국내', ['SOD-0333', '1'], {'Sales Unit Price': ('11670000', '10170000')}),
                edit(4, 242, 'SO_국내', ['SOD-0333', '2'], {'Sales Unit Price': ('34540000', '33040000')}),
                edit(5, 242, 'SO_국내', ['SOD-0900', '1'], {'Sales Unit Price': ('400000', '500000')}),
                edit(6, 242, 'SO_해외', ['SOO-0001', '1'], {'Sales Unit Price': ('100', '90')}),
                edit(7, 242, 'SO_국내', ['SOD-1068', '1'], {'Customer PO': ('A', 'B'), 'Item qty': ('3', '5')}),
                edit(8, 242, 'SO_국내', ['SOD-1116', '4'], {'Status': ('미출고', '출고 완료')}),
            ],
            closed=['2026-10']))

    def test_lines_carry_stage_and_amount(self, db):
        df = dashboard.load_so_unauth_changes().set_index('sync_log_id')
        assert set(df.index) == {1, 2, 3, 4, 5, 6}, 'Customer PO가 같이 바뀐 것·감시 필드가 아닌 것은 빠진다'

        assert df.loc[1, 'stage'] == '미출고' and df.loc[1, 'impact'] == -2100000
        assert '출고 뒤 변경' in df.loc[3, 'stage'] and df.loc[3, 'impact'] == -1500000
        assert '매출 인식' in df.loc[2, 'stage'] and df.loc[2, 'impact'] == -200000
        assert '마감된 달(2026-10)' in df.loc[5, 'stage'] and df.loc[5, 'impact'] == 100000
        assert list(df.loc[[1, 3, 2, 5], 'stage_rank']) == sorted(df.loc[[1, 3, 2, 5], 'stage_rank'])

        assert (df.loc[3, 'customer'], df.loc[3, 'customer_po'], df.loc[3, 'line']) == ('(주)삼신', 'SS-1', 1)
        assert (df.loc[6, 'currency'], df.loc[6, 'impact'], df.loc[6, 'customer']) == ('EUR', -20, 'ACME')
        assert set(df['so_state']) == {''}

    def test_amount_uses_the_other_value_as_it_was_then(self, tmp_path, use_db):
        """같은 라인이 여러 번 바뀌면 — 수량 5→7, 단가 100→120, 수량 7→9 — 각 변경은 그때 값으로 잰다

        전부 지금 값(9개·120원)으로 재면 240 + 180 + 240 = 660인데, 실제로 달라진 금액은
        9×120 − 5×100 = 580이다 (200 + 140 + 240). 뒤 변경을 이미 확인했어도 마찬가지다.
        """
        use_db(make_db(
            tmp_path, so_domestic=[so('SOD-1', 1, qty=9, price=120)],
            runs={1: '2026-11-01 09:00:00', 2: '2026-11-02 09:00:00', 3: '2026-11-03 09:00:00'},
            log=[edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('5', '7')}),
                 edit(2, 2, 'SO_국내', ['SOD-1', '1'], {'Sales Unit Price': ('100', '120')}),
                 edit(3, 3, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('7', '9')})]))
        impact = dashboard.load_so_unauth_changes().set_index('sync_log_id')['impact']
        assert impact.to_dict() == {1: 200, 2: 140, 3: 240}
        assert impact.sum() == 9 * 120 - 5 * 100

        assert dashboard._ack_so_changes([2, 3], None)
        dashboard.load_so_unauth_changes.clear()
        assert dashboard.load_so_unauth_changes().set_index('sync_log_id')['impact'].to_dict() == {1: 200}

    def test_cancelled_or_missing_line_is_flagged(self, tmp_path, use_db):
        use_db(make_db(
            tmp_path, so_domestic=[so('SOD-1', 1, qty=5, price=100, Status='Cancelled')],
            runs={1: '2026-11-01 09:00:00'},
            log=[edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('3', '5')}),
                 edit(2, 1, 'SO_국내', ['SOD-1', '2'], {'Item qty': ('3', '5')})]))
        state = dashboard.load_so_unauth_changes().set_index('sync_log_id')['so_state']
        assert state[1] == 'Cancelled' and '없음' in state[2]

    def test_one_card_per_order_and_sync(self, db):
        groups = dashboard._group_so_changes(dashboard.load_so_unauth_changes())
        assert [(g['SO_ID'], g['log_ids']) for g in groups] == [
            ('SOD-0900', [5]),        # 마감된 달 매출 — 가장 무겁다
            ('SOD-1068', [2]),        # 매출 인식됨
            ('SOD-0333', [3, 4]),     # 출고됨 — 한 주문의 두 라인이 카드 하나
            ('SOO-0001', [6]),        # 미출고, 최신
            ('SOD-1116', [1]),        # 미출고, 오래된 것
        ]
        samsin = groups[2]
        assert samsin['impact'] == -3000000 and samsin['customer'] == '(주)삼신'
        assert len({g['key'] for g in groups}) == len(groups), '위젯 키로 쓴다 — 겹치면 안 된다'

    def test_same_order_changed_in_two_syncs_is_two_cards(self, tmp_path, use_db):
        """동기화가 다르면 다른 사건이다 — 한쪽만 검토하고 둘 다 확인되면 안 된다"""
        use_db(make_db(
            tmp_path, so_domestic=[so('SOD-1', 1, qty=5, price=100)],
            runs={1: '2026-11-01 09:00:00', 2: '2026-11-02 09:00:00'},
            log=[edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('3', '4')}),
                 edit(2, 2, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('4', '5')})]))
        groups = dashboard._group_so_changes(dashboard.load_so_unauth_changes())
        assert [g['log_ids'] for g in groups] == [[2], [1]]

    def test_unknown_amount_keeps_the_group_total_unknown(self, tmp_path, use_db):
        """한 라인이라도 금액을 모르면 묶음 합계를 내지 않는다 (아는 것만 더한 숫자는 틀린 숫자다)"""
        use_db(make_db(
            tmp_path, so_domestic=[so('SOD-1', 1, qty=5, price=100)],   # 2번 라인은 그 뒤 지워졌다
            runs={1: '2026-11-01 09:00:00'},
            log=[edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('3', '5')}),
                 edit(2, 1, 'SO_국내', ['SOD-1', '2'], {'Sales Unit Price': ('10', '20')})]))
        df = dashboard.load_so_unauth_changes().set_index('sync_log_id')
        assert df.loc[1, 'impact'] == 200 and pd.isna(df.loc[2, 'impact'])
        assert dashboard._group_so_changes(df.reset_index())[0]['impact'] is None

    def test_context_failure_still_lists_the_change(self, tmp_path, use_db):
        """맥락을 못 붙여도(뷰 없음) 변경 자체는 보여야 한다 — 이 카드는 한 번이라도 놓치면 안 되는 것"""
        use_db(make_db(
            tmp_path, views=False, so_domestic=[so('SOD-1', 1, qty=5, price=100)],
            runs={1: '2026-11-01 09:00:00'},
            log=[edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('3', '5')})]))
        df = dashboard.load_so_unauth_changes()
        assert df['sync_log_id'].tolist() == [1]
        assert df.loc[0, 'stage_rank'] < 0 and '미출고' not in df.loc[0, 'stage']

    # ── 확인(ack) — 쓰기는 원본 DB로 ──
    @staticmethod
    def _acks(path):
        with sqlite3.connect(str(path)) as conn:
            return conn.execute(
                "SELECT sync_log_id, note FROM _so_change_ack ORDER BY sync_log_id").fetchall()

    def test_group_ack_writes_every_line_with_the_note(self, db):
        assert dashboard._ack_so_changes([3, 4], '  고객 요청으로 단가 조정  ') is True
        assert self._acks(db) == [(3, '고객 요청으로 단가 조정'), (4, '고객 요청으로 단가 조정')]
        dashboard.load_so_unauth_changes.clear()
        assert set(dashboard.load_so_unauth_changes()['sync_log_id']) == {1, 2, 5, 6}

    def test_blank_note_is_stored_as_null(self, db):
        assert dashboard._ack_so_changes([1], '   ') is True
        assert self._acks(db) == [(1, None)]

    def test_group_ack_is_all_or_nothing(self, db):
        """묶음 중 하나라도 못 쓰면 전부 취소 — 반만 확인된 카드는 다음에 반만 남아 헷갈린다"""
        assert dashboard._ack_so_changes([3, 9999], '사유') is False
        assert self._acks(db) == []

    def test_nothing_to_ack(self, db):
        assert dashboard._ack_so_changes([], '사유') is False

    def test_history_shows_who_when_why(self, db):
        dashboard._ack_so_changes([3, 4], '고객 요청')
        dashboard._ack_so_changes([1], None)
        dashboard.load_so_ack_history.clear()
        hist = dashboard.load_so_ack_history()
        assert len(hist) == 3
        row = hist[hist['sync_log_id'] == 3].iloc[0]
        assert row['사유'] == '고객 요청' and row['주문·라인'] == 'SOD-0333 | 1'
        assert row['변경'] == 'Sales Unit Price: 11670000 → 10170000'
        assert row['변경 시각'].startswith('2026-11-12 13:03')
        assert hist['확인 시각'].is_monotonic_decreasing, '최신 확인이 위'

    def test_history_without_any_ack(self, db):
        assert dashboard.load_so_ack_history().empty


# ═══════════════════════════════════════════════════════════════
# 4·5. 탐색 — 주요 변경
# ═══════════════════════════════════════════════════════════════
class TestMajorChanges:

    @staticmethod
    def _kept(frame, scope='전부'):
        """{id: 남은 컬럼 목록}"""
        out = dashboard._major_changes(frame, scope)
        return {int(i): (sorted(json.loads(c)) if c else None)
                for i, c in zip(out['id'], out['changes_json'])}

    def test_first_entry_and_formula_columns_drop_out(self):
        frame = log_frame([
            edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'AX Project number': (None, '82-1.001')}),   # 빈칸 채움
            edit(2, 1, 'SO_국내', ['SOD-1', '1'], {'Status': ('미출고', '출고 완료'),            # 파생
                                                    'EXW NOAH': ('2026-11-01', '2026-11-08'),   # 수식
                                                    'Seq': ('10', '11')}),
            edit(3, 1, 'SO_국내', ['SOD-1', '1'], {'Sales Unit Price': ('100', '90'),
                                                    'Sales amount': ('100', '90')}),            # 금액은 수식
            edit(4, 1, 'PO_국내', ['ND-1', '1', '1'], {'Status': ('Confirmed', 'Invoiced P10')}),  # 진행 표시
        ])
        assert self._kept(frame) == {3: ['Sales Unit Price']}

    def test_same_column_name_is_judged_per_sheet(self):
        """`Customer name`은 SO에서는 사람이 넣고 DN에서는 수식이 따라온다"""
        frame = log_frame([
            edit(1, 1, 'SO_해외', ['SOO-1', '1'], {'Customer name': ('ASC', 'Reliable')}),
            edit(2, 1, 'DN_해외', ['DNO-1', 'SOO-1', '1', '1'], {'Customer name': ('ASC', 'Reliable')}),
        ])
        assert self._kept(frame) == {1: ['Customer name']}

    @pytest.mark.parametrize('sheet,old,new', [
        ('SO_국내', '미출고', 'Cancelled'), ('SO_국내', 'Hold', '미출고'),
        ('PO_국내', 'Confirmed', 'Cancelled'), ('PO_해외', 'Holding', 'Confirmed'),
        ('SO_국내', None, 'Cancelled'), ('SO_국내', 'Cancelled', None),   # 빈칸이 끼어도 취소는 취소다
    ])
    def test_cancel_and_hold_status_changes_stay(self, sheet, old, new):
        """Status는 대부분 진행 표시지만 취소·보류는 잔고를 움직인다 (90일 실측 93건)"""
        frame = log_frame([edit(1, 1, sheet, ['X-1', '1'], {'Status': (old, new)})])
        assert self._kept(frame) == {1: ['Status']}
        assert self._kept(frame, '취소·보류') == {1: ['Status']}

    def test_erasing_a_value_stays_but_first_entry_does_not(self):
        """값을 지운 것은 최초 입력과 다르다 — 실측: 해외 선적일을 지운 15건(90일)은 매출 인식 취소"""
        frame = log_frame([
            edit(1, 1, 'DN_해외', ['DNO-0184', 'SOO-0207', '1', '1'], {'선적일': ('2026-09-27 00:00:00', None)}),
            edit(2, 1, 'DN_해외', ['DNO-0185', 'SOO-0208', '1', '1'], {'선적일': (None, '2026-09-27 00:00:00')}),
            edit(3, 1, 'DN_국내', ['DND-1', 'SOD-1', '1', '1'], {'세금계산서 발행일': ('2026-10-31', '')}),
        ])
        assert self._kept(frame) == {1: ['선적일'], 3: ['세금계산서 발행일']}
        assert self._kept(frame, '출고·매출 인식') == {1: ['선적일'], 3: ['세금계산서 발행일']}

    @pytest.mark.parametrize('old,new', [
        ('2026-11-15', '2026-11-15 00:00:00'),   # 실측: 한 동기화에서 요청납기 18건 — 셀 서식만 바뀌었다
        ('7300000', '7300000.0'), ('고객 A ', '고객 A'),
    ])
    def test_same_value_in_another_notation_is_not_a_change(self, old, new):
        frame = log_frame([edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Requested delivery date': (old, new)})])
        assert self._kept(frame) == {}

    def test_a_real_date_or_number_change_is_still_a_change(self):
        frame = log_frame([edit(1, 1, 'SO_국내', ['SOD-1', '1'], {
            'Requested delivery date': ('2027-01-15 00:00:00', '2027-02-01 00:00:00'),
            'Item qty': ('4', '3.0')})])
        assert self._kept(frame) == {1: ['Item qty', 'Requested delivery date']}

    def test_dn_amount_columns_stay_visible(self):
        """DN 단가·금액은 수식 열이지만 숨기지 않는다 — 이미 나간 매출이 달라진 것이고 드물다(90일 2건).

        숨기면 사람이 수식을 값으로 덮어쓴 경우를 '단가 불일치' 검사의 문턱(1원·0.1%) 아래에서 놓친다.
        """
        frame = log_frame([edit(1, 1, 'DN_국내', ['DND-1', 'SOD-1', '1', '1'], {
            'Unit Price': ('7410000', '7300000'), 'Total Sales': ('7410000', '7300000'),
            'Customer name': ('A', 'B'), '_row_seq': ('1', '2')})])
        assert self._kept(frame) == {1: ['Total Sales', 'Unit Price']}

    def test_status_of_other_sheets_is_ordinary_input(self):
        frame = log_frame([edit(1, 1, 'PMT_국내', ['ADV-1'], {'Status': ('입금대기', '입금완료')})])
        assert self._kept(frame) == {1: ['Status']}
        assert self._kept(frame, '취소·보류') == {}

    def test_new_rows_drop_and_deletes_stay_only_in_everything(self):
        frame = log_frame([
            added(1, 1, 'SO_국내', ['SOD-9', '1'], {'Item qty': '3'}),
            deleted(2, 1, 'SO_국내', ['SOD-8', '1'], {'Item qty': '3'}),
            edit(3, 1, 'PO_국내', ['ND-1', '1', '1'], {'ICO Unit': ('175576', '441818')}),
        ])
        assert self._kept(frame) == {2: None, 3: ['ICO Unit']}
        assert self._kept(frame, '매입 수량·단가') == {3: ['ICO Unit']}

    def test_scopes_pick_sheets_and_columns(self):
        frame = log_frame([
            edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('4', '3'), 'Remarks': ('a', 'b')}),
            edit(2, 1, 'PO_국내', ['ND-1', '1', '1'], {'Item qty': ('4', '3'), 'ICO Unit': ('10', '12')}),
            edit(3, 1, 'PO_국내', ['ND-1', '1', '1'], {'공장 EXW date': ('2026-11-01', '2026-11-20')}),
            edit(4, 1, 'DN_국내', ['DND-1', 'SOD-1', '1', '1'], {'출고일': ('2026-11-01', '2026-10-27')}),
            edit(5, 1, 'SO_국내', ['SOD-1', '1'], {'Customer PO': ('SP-005C', 'SP-005C R3')}),
        ])
        assert self._kept(frame, '판매 수량·단가') == {1: ['Item qty']}
        assert self._kept(frame, '매입 수량·단가') == {2: ['ICO Unit', 'Item qty']}
        assert self._kept(frame, '납기·일정') == {3: ['공장 EXW date']}
        assert self._kept(frame, '출고·매출 인식') == {4: ['출고일']}
        assert self._kept(frame, '거래처·주문번호') == {5: ['Customer PO']}
        assert self._kept(frame)[1] == ['Item qty', 'Remarks'], "'전부'에서는 프리셋에 없는 열도 보인다"

    def test_changes_json_is_narrowed_for_downstream_views(self):
        frame = log_frame([edit(1, 1, 'DN_국내', ['DND-1', 'SOD-1', '1', '1'],
                                {'출고일': ('2026-11-01', '2026-10-27'), 'Customer name': ('A', 'B')})])
        out = dashboard._major_changes(frame, '전부')
        narrowed = out['changes_json'].iloc[0]
        assert json.loads(narrowed) == {'출고일': {'old': '2026-11-01', 'new': '2026-10-27'}}
        assert '출고일' in narrowed, '한글을 \\uXXXX로 바꾸면 화면의 컬럼명 검색이 한글 열을 못 찾는다'
        assert json.loads(frame['changes_json'].iloc[0]).keys() == {'출고일', 'Customer name'}, \
            '원본(캐시된 로그)은 건드리지 않는다'

    def test_broken_or_empty_log_does_not_crash(self):
        frame = log_frame([edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('4', '3')})])
        frame.loc[0, 'changes_json'] = '{깨진'
        assert dashboard._major_changes(frame, '전부').empty
        empty = dashboard._major_changes(frame.iloc[0:0], '전부')
        assert empty.empty and list(empty.columns) == list(frame.columns)

    def test_every_scope_is_usable(self):
        frame = log_frame([edit(1, 1, 'SO_국내', ['SOD-1', '1'], {'Item qty': ('4', '3')})])
        for scope in dashboard._MAJOR_SCOPES:
            dashboard._major_changes(frame, scope)
        assert list(dashboard._MAJOR_SCOPES)[0] == '전부'

    @pytest.mark.parametrize('old,new,want', [
        ('175576', '441818', 151.6), ('100', '90', -10.0), ('0', '804483', None),
        ('SP-005C', 'SP-005C R3', None), ('2026-11-01 00:00:00', '2026-10-27 00:00:00', None),
    ])
    def test_percent_change_only_for_numbers(self, old, new, want):
        got = dashboard._pct_change(old, new)
        assert got == want if want is None else got == pytest.approx(want, abs=0.05)

    def test_detail_rows_one_per_changed_column_and_one_per_delete(self):
        frame = dashboard._major_changes(log_frame([
            edit(1, 1, 'PO_국내', ['ND-1006', '2', '1'], {'ICO Unit': ('175576', '441818'),
                                                          'Item qty': ('1', '2')}),
            deleted(2, 1, 'SO_국내', ['SOD-0933', '1'], {'Customer name': '굿이엔지', 'Item name': 'NA050',
                                                         'Item qty': '3', 'Sales amount': '1671000'}),
        ]), '전부')
        detail = dashboard._major_detail(frame, live_ids={'SO_국내': frozenset()})
        assert len(detail) == 3
        ico = detail[detail['컬럼'] == 'ICO Unit'].iloc[0]
        assert (ico['이전값'], ico['변경값']) == ('175576', '441818')
        assert ico['변동률(%)'] == pytest.approx(151.6, abs=0.05)
        gone = detail[detail['유형'] == '삭제'].iloc[0]
        assert gone['컬럼'] == '문서째 삭제' and '굿이엔지' in gone['이전값'] and 'NA050' in gone['이전값']
        assert detail['변동률(%)'].dtype == float and pd.isna(gone['변동률(%)']), \
            "숫자 열이어야 정렬되고, 빈 칸에 'None'이 찍히지 않는다"
        assert dashboard._major_detail(frame.iloc[0:0])['변동률(%)'].dtype == float


# ═══════════════════════════════════════════════════════════════
# 6. 주문 타임라인
# ═══════════════════════════════════════════════════════════════
class TestOrderLookup:

    @pytest.fixture
    def db(self, tmp_path, use_db):
        return use_db(make_db(
            tmp_path,
            so_domestic=[
                so('SOD-1', 1, **{'Customer PO': 'SP2602-005C R3', 'AX Project number': '82-15689.001'}),
                so('SOD-2', 1, **{'Customer PO': 'WG-77'}), so('SOD-3', 1, **{'Customer PO': 'WG-77'}),
                so('SOD-4', 1, **{'Customer PO': 'OTHER'})],
            po_domestic=[po('ND-1', 'SOD-1', 1), po('ND-2', 'SOD-2', 1), po('ND-2', 'SOD-3', 1)],
            dn_domestic=[dn('DND-1', 'SOD-1', 1, 1, 100), dn('DND-7', 'SOD-2', 1, 1, 100)]))

    def test_by_document_number_as_before(self, db):
        assert dashboard.resolve_related_ids('ND-1') == ['DND-1', 'ND-1', 'SOD-1']
        assert dashboard.resolve_related_ids(' DND-7 ') == ['DND-7', 'ND-2', 'SOD-2', 'SOD-3']

    def test_by_customer_po(self, db):
        """고객이 부르는 번호로 찾는다 — 대소문자·앞뒤 공백은 따지지 않는다"""
        assert dashboard.resolve_related_ids('sp2602-005c r3') == ['DND-1', 'ND-1', 'SOD-1']
        assert dashboard.resolve_related_ids(' WG-77') == ['DND-7', 'ND-2', 'SOD-2', 'SOD-3']

    def test_by_ax_project_number(self, db):
        assert dashboard.resolve_related_ids('82-15689.001') == ['DND-1', 'ND-1', 'SOD-1']

    def test_customer_po_itself_is_not_listed_as_an_id(self, db):
        assert 'WG-77' not in dashboard.resolve_related_ids('WG-77')

    def test_unknown_input_is_kept_for_log_matching(self, db):
        """지금 시트에 없는 번호(지워진 주문)도 로그에는 남아 있다 — 입력을 버리면 그 이력을 못 찾는다"""
        assert dashboard.resolve_related_ids('SOD-0933') == ['SOD-0933']
        assert dashboard.resolve_related_ids('   ') == []

    def test_partial_customer_po_does_not_match(self, db):
        assert dashboard.resolve_related_ids('WG') == ['WG']

    def test_only_document_numbers_are_matched_against_the_log(self):
        """PK에는 라인 번호·행 순번도 들어 있다 — 입력 '10'이 모든 문서의 10번 라인에 걸리면 안 된다"""
        assert dashboard._pk_doc_ids('["SOD-1","10"]') == ['SOD-1']
        assert dashboard._pk_doc_ids('["DND-1","SOD-1","1.0","2"]') == ['DND-1', 'SOD-1']
        assert dashboard._pk_doc_ids('["EUR","2026-11"]') == ['EUR', '2026-11']
        assert dashboard._pk_doc_ids('') == [] and dashboard._pk_doc_ids('{깨진') == []


class TestDeleteKind:

    def test_whole_document_vs_rows_only(self, tmp_path, use_db):
        use_db(make_db(tmp_path, so_domestic=[so('SOD-1', 1)], po_domestic=[po('ND-1', 'SOD-1', 1)],
                       dn_domestic=[dn('DND-1', 'SOD-1', 1, 1, 100)]))
        live = dashboard.load_live_doc_ids()
        kind = dashboard._delete_kind
        assert kind('SO_국내', '["SOD-1","2"]', live) == '행 삭제'
        assert kind('SO_국내', '["SOD-0933","1"]', live) == '문서째 삭제'
        assert kind('DN_국내', '["DND-1","SOD-1","1","1"]', live) == '행 삭제'
        assert kind('DN_국내', '["DND-9","SOD-1","1","1"]', live) == '문서째 삭제', 'DN은 DN 번호로 본다'
        assert kind('PO_국내', '["ND-1","1","1"]', live) == '행 삭제'

    def test_unknown_sheet_or_broken_key_is_just_a_delete(self, tmp_path, use_db):
        use_db(make_db(tmp_path))
        live = dashboard.load_live_doc_ids()
        assert dashboard._delete_kind('FX', '["EUR","2026-11"]', live) == '삭제'
        assert dashboard._delete_kind('SO_국내', '', live) == '삭제'
        assert dashboard._delete_kind('SO_국내', '["SOD-1","1"]', None) == '삭제'


class TestCardSummary:

    def test_new_row_shows_what_was_ordered_not_just_ids(self):
        """예전엔 딕셔너리 앞 3개 — PO 신규는 `PO_ID, SO_ID, Order type`만 보였다"""
        row = log_frame([added(1, 1, 'PO_국내', ['ND-1173', '1', '1'], {
            'PO_ID': 'ND-1173', 'SO_ID': 'SOD-1173', 'Order type': 'Product Sales',
            'Item name': 'NA050', 'Item qty': '3', 'ICO Unit': '512336', 'Model': 'NA'})]).iloc[0]
        body, detail = dashboard._timeline_event_summary(row)
        assert 'NA050' in body and '512336' in body and 'Order type' not in body
        assert list(detail.columns) == ['컬럼', '값'] and len(detail) == 7, '전체 보기에는 전부'

    def test_edit_puts_the_real_change_first(self):
        row = log_frame([edit(1, 1, 'SO_국내', ['SOD-1', '1'], {
            'Seq': ('1', '2'), 'Status': ('미출고', '출고 완료'), 'EXW NOAH': ('a', 'b'),
            'AX Period': (None, 'P11'), 'Sales Unit Price': ('7410000', '7300000')})]).iloc[0]
        body, detail = dashboard._timeline_event_summary(row, max_inline=1)
        assert 'Sales Unit Price' in body and 'Seq' not in body
        assert detail['컬럼'].iloc[0] == 'Sales Unit Price'
        assert list(detail.columns) == ['컬럼', '이전값', '변경값'] and len(detail) == 5

    def test_delete_says_which_kind(self):
        row = log_frame([deleted(1, 1, 'SO_국내', ['SOD-0933', '1'],
                                 {'Customer name': '굿이엔지', 'Item name': 'NA050'})]).iloc[0]
        body, _ = dashboard._timeline_event_summary(row, live_ids={'SO_국내': frozenset()})
        assert '문서째 삭제' in body and '굿이엔지' in body
        body, _ = dashboard._timeline_event_summary(row, live_ids={'SO_국내': frozenset({'SOD-0933'})})
        assert '행 삭제' in body

    def test_table_summary_uses_the_same_words(self):
        row = log_frame([deleted(1, 1, 'SO_국내', ['SOD-0933', '1'], {'Item name': 'NA050'})]).iloc[0]
        assert dashboard._changes_summary(row, {'SO_국내': frozenset()}).startswith('문서째 삭제')
        assert dashboard._changes_summary(row, {'SO_국내': frozenset({'SOD-0933'})}).startswith('행 삭제')
        assert dashboard._changes_summary(row).startswith('삭제'), '문서 목록이 없으면 구분하지 않는다'

    def test_reload_marker_is_not_drawn_as_a_delete(self):
        """'재적재'는 PK 정의가 바뀌어 테이블을 다시 실은 표시다 — 예전엔 삭제 카드로 그려졌다"""
        row = pd.Series({'sheet_name': 'DN_국내', 'change_type': '재적재', 'pk': 'dn_domestic',
                         'pk_json': '["dn_domestic"]', 'row_snapshot_json': None,
                         'changes_json': '{"reason": "PK 정의 변경 → 테이블 재생성", "rows": 1370}'})
        body, _ = dashboard._timeline_event_summary(row)
        assert '재적재' in body and '삭제' not in body
        assert dashboard._changes_summary(row).startswith('재적재')

    def test_key_values_fall_back_when_no_key_field(self):
        assert dashboard._key_values({'Turns': '12', 'ALS': 'Y'}) == [('Turns', '12'), ('ALS', 'Y')]
        assert dashboard._key_values({'Item name': 'NA050', 'Turns': '12', 'Item qty': None}) == [('Item name', 'NA050')]
