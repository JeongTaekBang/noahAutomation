"""
AX 오더북 저장소 — 우리 판 'Table 50229'(AX Sales Order Book)
============================================================

AX는 오더북을 월말에 한 번 계산해 **저장**하고, 다음 달은 저장된 End에서 시작한다(Record info: Table 50229,
'Auto-report', Created by·시각, Last calculated, Period status — 2026-11-03 사용자 확인). 그래서 지난 달 숫자는
바뀌지 않고, 그 뒤의 취소·변경·청구는 일어난 달로 들어간다. 같은 일을 SQLite 파일 하나로 한다.

    periods  기간마다 한 행 — 상태(Open/Closed)·계산 방식·계산 시각·계산한 사람·쓴 IR/IT 파일·추출일·행 수
    rows     그 달 Raw 행(48열 + 키) — 결과 파일 Raw 시트가 이것
    billed   그 달 말까지 Output으로 인식한 청구 누계 — (프로젝트|라인번호|품목번호)별 수량·금액(거래통화).
             다음 달 Output = 그때 청구 누계 − 이것. 마지막 달과 그 앞 달 것만 남긴다(마지막 달을 다시 계산할 때 쓰는 만큼)

기간 규칙: 비어 있으면 기준선(P01~Pm 한 번에), 그 뒤엔 다음 달을 차례로. 마지막 달만 Open이라 다시 계산할 수 있고
(Last calculated), 앞 달은 Closed. 계산은 ax_orderbook.py가 하고(저장소를 모른다) 여기는 읽기·쓰기와 규칙만이다 —
dn_recorder(계산)/dn_writer(쓰기)를 나눈 것과 같은 이유.
OneDrive 폴더라 WAL을 쓰지 않는다 — 사이드카(-wal)를 따로 올려 커밋이 빠진 사본이 생길 수 있다
(CLAUDE.md 'OneDrive로 공유하는 SQLite는 원본을 열지 않는다').
"""

from __future__ import annotations

import getpass
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd

from po_generator.ax_orderbook import BILLED_COLUMNS, NUM_COLUMNS, RAW_COLUMNS

OPEN, CLOSED = 'Open', 'Closed'
BASELINE, MONTHLY = '기준선', '월 계산'
DATE_COLUMNS: tuple[str, ...] = ('Report Start', 'Receipt date', 'Project creation date', 'Requested ship date')
_ROW_COLUMNS: tuple[str, ...] = ('year', 'pn', '_key') + RAW_COLUMNS
_PERIOD_COLUMNS: tuple[str, ...] = ('year', 'pn', 'status', 'basis', 'calculated_at', 'calculated_by', 'ir_file',
                                    'it_file', 'snapshot', 'row_count', 'billed_rows')


class PeriodError(ValueError):
    """기간 규칙 위반 — 닫힌 달을 다시 계산하거나 달을 건너뛰려 할 때, 저장 상태가 모자랄 때"""


@dataclass(frozen=True)
class Plan:
    kind: str                           # 'baseline' | 'next' | 'recompute'
    prev: tuple[int, int] | None        # 시작할 저장된 달 (연, 월) — 기준선이면 None


def _ord(year: int, pn: int) -> int:
    return year * 100 + pn


def previous_period(year: int, pn: int) -> tuple[int, int]:
    return (year - 1, 12) if pn == 1 else (year, pn - 1)


def next_period(year: int, pn: int) -> tuple[int, int]:
    return (year + 1, 1) if pn == 12 else (year, pn + 1)


def _label(period: tuple[int, int]) -> str:
    return f"{period[0]} P{period[1]:02d}"


def _quote(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _sql_value(v):
    """DataFrame 값 → SQLite에 넣을 값. 날짜는 'YYYY-MM-DD', 빈 값은 NULL."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime('%Y-%m-%d')
    if isinstance(v, np.generic):
        return v.item()
    return v


class OrderbookStore:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)

    def _connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.path)
        conn.execute('PRAGMA journal_mode=DELETE')
        row_cols = ', '.join(f"{_quote(c)} REAL" if c in NUM_COLUMNS else _quote(c) for c in RAW_COLUMNS)
        conn.executescript(f"""
            CREATE TABLE IF NOT EXISTS periods (
                year INTEGER NOT NULL, pn INTEGER NOT NULL, status TEXT NOT NULL, basis TEXT NOT NULL,
                calculated_at TEXT NOT NULL, calculated_by TEXT, ir_file TEXT, it_file TEXT, snapshot TEXT,
                row_count INTEGER NOT NULL, billed_rows INTEGER NOT NULL, PRIMARY KEY (year, pn));
            CREATE TABLE IF NOT EXISTS rows (year INTEGER NOT NULL, pn INTEGER NOT NULL, "_key" TEXT NOT NULL,
                {row_cols});
            CREATE INDEX IF NOT EXISTS rows_period ON rows (year, pn);
            CREATE TABLE IF NOT EXISTS billed (
                year INTEGER NOT NULL, pn INTEGER NOT NULL, bgroup TEXT NOT NULL, Project TEXT, Key TEXT, cur TEXT,
                qty REAL NOT NULL, amt REAL NOT NULL, PRIMARY KEY (year, pn, bgroup));
        """)
        return conn

    # ------------------------------------------------------------------ 읽기
    def _read(self, sql: str, params: list, columns: Sequence[str]) -> pd.DataFrame:
        """읽기는 파일을 만들지 않는다 — --dry-run이 빈 저장소를 남기지 않게"""
        if not self.path.exists():
            return pd.DataFrame(columns=list(columns))
        conn = self._connect()
        try:
            return pd.read_sql(sql, conn, params=params or None)
        finally:
            conn.close()

    def periods(self, year: int | None = None) -> pd.DataFrame:
        sql = 'SELECT * FROM periods' + (' WHERE year = ?' if year is not None else '') + ' ORDER BY year, pn'
        return self._read(sql, [year] if year is not None else [], _PERIOD_COLUMNS)

    def rows(self, year: int | None = None, pn: int | None = None) -> pd.DataFrame:
        """저장된 Raw 행 (year·pn·_key 열 포함). 인자 없이 부르면 전부."""
        where, params = [], []
        if year is not None:
            where.append('year = ?')
            params.append(year)
        if pn is not None:
            where.append('pn = ?')
            params.append(pn)
        sql = 'SELECT * FROM rows' + (' WHERE ' + ' AND '.join(where) if where else '') + ' ORDER BY year, pn, rowid'
        df = self._read(sql, params, _ROW_COLUMNS)
        for c in DATE_COLUMNS:
            df[c] = pd.to_datetime(df[c], errors='coerce')
        for c in NUM_COLUMNS:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        return df

    def billed(self, year: int, pn: int) -> pd.DataFrame:
        """그 달 말까지 인식한 청구 누계 (index = bgroup). 저장 때 센 행 수와 다르면 멈춘다 — 누계가 비면
        다음 달 Output이 그동안의 청구 전부가 되기 때문이다."""
        p = self.periods()
        hit = p[(p['year'] == year) & (p['pn'] == pn)]
        if hit.empty:
            raise PeriodError(f"{_label((year, pn))}이 저장돼 있지 않습니다")
        df = self._read('SELECT bgroup, Project, Key, cur, qty, amt FROM billed WHERE year = ? AND pn = ?',
                        [year, pn], ('bgroup',) + BILLED_COLUMNS)
        if len(df) != int(hit['billed_rows'].iloc[0]):
            raise PeriodError(f"{_label((year, pn))}의 청구 누계가 남아 있지 않습니다 "
                              f"(저장 {int(hit['billed_rows'].iloc[0]):,}행 / 지금 {len(df):,}행) — "
                              "마지막 달과 그 앞 달 것만 남긴다")
        return df.set_index('bgroup')[list(BILLED_COLUMNS)]

    # ------------------------------------------------------------------ 규칙
    def plan(self, year: int, month: int) -> Plan:
        """(year, month)를 계산할 수 있는지와 어디서 시작하는지. AX처럼 마지막 달만 다시 계산할 수 있다."""
        p = self.periods()
        if p.empty:
            return Plan('baseline', None)
        stored = {(int(y), int(m)): s for y, m, s in p[['year', 'pn', 'status']].itertuples(index=False)}
        latest = max(stored, key=lambda t: _ord(*t))
        target, prev, nxt = (year, month), previous_period(year, month), next_period(*latest)
        if target == latest:
            if stored[latest] == CLOSED:
                raise PeriodError(f"{_label(target)}은 닫힌 달이라 다시 계산할 수 없습니다")
            return Plan('recompute', prev) if prev in stored else Plan('baseline', None)
        if target == nxt:
            return Plan('next', latest)
        if _ord(*target) < _ord(*latest):
            raise PeriodError(f"{_label(target)}은 이미 닫혔습니다 — 다시 계산할 수 있는 건 마지막 달 {_label(latest)}, "
                              f"새로 계산할 달은 {_label(nxt)}")
        raise PeriodError(f"달을 건너뛸 수 없습니다 — 저장된 마지막 달이 {_label(latest)}이라 다음은 {_label(nxt)}")

    # ------------------------------------------------------------------ 쓰기
    def save(self, year: int, months: Sequence[int], rows: pd.DataFrame, billed: pd.DataFrame, basis: str,
             ir_file: str = '', it_file: str = '', snapshot: str = '') -> None:
        """그 달들을 한 트랜잭션으로 바꿔 넣는다 — 마지막 달만 Open, 그 앞은 전부 Closed.

        rows = Raw 행(RAW_COLUMNS + _key + pn), billed = 달별 청구 누계(bgroup + BILLED_COLUMNS + pn).
        청구 누계는 마지막 달과 그 앞 달 것만 남긴다(마지막 달을 다시 계산할 때 앞 달 누계가 필요하다)."""
        months = sorted(int(m) for m in months)
        p = self.periods(year)
        closed = sorted(set(months) & set(p.loc[p['status'].eq(CLOSED), 'pn'].astype(int)))
        if closed:
            raise PeriodError(f"닫힌 달은 바꿔 넣을 수 없습니다: {', '.join(f'P{m:02d}' for m in closed)}")
        last = months[-1]
        keep = {_ord(year, last), _ord(*previous_period(year, last))}
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        try:
            user = getpass.getuser()
        except (KeyError, OSError):
            user = ''
        row_sql = (f"INSERT INTO rows ({', '.join(_quote(c) for c in _ROW_COLUMNS)}) "
                   f"VALUES ({', '.join('?' * len(_ROW_COLUMNS))})")
        row_values = [tuple(_sql_value(v) for v in r)
                      for r in rows.assign(year=year)[list(_ROW_COLUMNS)].itertuples(index=False, name=None)]
        b = billed[billed['pn'].isin(months)]
        billed_count = b['pn'].astype(int).value_counts()
        b = b[b['pn'].map(lambda m: _ord(year, int(m)) in keep)]
        billed_values = [(year, int(r.pn), r.bgroup, r.Project, r.Key, r.cur, float(r.qty), float(r.amt))
                         for r in b.itertuples(index=False)]
        row_count = rows['pn'].astype(int).value_counts()
        period_values = [(year, m, OPEN if m == last else CLOSED, basis, now, user, ir_file, it_file, snapshot,
                          int(row_count.get(m, 0)), int(billed_count.get(m, 0))) for m in months]
        conn = self._connect()
        try:
            with conn:
                for m in months:
                    for table in ('rows', 'billed', 'periods'):
                        conn.execute(f'DELETE FROM {table} WHERE year = ? AND pn = ?', (year, m))
                conn.executemany(row_sql, row_values)
                conn.executemany('INSERT INTO billed VALUES (?, ?, ?, ?, ?, ?, ?, ?)', billed_values)
                conn.executemany(f"INSERT INTO periods ({', '.join(_PERIOD_COLUMNS)}) "
                                 f"VALUES ({', '.join('?' * len(_PERIOD_COLUMNS))})", period_values)
                conn.execute('UPDATE periods SET status = ? WHERE year * 100 + pn < ?', (CLOSED, _ord(year, last)))
                conn.execute(f"DELETE FROM billed WHERE year * 100 + pn NOT IN ({', '.join('?' * len(keep))})",
                             tuple(keep))
        finally:
            conn.close()
