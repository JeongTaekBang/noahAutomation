"""
AX 오더북 결과 HTML — Edge에서 더블클릭으로 여는 인터랙티브 보고서
================================================================

엑셀 결과와 같은 숫자를 HTML 파일 하나에 담는다. 회사 PC는 서버를 못 띄우고 CDN도 믿을 수 없어서, 데이터(JSON)와
차트 엔진(이미 설치된 Plotly의 JS)을 파일 안에 통째로 넣는다 — 인터넷·설치 없이 열린다. 화면과 필터·합계 로직은
템플릿(`templates/ax_orderbook_report.html`, 빌드 도구 없는 순수 JS)이 갖고, 여기는 데이터를 고르고 줄여서 끼워 넣기만 한다.

    report_data(...)  결과 → JSON으로 바꿀 수 있는 dict (순수 함수 — 테스트는 여기서)
    write_report(...) 템플릿의 자리표시자 두 개(데이터, Plotly 엔진)를 채워 파일로

데이터 모양 — 표는 열 단위 {'n': 행 수, 'cols': {열: [값...]}, 'str': [사전 열...]}. 반복되는 문자열은 공유 사전
`strings`의 번호로 넣는다(Raw 8,911행의 프로젝트·고객·코드가 대부분 반복). 금액은 원(소수 2자리까지).

    keys    키마다 한 행 — 프로젝트·고객·Order Book Code·Market·Cost center 등. 필터는 이 속성으로 건다
            (우리 Raw의 마지막 행, 우리에게 없는 키는 회계 행에서)
    ours    달·키별 우리 계산 (pn, k=키 번호, S·I·O·V·E·Eq)
    acct    달·키별 회계 오더북 (같은 해, 계산한 달까지) — 필터한 회계 숫자와 프로젝트 일치율을 JS가 바로 낸다
    diff    회계와의 차이와 원인 (엑셀 '차이_키별')
    lines   마지막 달 라인내역 / pending 출고 후 미청구 / review 확인목록 / fx 환율 / store 저장 기록
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd

from po_generator.ax_orderbook import DIFF_CAUSE_NOTES, DIFF_CAUSES, OrderbookResult
from po_generator.config import AX_ORDERBOOK_REPORT_TEMPLATE

PLACEHOLDER_DATA = '__AX_DATA__'
PLACEHOLDER_PLOTLY = '__PLOTLY_JS__'
# 키 속성 — JS 이름 → Raw 열
KEY_ATTRIBUTES: dict[str, str] = {
    'key': 'Key', 'project': 'Project', 'customer': 'Customer Name', 'account': 'Account number',
    'obc': 'Order Book Code', 'market': 'Market', 'cc': 'Cost center', 'dept': 'Prod/Dept (Revenue)',
    'industry': 'Industry', 'name': 'Name',
}


class _Strings:
    """반복 문자열 사전 — 같은 문자열은 같은 번호"""

    def __init__(self) -> None:
        self.values: list[str] = []
        self._index: dict[str, int] = {}

    def idx(self, value) -> int | None:
        if _blank(value):
            return None
        text = str(value)
        if text not in self._index:
            self._index[text] = len(self.values)
            self.values.append(text)
        return self._index[text]


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and v == '') or (not isinstance(v, str) and pd.isna(v))


def _number(v, digits: int):
    """소수 digits자리로, 정수면 int로 (JSON이 '184168.0'이 아니라 '184168'이 되게)"""
    if _blank(v):
        return None
    x = round(float(v), digits)
    return int(x) if x.is_integer() else x


def _day(v) -> str | None:
    if _blank(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime('%Y-%m-%d')
    return str(v)


def _plain(v):
    """표 한 칸 → JSON 값 (사전 없이)"""
    if _blank(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return _day(v)
    if isinstance(v, (bool, np.bool_)):
        return bool(v)
    if isinstance(v, (int, float, np.integer, np.floating)):
        return _number(v, 4)
    return str(v)


def _pack(df: pd.DataFrame, spec: dict[str, tuple[str, str]], strings: _Strings) -> dict:
    """DataFrame → 열 단위 표. spec: 출력 열 → (원본 열, 종류). 종류 's' 문자열(사전), 't' 문자열(그대로),
    'i' 정수, 'm' 금액(소수 2자리), 'q' 수량(4자리), 'd' 날짜."""
    convert = {
        's': strings.idx, 't': lambda v: None if _blank(v) else str(v),
        'i': lambda v: None if _blank(v) else int(v), 'm': lambda v: _number(v, 2), 'q': lambda v: _number(v, 4),
        'd': _day,
    }
    cols = {name: [convert[kind](v) for v in (df[src] if src in df.columns else [None] * len(df))]
            for name, (src, kind) in spec.items()}
    return {'n': len(df), 'cols': cols, 'str': [name for name, (_, kind) in spec.items() if kind == 's']}


def _records(df: pd.DataFrame) -> list[dict]:
    return [{str(c): _plain(v) for c, v in row.items()} for row in df.to_dict('records')] if len(df) else []


def _key_attributes(raw_rows: pd.DataFrame, official: pd.DataFrame | None) -> pd.DataFrame:
    """키마다 한 행 — 우리 Raw의 마지막 행, 우리에게 없는(회계에만 있는) 키는 회계의 마지막 행"""
    cols = [c for c in KEY_ATTRIBUTES.values() if c != 'Key']
    ours = (raw_rows.assign(Key=raw_rows['_key']).sort_values('pn', kind='stable')
            .drop_duplicates('Key', keep='last')[['Key', *cols]])
    if official is None or official.empty:
        return ours.reset_index(drop=True)
    off = official.sort_values('pn', kind='stable').drop_duplicates('Key', keep='last')
    off = off.loc[~off['Key'].isin(ours['Key']), ['Key', *cols]]
    return pd.concat([ours, off], ignore_index=True) if len(off) else ours.reset_index(drop=True)


def report_data(res: OrderbookResult, diff: pd.DataFrame, official: pd.DataFrame | None, store_periods: pd.DataFrame,
                files: dict[str, str], note: str, preview: bool = False) -> dict:
    """결과 → HTML에 넣을 dict. res.keys·raw_rows = 보여 줄 달 전부(저장된 달 + 이번 달), official = 같은 해 회계 행
    (계산한 달까지 — 없으면 빈 표), diff = classify_differences 결과."""
    strings = _Strings()
    official = official if official is not None else pd.DataFrame(columns=['pn', 'Key'])
    keys = _key_attributes(res.raw_rows, official)
    kidx = {k: i for i, k in enumerate(keys['Key'])}
    to_k = lambda s: s.map(kidx).astype('Int64')      # noqa: E731 — 키 → 번호 (없으면 빈 값)

    ours = res.keys.assign(k=lambda d: to_k(d['Key']))
    acct = pd.DataFrame(columns=['pn', 'k', 'S', 'I', 'O', 'V', 'E'])
    if len(official):
        acct = (official.groupby(['pn', 'Key'], as_index=False)[
                    ['Value Start', 'Value Input', 'Value Output', 'Value Variance', 'Value End']].sum()
                .assign(k=lambda d: to_k(d['Key'])))
    official_months = sorted({int(p) for p in official['pn'].dropna()}) if len(official) else []

    d = diff.copy() if len(diff) else pd.DataFrame(columns=['Period', 'Key', '원인'])
    d['pn'] = d['Period'].astype(str).str[1:].astype(int) if len(d) else []
    d['c'] = d['원인'].map({c: i for i, c in enumerate(DIFF_CAUSES)})
    d['k'] = to_k(d['Key']) if len(d) else []

    lines = res.lines.copy()
    pending = res.pending.assign(k=lambda p: to_k(p['Key'])) if len(res.pending) else res.pending
    return {
        'meta': {
            'title': f"AX 오더북 {res.year} {res.period}", 'year': res.year, 'upto': res.upto, 'period': res.period,
            'generated': f"{datetime.now():%Y-%m-%d %H:%M}", 'files': files, 'note': note,
            'filter_note': res.next_filter_note, 'preview': bool(preview),
        },
        'periods': [{'pn': m, 'label': f"P{m:02d}", 'official': m in official_months} for m in range(1, res.upto + 1)],
        'causes': [{'name': c, 'note': DIFF_CAUSE_NOTES.get(c, '')} for c in DIFF_CAUSES],
        'keys': _pack(keys, {name: (src, 's') for name, src in KEY_ATTRIBUTES.items()}, strings),
        'ours': _pack(ours, {'pn': ('pn', 'i'), 'k': ('k', 'i'), 'S': ('Sv', 'm'), 'I': ('Iv', 'm'), 'O': ('Ov', 'm'),
                             'V': ('Vv', 'm'), 'E': ('Ev', 'm'), 'Eq': ('Eq', 'q')}, strings),
        'acct': _pack(acct, {'pn': ('pn', 'i'), 'k': ('k', 'i'), 'S': ('Value Start', 'm'), 'I': ('Value Input', 'm'),
                             'O': ('Value Output', 'm'), 'V': ('Value Variance', 'm'), 'E': ('Value End', 'm')}, strings),
        'diff': _pack(d, {'pn': ('pn', 'i'), 'k': ('k', 'i'), 'c': ('c', 'i'), 'dS': ('d_Start', 'm'),
                          'dI': ('d_Input', 'm'), 'dO': ('d_Output', 'm'), 'dE': ('d_End', 'm'),
                          'dV': ('d_Variance', 'm'), 'oE': ('Ev', 'm'), 'aE': ('Value End', 'm')}, strings),
        'lines': _pack(lines, {'project': ('Project', 's'), 'key': ('Key', 's'), 'line': ('Line number', 's'),
                               'item': ('Item number', 's'), 'itemName': ('Item name', 's'), 'cur': ('cur', 's'),
                               'status': ('Line status', 's'), 'created': ('created_kst', 'd'),
                               'ship': ('Actual ship date', 'd'), 'qty': ('Quantity', 'q'), 'net': ('Net amount', 'm'),
                               'inV': ('in_v', 'm'), 'outV': ('out_v', 'm'), 'endV': ('end_v', 'm'),
                               'inQ': ('in_q', 'q'), 'outQ': ('out_q', 'q'), 'endQ': ('end_q', 'q')}, strings),
        'pending': _pack(pending, {'project': ('Project', 's'), 'key': ('Key', 's'), 'k': ('k', 'i'),
                                   'line': ('Line number', 's'), 'itemName': ('Item name', 's'),
                                   'first': ('첫 출고일', 'd'), 'last': ('마지막 출고일', 'd'),
                                   'months': ('경과 개월', 'i'), 'qty': ('수량', 'q'), 'cur': ('통화', 's'),
                                   'net': ('금액(통화)', 'm'), 'krw': ('금액(원)', 'm')}, strings),
        'review': _pack(res.review, {'kind': ('구분', 's'), 'key': ('Key', 's'), 'project': ('Project', 's'),
                                     'customer': ('고객', 's'), 'amount': ('금액', 'm'), 'detail': ('내용', 't')}, strings),
        'store': _records(store_periods),
        'fx': _records(res.fx_used),
        'strings': strings.values,
    }


def _payload(data: dict) -> str:
    """dict → <script> 안에 그대로 넣을 JSON. '<'를 전부 \\u003c로 — 문자열 속 '</script>'나 '<!--'가 태그를 닫지 못하게"""
    return json.dumps(data, ensure_ascii=False, allow_nan=False, separators=(',', ':')).replace('<', '\\u003c')


def write_report(data: dict, path: Path, template: Path = AX_ORDERBOOK_REPORT_TEMPLATE) -> Path:
    """템플릿에 데이터와 Plotly 엔진을 끼워 HTML 파일 하나로. 그 이름이 잠겨 있으면 시각을 붙인 이름으로(엑셀과 같게)."""
    from plotly.offline import get_plotlyjs      # 이 함수만 쓰므로 여기서 — 계산 모듈은 Plotly 없이도 돈다

    pieces = {PLACEHOLDER_DATA: _payload(data), PLACEHOLDER_PLOTLY: get_plotlyjs()}
    parts = re.split('(' + '|'.join(map(re.escape, pieces)) + ')', template.read_text(encoding='utf-8'))
    found = sorted(p for p in parts if p in pieces)
    if found != sorted(pieces):
        raise ValueError(f"{template.name}: 자리표시자 {sorted(pieces)}가 한 번씩 있어야 합니다 (찾은 것: {found})")
    html = ''.join(pieces.get(p, p) for p in parts)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        path.write_text(html, encoding='utf-8')
        return path
    except PermissionError:
        other = path.with_name(f"{path.stem}_{datetime.now():%H%M%S}{path.suffix}")
        other.write_text(html, encoding='utf-8')
        return other
