"""
Final Invoice의 인보이스 단위 — INV 번호 (COM 없는 순수 판정)
==============================================================

**INV 번호가 인보이스다.** `DN_해외`의 INV 컬럼(`config.COLUMN_ALIASES['invoice_no']`)에
같은 번호를 넣은 행 전부가 한 장이고, DN·발주번호는 그 인보이스를 찾아가는 입구일 뿐이다 —
어느 입구로 들어가도 같은 번호는 같은 문서가 나와야 한다.

왜 DN 번호를 그대로 쓰지 않나 (2026-11-11 사용자 요청): DN 하나에 PO가 여럿이면 FI를
발주번호별로 나눠 내는데 그 전부에 같은 DN 번호가 찍힌다 — 업체 AP는 서로 다른 청구 여러 장을
같은 번호로 받는다. 실측 DN 215건 중 39건(`DNO-2026-0020`은 9장).

번호는 **사람이 시트에 값으로** 넣는다. 수식으로 매기면 행을 더하거나 정렬할 때 이미 발행한
번호가 바뀌고, 코드가 매기려면 마스터 워크북에 써야 한다(그 경로는 `create_dn.py` 하나뿐이다).
그래서 여기는 읽고 검증하고, 넣을 번호를 권하기만 한다 (`suggest_invoice_numbers` —
INV 번호는 DN 번호를 따르고 복수 PO 선적은 Customer PO별로 `-1`, `-2`).

Excel을 만지는 쪽(`fi_generator`)과 가른 이유는 `hs_code`·`dn_recorder`와 같다 — 판정을
COM 없이 테스트하고, 조회 서비스가 xlwings를 끌고 오지 않게.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime

import pandas as pd

from po_generator.config import COLUMN_ALIASES
from po_generator.utils import resolve_column, to_text

# 시트에 새로 만들 때 권하는 컬럼 이름 (안내 문구용 — 찾을 때는 별칭 전부를 본다)
INVOICE_COLUMN_NAME: str = COLUMN_ALIASES['invoice_no'][0]

# 권장 번호를 만들 때만 쓴다 — 이 모양이 아닌 DN은 권하지 않을 뿐, 생성에는 지장 없다
_DN_PATTERN = re.compile(r'^DNO-(\d{4})-(\d+)$', re.IGNORECASE)


def unique_texts(items: pd.DataFrame, key: str) -> list[str]:
    """그 컬럼의 값들 — 공란 제외, 처음 나온 순서, 중복 없이 (컬럼이 없으면 [])"""
    col = resolve_column(items.columns, key)
    if col is None:
        return []
    seen: list[str] = []
    for value in items[col]:
        text = to_text(value).strip()
        if text and text not in seen:
            seen.append(text)
    return seen


def invoice_numbers(items: pd.DataFrame) -> pd.Series:
    """라인별 INV 번호 — 앞뒤 공백을 떼고, 공란이거나 컬럼이 없으면 ''"""
    col = resolve_column(items.columns, 'invoice_no')
    if col is None:
        return pd.Series('', index=items.index, dtype=object)
    return items[col].map(lambda value: to_text(value).strip()).astype(object)


def invoices_of(items: pd.DataFrame, label: str) -> tuple[list[str], str | None]:
    """고른 라인(DN 한 건·발주번호 한 건)이 실린 인보이스 — (INV 번호 목록, 오류)

    - 전부 공란(또는 컬럼 없음) → `([], None)`: INV 번호를 쓰기 전 방식 그대로 만든다.
      이미 DN 번호로 나간 지난 건을 다시 뽑아도 번호가 바뀌지 않는다
    - 일부만 채움 → `([], 오류)`: 빈 라인은 어느 인보이스에도 안 실려 청구에서 빠진다.
      DN 번호로 따로 내면 한 선적에 번호 체계가 둘이 된다 — 사람이 채우고 다시 돌린다
    - 전부 채움 → `(정렬한 번호, None)`
    """
    numbers = invoice_numbers(items)
    filled = numbers != ''
    if not filled.any():
        return [], None
    if not filled.all():
        blank = items[~filled]
        orders = unique_texts(blank, 'rck_po')
        where = f" ({', '.join(orders)})" if orders else ''
        return [], (
            f"{label}: INV 번호가 일부 라인에만 있습니다 — 빈 라인 {len(blank)}건{where}은 "
            f"어느 인보이스에도 실리지 않습니다. 번호를 채운 뒤 다시 생성하세요"
        )
    return sorted(set(numbers)), None


def check_invoice_lines(items: pd.DataFrame) -> list[str]:
    """한 인보이스로 낼 수 없는 사유 — 비어 있으면 통과

    번호를 손으로 넣다 보니 **같은 번호가 남의 행에 들어갈 수 있다.** 그대로 만들면 한
    거래처의 청구서에 다른 거래처 품목·단가가 찍힌다(거래명세표의 '한 DN에 두 거래처'와
    같은 사고). 통화가 섞이면 Total이 서로 다른 통화를 더한 숫자가 된다.
    """
    errors: list[str] = []

    # 고객코드가 있으면 그것으로 — 같은 거래처의 이름 표기 차이를 둘로 세지 않는다
    customers = unique_texts(items, 'customer_code') or unique_texts(items, 'customer_name')
    if len(customers) > 1:
        names = unique_texts(items, 'customer_name') or customers
        errors.append(
            f"한 INV 번호에 거래처가 {len(customers)}곳입니다 ({', '.join(names)}) — "
            f"같은 번호를 다른 거래처 행에 넣지 않았는지 확인하세요"
        )

    currencies = unique_texts(items, 'currency')
    if len(currencies) > 1:
        errors.append(
            f"한 INV 번호에 통화가 {len(currencies)}가지입니다 ({', '.join(currencies)}) — "
            f"합계를 낼 수 없습니다"
        )
    return errors


@dataclass(frozen=True)
class FiReferences:
    """FI 헤더에 찍는 번호 — 서비스가 정하고 생성기는 쓰기만 한다 (Invoice Date와 같은 규약)"""

    invoice_no: str         # Invoice No — INV 번호 (없으면 예전처럼 DN 번호)
    customer_po: str        # Customer PO — 실린 라인의 것 전부 (대개 하나)
    customer_po_date: str   # Customer PO date — Customer PO와 같은 순서로 하나씩
    dn_ids: str             # Delivery Note No. — 실린 라인의 DN 전부 (CI·PL은 이 번호로 나간다)


def _date_text(value: object) -> str:
    """날짜 칸의 값 → 'YYYY-MM-DD' (공란은 '', 날짜가 아닌 값은 그대로)"""
    if value is None or isinstance(value, str) and not value.strip():
        return ''
    if isinstance(value, (datetime, date)):
        return '' if pd.isna(value) else value.strftime('%Y-%m-%d')
    return '' if pd.isna(value) else str(value)


def _po_dates(items: pd.DataFrame, customer_pos: list[str]) -> list[str]:
    """Customer PO마다 그 발주의 PO 일자 — `customer_pos`와 같은 순서 (없으면 '')

    PO가 하나도 없는 시트면 첫 날짜 하나. 한 발주를 SO 여러 개로 받았으면 먼저 나온 날짜.
    """
    date_col = resolve_column(items.columns, 'po_receipt_date')
    if date_col is None:
        return []
    po_col = resolve_column(items.columns, 'customer_po')
    pos = [to_text(v).strip() for v in items[po_col]] if po_col else [''] * len(items)

    first: dict[str, str] = {}
    for po, value in zip(pos, items[date_col]):
        text = _date_text(value)
        if text:
            first.setdefault(po, text)

    if not customer_pos:
        return [next(iter(first.values()), '')]
    return [first.get(po, '') for po in customer_pos]


def build_fi_references(
    items: pd.DataFrame,
    invoice_no: str | None = None,
) -> tuple[FiReferences, list[str]]:
    """인보이스에 실리는 라인에서 헤더 번호를 정한다 — (번호, 경고)

    Args:
        items: 인보이스에 실리는 라인 전체
        invoice_no: INV 번호. None이면 INV 번호를 쓰기 전 방식 — 첫 라인의 DN 번호
            (발주번호 통합이 여러 DN에 걸치면 대표 DN만 찍히던 그대로다)

    DN 번호를 따로 싣는 이유: CI·PL·COO는 선적 단위 서류라 계속 DN 번호로 나간다.
    Invoice No가 INV 번호로 바뀌면 업체가 청구서와 선적서류를 이을 단서가 이 칸뿐이다.
    """
    dn_ids = unique_texts(items, 'dn_id')
    customer_pos = unique_texts(items, 'customer_po')
    warnings: list[str] = []

    if invoice_no is None:
        number = dn_ids[0] if dn_ids else ''
    else:
        number = invoice_no
        if len(dn_ids) > 1:
            # 의도한 통합일 수도, 같은 번호를 두 선적에 잘못 넣은 것일 수도 있다
            warnings.append(
                f"DN {len(dn_ids)}건에 걸친 인보이스입니다 ({', '.join(dn_ids)}) — "
                f"같은 INV 번호를 의도한 것인지 확인하세요"
            )

    if len(customer_pos) > 1:
        # 품목 표에는 PO 열이 없어 어느 줄이 어느 PO인지 문서만으로는 알 수 없다
        warnings.append(
            f"한 인보이스에 Customer PO가 {len(customer_pos)}개 실렸습니다 "
            f"({', '.join(customer_pos)}) — PO별로 따로 청구하려면 PO마다 다른 INV 번호를 넣으세요"
        )

    # PO가 여럿이면 날짜도 같은 순서로 하나씩 — 빈 것은 자리를 지켜 짝이 밀리지 않게
    po_dates = _po_dates(items, customer_pos)
    if len(po_dates) > 1:
        po_dates = [d or '-' for d in po_dates]

    return FiReferences(
        invoice_no=number,
        customer_po=', '.join(customer_pos),
        customer_po_date=', '.join(po_dates),
        dn_ids=', '.join(dn_ids),
    ), warnings


def suggest_invoice_numbers(items: pd.DataFrame) -> list[tuple[str, str]]:
    """DN 한 건의 라인에 넣을 INV 번호 — `[(번호, Customer PO), …]` (권할 수 없으면 [])

    번호 규칙 (2026-11-11 사용자 결정): **INV 번호는 DN 번호를 따른다** — `DNO-2026-0215` →
    `INV-2026-0215`. Customer PO가 여럿인 선적은 PO마다 한 장이라 시트에 나온 순서대로
    `-1`, `-2`…를 붙인다 (`DNO-2026-0191` → `INV-2026-0191-1`, `-2`).

    번호는 사람이 넣으므로 안내만 한다. '지금까지 쓴 가장 큰 번호 + 1'로 권하면 이 규칙과
    어긋난다 — 0191만 번호가 있을 때 `DNO-2026-0215`에 0192를 권했다 (같은 날 실측).
    DN이 하나가 아니거나(발주번호 통합) 번호 모양이 다르면 권하지 않는다.
    """
    dn_ids = unique_texts(items, 'dn_id')
    if len(dn_ids) != 1:
        return []
    match = _DN_PATTERN.match(dn_ids[0])
    if not match:
        return []
    base = f"INV-{match.group(1)}-{match.group(2)}"
    customer_pos = unique_texts(items, 'customer_po')
    if len(customer_pos) <= 1:
        return [(base, customer_pos[0] if customer_pos else '')]
    return [(f"{base}-{n}", po) for n, po in enumerate(customer_pos, start=1)]
