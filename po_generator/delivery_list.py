"""
공장 출고리스트(`2026리스트_RCK_{period}.xlsx`) Delivery 시트 읽기
================================================================

출고리스트는 공장이 만들어 보내는 파일이라 **헤더가 예고 없이 바뀐다**. 2026-11(P11)에
영문·혼용 헤더가 한글로 바뀌었고 `납품 예정일`이 빠지고 `분할 납품`이 생겼다:

    P03~P10          P11~
    ---------------  ------------------
    DATE             구분
    JOB NO           수주번호
    BUYER            고객명              ← 늘 'RCK'. 거래처가 아니다
    Customer         발주처              ← 거래처는 이쪽
    AMOUNT           발주 총 금액
    RCK ODER         발주번호
    납품 예정일      (없음)
    납품완료         납품 일자
    계산서금액       납품된 항목 금액
    (없음)           분할 납품
    SO_ID · SECTOR · Type · 비고 는 그대로

이 시트를 읽는 곳이 둘이다 — `dn_recorder`(DN 출고기록 자동 입력)와 `reconcile_po`(매입대사).
그래서 읽는 순간 **예전 이름으로 맞추고**(`normalize_columns`) 그 뒤 코드·메시지·결과 파일은
양식과 무관하게 한 이름만 쓴다. 지난 달 파일을 다시 돌려도 같고, 회계팀에 가는
`AX_PO_매핑`의 헤더도 달마다 같다. 헤더가 또 바뀌면 `HEADER_ALIASES`에 한 줄.

COM·DB 없음 — pandas만 쓴다.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

DELIVERY_SHEET = 'Delivery'

# 기준 이름 = P03~P10 헤더 (코드·메시지·결과 파일이 쓰는 이름)
COL_SO_ID = 'SO_ID'
COL_SHIP_DATE = '납품완료'
COL_INVOICE_AMOUNT = '계산서금액'
COL_CUSTOMER = 'Customer'
COL_RCK_ORDER = 'RCK ODER'
COL_JOB_NO = 'JOB NO'       # 공장 발주번호 = PO_국내 `NOAH O.C No.` (+ 회차 접미사 `-1R`)

# 기준 이름 → 다른 양식의 헤더.
# `고객명`을 Customer로 풀면 안 된다 — 공장 입장의 고객(= RCK)이라 전 행이 'RCK'다.
# 거래처는 `발주처`다 (P11 실측 105행: 고객명 = 'RCK' 105행).
HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    'DATE': ('구분',),
    COL_JOB_NO: ('수주번호',),
    'BUYER': ('고객명',),
    COL_CUSTOMER: ('발주처',),
    'AMOUNT': ('발주 총 금액',),
    COL_RCK_ORDER: ('발주번호',),
    COL_SHIP_DATE: ('납품 일자',),
    COL_INVOICE_AMOUNT: ('납품된 항목 금액',),
}

_SPACE_RE = re.compile(r'\s+')


def _key(header: object) -> str:
    """헤더 대조 키 — 공백 차이는 같은 헤더로 본다 (`납품 일자` = `납품일자`)"""
    return _SPACE_RE.sub('', str(header))


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """출고리스트 헤더를 기준 이름으로 맞춘다 (모르는 열은 그대로 둔다)

    기준 이름의 열이 이미 있으면 그 별칭은 건드리지 않는다 — 같은 이름의 열이 둘이 되면
    `df['계산서금액']`이 Series가 아니라 DataFrame이 되어 엉뚱한 곳에서 터진다.
    """
    by_key = {_key(c): c for c in df.columns}
    rename: dict[object, str] = {}
    for canonical, aliases in HEADER_ALIASES.items():
        actual = by_key.get(_key(canonical))
        if actual is not None:
            if actual != canonical:
                rename[actual] = canonical      # 공백만 다른 기준 이름
            continue
        for alias in aliases:
            actual = by_key.get(_key(alias))
            if actual is not None:
                rename[actual] = canonical
                break
    if not rename:
        return df
    logger.debug("출고리스트 헤더 맞춤: %s",
                 ", ".join(f"{old} → {new}" for old, new in rename.items()))
    return df.rename(columns=rename)


def read_delivery_sheet(delivery_file: Path,
                        required: tuple[str, ...] = ()) -> pd.DataFrame:
    """Delivery 시트를 읽어 헤더를 기준 이름으로 맞춘다

    Args:
        delivery_file: 출고리스트 파일
        required: 꼭 있어야 하는 열(기준 이름). 없으면 ValueError — 열이 빠진 채 넘어가면
            금액이 0으로 읽혀 전 건이 '금액 불일치'로 뜨는 식으로 엉뚱하게 드러난다

    Raises:
        ValueError: 필요한 열이 없을 때 (찾은 이름과 시트의 실제 헤더를 같이 알려 준다)
    """
    raw = pd.read_excel(delivery_file, sheet_name=DELIVERY_SHEET)
    df = normalize_columns(raw)
    missing = [c for c in required if c not in df.columns]
    if missing:
        wanted = ", ".join(
            " / ".join(f"'{n}'" for n in (c, *HEADER_ALIASES.get(c, ())))
            for c in missing)
        raise ValueError(
            f"출고리스트에 필요한 열이 없습니다: {wanted} ({delivery_file.name}) "
            f"— 시트의 헤더: {', '.join(str(c) for c in raw.columns)}. "
            "헤더가 바뀌었으면 po_generator/delivery_list.py의 HEADER_ALIASES에 추가하세요.")
    return df
