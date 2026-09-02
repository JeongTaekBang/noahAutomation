"""
Sectoriel 라인별 HS CODE 판정
==============================

거래처 SECTORIEL이 **자기네 수입통관에 쓰는 HS CODE**를 CI/PL 라인마다 찍어달라고
요청해서(2026-09-02) 만든 모듈이다. 기존 CI/PL은 문서 단위 HS 하나(`I12`)만 갖는다.

**왜 별도 모듈인가.** 판정은 순수 계산이라 Excel COM 없이 전수 테스트할 수 있어야 한다
(`dn_recorder`와 `dn_writer`를 가른 것과 같은 이유). 생성기는 결과 열을 쓰기만 한다.

**판정 우선순위가 곧 안전장치다** — 통관서류라 틀린 코드의 비용이 크다:

1. **운임·수수료 줄** → 빈칸, 경고 없음. `FEDEX COST` 라인의 `OS name`이 실제로
   `Noah NA`/`SA`로 찍혀 있어서(2026 실측 2건) 이걸 먼저 걸러내지 않으면
   **운임에 액추에이터 HS가 붙는다**.
2. **Model number 정확 매칭** — 스페어는 `980748` 같은 안정 코드를 갖는다.
   품목명은 프랑스어·영어가 섞이고 표기가 흔들리지만(`CI PUIS.+LED SA05`,
   `CIRCUIT LED BOARD SA05/09`) 모델번호는 안 흔들린다.
3. **품목명 키워드** — 고객 표에는 있지만 아직 출고 이력이 없는 항목들(HEATER·Inner
   Frame·Governor·Battery Pack·Handle Wheel) 대비. 보드류도 모델번호가 새로 생겼을 때
   여기서 잡힌다.
4. **`OS name`이 액추에이터 계열** → `85013100`. 200줄 중 186줄이 여기서 끝난다.
   **SR을 포함한다** — SR05/SR10/SR30은 스프링리턴 '액추에이터 본체'고, 고객 표의
   'Governor for SR'은 SR**용 부품**이지 SR 본체가 아니다 (2026-09-02 사용자 확인).
   덤으로 `OS name` 오타(SA05X가 SR로, SR03이 SA로 찍힌 실측 2건)가 무해해진다.

   **품목명보다 뒤에 두는 이유**: `OS name`은 '이 라인에 실린 물건'이 아니라 '이 주문
   라인이 속한 제품군'이라 이미 한 번 거짓말했다(운임 3줄). 액추에이터에 딸려 나가는
   히터·핸들휠은 제품군이 `Noah NA`로 찍힐 수 있는데, 고객 표는 그것들을 별도 세번으로
   나눴다. 순서를 바꿔도 실데이터 200줄 결과는 **완전히 동일**하다(오분류 0건, 2026-09-02
   실측) — 공짜로 얻는 안전장치라 뒤로 뺐다. 대신 보드 규칙은 액추에이터 설명문을
   훔치지 않게 앵커링한다(`BOARD` 단독 금지).
5. **그 외 → 빈칸 + 미매칭 경고**. 코드를 추측하지 않는다. 2026년 실적에선 전기모터
   (`MOTEUR ELEC.SEUL`, 980737/980739)와 스템 커버(`CAPOT TIGE`, 980800)가 여기 걸린다 —
   고객이 코드를 알려주면 `SPARE_HS_BY_MODEL`에 한 줄씩 추가하면 된다.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

import pandas as pd

from po_generator.config import HS_LINE_CUSTOMERS
from po_generator.utils import get_value, to_text

logger = logging.getLogger(__name__)


# === 고객 제공 HS 코드표 (SECTORIEL, 2026-09-02) ===
# 왼쪽 이름은 고객이 쓰는 designation 그대로 — 나중에 대조할 때 헷갈리지 않게 둔다.
HS_ELECTRIC_ACTUATOR = '85013100'   # ELECTRIC ACTUATOR
HS_HEATER_SMALL = '85352900'        # HEATER for NA006 - NA009
HS_HEATER_LARGE = '84818099'        # HEATER for NA015 - NA250
HS_INNER_FRAME = '85030099'         # Inner Frame assembly for SA005
HS_BOARDS = '85352900'              # Boards for All Model
HS_GOVERNOR = '85352900'            # Governor for SR
HS_BATTERY_PACK = '85352900'        # Battery Pack for NA/SA RBP
HS_HANDLE_WHEEL = '84835080'        # Handle Wheel


# === 1. 물품이 아닌 줄 (HS 자체가 없다) ===
# 운임 청구 줄에도 OS name이 붙어 있어서 OS name 규칙보다 먼저 본다.
NON_GOODS_PATTERN = re.compile(
    r'\bFEDEX\b|\bFREIGHT\b|\bCOURIER\b|\bDHL\b|\bEMS\b|운임|택배',
)

# === 2. 모델번호 → HS (스페어의 안정 키) ===
SPARE_HS_BY_MODEL: dict[str, str] = {
    '980748': HS_BOARDS,          # CIRCUIT LED BOARD SA05/09 230VAC
    '980750': HS_BOARDS,          # CI PUIS.+LED SA05 (PC board + LED board)
    '980770': HS_BOARDS,          # CI NA-PCU 230V
    '980772': HS_BOARDS,          # CI CONVERTI 24VCA a 24VCC NA06
    '980773': HS_BOARDS,          # CI+CONDENSATEUR SA05SCP 230VCA
    '0812880001VCDE': HS_BOARDS,  # CIRCUIT IMPRIME NA-PCU 24VCC
}

# === 3. 액추에이터 계열 OS name ===
# 시트 표기 그대로 정규화해서 비교한다 ('Noah NA'가 실제 값).
ACTUATOR_OS_NAMES: frozenset[str] = frozenset({
    'NOAH NA', 'NA', 'SA', 'SR', 'MS', 'NL',
})

# === 4. 품목명 키워드 규칙 (위에서부터 먼저 맞는 것) ===
# 보드 규칙이 가장 넓어서 맨 끝에 둔다 — HEATER/GOVERNOR가 보드로 새면 코드가 갈린다.
_HEATER_PATTERN = re.compile(r'\bHEATER\b|RESISTANCE\s*CHAUFF|CHAUFFANT')
_HEATER_MODEL_PATTERN = re.compile(r'\b(?:NA|SA)\s*0*(\d{1,3})\b')

_ITEM_RULES: tuple[tuple[re.Pattern[str], str, str], ...] = (
    (re.compile(r'INNER\s*FRAME|CHASSIS\s*INTERIEUR'), HS_INNER_FRAME, 'Inner Frame assembly'),
    (re.compile(r'\bGOVERNOR\b|REGULATEUR|RÉGULATEUR'), HS_GOVERNOR, 'Governor for SR'),
    (re.compile(r'\bBATTERY\b|\bRBP\b|BATTERIE'), HS_BATTERY_PACK, 'Battery Pack'),
    (re.compile(r'HAND\s*WHEEL|HANDWHEEL|\bVOLANT\b'), HS_HANDLE_WHEEL, 'Handle Wheel'),
    # 보드는 **앵커링한다** — `\bBOARD\b` 단독으로 두면 액추에이터 설명문에 'board'가
    # 한 번만 섞여도 스페어 코드를 받는다 (이 규칙이 OS 폴백보다 앞에 있으므로
    # 186줄 위를 지나간다). 실측 200줄에서 앵커/비앵커 결과는 동일 — 안전한 쪽을 쓴다.
    (re.compile(
        r'^CI[\s+.]|CIRCUIT\s+IMPRIME|CIRCUIT\s+LED|\bPCB\b|\bCARTE\b'
        r'|(?:PCU|POWER|MAIN|CONTROL|TERMINAL|STATUS|LED|PC|RELAY)[\s\-]*BOARD'
    ), HS_BOARDS, 'Boards for All Model'),
)


@dataclass(frozen=True)
class HsResult:
    """한 라인의 HS 판정 결과

    Attributes:
        code: 셀에 쓸 HS 코드. **빈 문자열이면 아무것도 쓰지 않는다**
        rule: 어느 규칙이 정했는지 (로그·검증용)
        unmatched: True면 사람이 코드를 정해야 하는 줄 (경고 대상).
                   운임처럼 '원래 HS가 없는' 줄은 code가 비어도 False다
    """
    code: str = ''
    rule: str = ''
    unmatched: bool = False


def is_hs_line_customer(customer_name: object) -> bool:
    """라인별 HS CODE 판을 따로 받는 거래처인가 (고객명 부분일치)

    시트 표기가 흔들려도 걸리도록 부분일치로 본다 (`TS_PO_REMARK_CUSTOMERS`와 같은 규약).
    """
    name = _norm_text(customer_name)
    return bool(name) and any(keyword.upper() in name for keyword in HS_LINE_CUSTOMERS)


def _norm_text(value: object) -> str:
    """대문자 + 공백 정규화 (판정 비교용)"""
    text = to_text(value)
    return re.sub(r'\s+', ' ', text).strip().upper()


def _norm_model(value: object) -> str:
    """모델번호 정규화 — Excel이 숫자로 읽어 `980748.0`이 되는 것까지 흡수"""
    return _norm_text(value).replace(' ', '')


def _heater_code(item_upper: str) -> tuple[str, str]:
    """HEATER는 대상 모델 크기로 코드가 갈린다 (NA006-009 / NA015-250)"""
    match = _HEATER_MODEL_PATTERN.search(item_upper)
    if not match:
        return '', ''
    size = int(match.group(1))
    if 6 <= size <= 9:
        return HS_HEATER_SMALL, 'HEATER for NA006 - NA009'
    if 15 <= size <= 250:
        return HS_HEATER_LARGE, 'HEATER for NA015 - NA250'
    return '', ''


def resolve_hs_code(
    os_name: object = '',
    model_number: object = '',
    item_name: object = '',
) -> HsResult:
    """한 라인의 HS 코드를 판정한다 (모듈 docstring의 우선순위대로)

    Args:
        os_name: `SO_해외.OS name` (Noah NA / SA / SR / MS / Electric Spares)
        model_number: `SO_해외.Model number`
        item_name: 품목명 (DN의 `Item`)

    Returns:
        HsResult — `code`가 비어 있고 `unmatched`가 True면 사람이 정해야 하는 줄
    """
    item_upper = _norm_text(item_name)
    model_key = _norm_model(model_number)
    os_key = _norm_text(os_name)

    # 1. 운임·수수료 — 물품이 아니라 HS가 없다 (경고 대상 아님)
    if NON_GOODS_PATTERN.search(item_upper):
        return HsResult(code='', rule='운임·수수료(HS 없음)', unmatched=False)

    # 2. 모델번호 정확 매칭
    if model_key and model_key in SPARE_HS_BY_MODEL:
        return HsResult(
            code=SPARE_HS_BY_MODEL[model_key],
            rule=f'모델번호 {model_key}',
        )

    # 3. 품목명 키워드 (OS 폴백보다 먼저 — docstring의 이유)
    if _HEATER_PATTERN.search(item_upper):
        heater_code, heater_rule = _heater_code(item_upper)
        if heater_code:
            return HsResult(code=heater_code, rule=heater_rule)
        # 히터인 건 알겠는데 대상 모델 구간을 못 읽었다 — 둘 중 하나를 찍지 않는다
        return HsResult(code='', rule='히터(모델 구간 불명)', unmatched=True)

    for pattern, code, label in _ITEM_RULES:
        if pattern.search(item_upper):
            return HsResult(code=code, rule=label)

    # 4. 액추에이터 계열 OS name (제품군 폴백)
    if os_key in ACTUATOR_OS_NAMES:
        return HsResult(code=HS_ELECTRIC_ACTUATOR, rule=f'OS name {os_key}')

    # 5. 모르는 것은 비운다 — 추측하지 않는다
    return HsResult(code='', rule='미매칭', unmatched=True)


# SO_해외의 제품군 컬럼. DN에는 없어서 `document_service._enrich_from_so_lines`가
# (SO_ID, Line item)로 조인해 붙인다.
OS_NAME_COLUMN = 'OS name'
# 판정 결과가 실리는 열. 이 열이 있으면 생성기가 HS 판으로 만든다.
HS_COLUMN = '_hs_code'


def enrich_hs_codes(items_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """아이템 DataFrame에 `_hs_code` 열을 붙이고 미매칭 품목명을 돌려준다

    생성기는 이 열이 있으면 D열에 쓴다. **열이 행과 함께 다니므로** 생성기가
    품목을 정렬해도 코드가 행에서 떨어져 나가지 않는다 (별도 리스트로 넘기면 어긋난다).

    Args:
        items_df: DN 아이템 (Model number·OS name이 보강된 상태)

    Returns:
        (`_hs_code` 열이 추가된 사본, 사람이 코드를 정해야 하는 품목명 목록)
    """
    enriched = items_df.copy()
    codes: list[str] = []
    unmatched: list[str] = []

    for _, item in enriched.iterrows():
        result = resolve_hs_code(
            os_name=item.get(OS_NAME_COLUMN, ''),
            model_number=get_value(item, 'model', ''),
            item_name=get_value(item, 'item_name', ''),
        )
        codes.append(result.code)
        if result.unmatched:
            label = to_text(get_value(item, 'item_name', '')) or '(품목명 없음)'
            model = to_text(get_value(item, 'model', ''))
            unmatched.append(f"{label} [{model}]" if model else label)

    enriched[HS_COLUMN] = codes

    if unmatched:
        logger.warning(f"HS 코드 미지정 {len(unmatched)}줄: {unmatched}")
    return enriched, unmatched
