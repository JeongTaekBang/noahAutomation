"""
CI/PL 라인별 HS CODE 테스트 (Excel 없이)
========================================

두 경로를 다 지킨다 — 거래처 고유 코드를 받지 않은 문서의 **기본값 일괄**
(`TestDefaultCode`)과, 고객 코드표가 있는 거래처의 **판정 사다리**(나머지).
어느 쪽이든 운임 줄은 비운다.

품목명·모델번호는 **2026년 SECTORIEL 실출고 데이터에서 그대로 가져왔다**
(DN 200줄 / 스페어 12줄). 가공한 예시로 테스트하면 프랑스어 표기나 OS name 오염
같은 실제 함정을 못 잡는다.
"""

import pandas as pd
import pytest

from po_generator.hs_code import (
    HS_BOARDS,
    HS_ELECTRIC_ACTUATOR,
    HS_HEATER_LARGE,
    HS_HEATER_SMALL,
    HS_COLUMN,
    default_hs_result,
    enrich_hs_codes,
    resolve_hs_code,
)


class TestActuators:
    """OS name이 액추에이터 계열이면 본체 코드 — 200줄 중 166줄이 여기서 끝난다"""

    @pytest.mark.parametrize('os_name', ['Noah NA', 'SA', 'SR', 'MS', 'NL'])
    def test_액추에이터_계열은_전동액추에이터_코드(self, os_name):
        result = resolve_hs_code(os_name=os_name, item_name='NA15 150Nm 20s. 230V50HZ')
        assert result.code == HS_ELECTRIC_ACTUATOR

    def test_SR_본체는_액추에이터다(self):
        """'Governor for SR'은 SR용 부품이지 SR 본체가 아니다 (실데이터 22줄)"""
        result = resolve_hs_code(
            os_name='SR',
            model_number='023080',
            item_name='SR05 50Nm 230V50HZ RA.RESSORT With 2 extra switches ALS',
        )
        assert result.code == HS_ELECTRIC_ACTUATOR
        assert not result.unmatched

    def test_OS_name_오타가_결과를_바꾸지_않는다(self):
        """실측: SA05X가 OS=SR로, SR03이 OS=SA로 찍혀 있다 — 둘 다 액추에이터라 무해"""
        mislabeled = resolve_hs_code(
            os_name='SR',
            model_number='023310',
            item_name='SA05X 50Nm 17s. 230V50HZ ATEX Colour YELLOW 1018',
        )
        assert mislabeled.code == HS_ELECTRIC_ACTUATOR

    def test_대소문자와_공백은_무시한다(self):
        assert resolve_hs_code(os_name='  noah   na ').code == HS_ELECTRIC_ACTUATOR


class TestNonGoods:
    """운임 줄은 HS가 없다 — 그리고 경고 대상도 아니다"""

    def test_운임은_OS_name보다_먼저_걸러진다(self):
        """실측: FEDEX COST 라인의 OS name이 'Noah NA'/'SA'로 찍혀 있다.
        OS name을 먼저 보면 운임에 액추에이터 HS가 붙는다."""
        result = resolve_hs_code(os_name='Noah NA', item_name='FEDEX COST')
        assert result.code == ''
        assert not result.unmatched, "운임은 원래 HS가 없다 — 사람이 채울 줄이 아니다"

    @pytest.mark.parametrize('text', ['FEDEX COST', 'Freight charge', 'DHL COURIER', '운임'])
    def test_운임_표기_변형(self, text):
        assert resolve_hs_code(os_name='SA', item_name=text).code == ''


class TestSpares:
    """스페어는 모델번호가 안정 키 — 품목명은 프랑스어·영어가 섞여 흔들린다"""

    @pytest.mark.parametrize('model,item', [
        ('980748', 'CIRCUIT LED BOARD SA05/09 230VAC'),
        ('980750', 'CI PUIS.+LED SA05 230V 50Hz SA05 PC board + LED BOARD 230V 50Hz'),
        ('980770', 'CI NA-PCU 230V'),
        ('980772', 'CI CONVERTI 24VCA a 24VCC NA06 Converting Board AC to DC for NA06'),
        ('980773', 'CI+CONDENSATEUR SA05SCP 230VCA SA05 SCP Set 230V5 50HZ'),
        ('0812880001VCDE', 'CIRCUIT IMPRIME NA-PCU 24VCC'),
    ])
    def test_보드류_6종(self, model, item):
        result = resolve_hs_code(os_name='Electric Spares', model_number=model, item_name=item)
        assert result.code == HS_BOARDS
        assert not result.unmatched

    def test_엑셀이_숫자로_읽은_모델번호도_매칭(self):
        """980748이 float 980748.0으로 읽혀도 같은 줄로 봐야 한다"""
        assert resolve_hs_code(model_number=980748.0, item_name='CIRCUIT LED BOARD').code == HS_BOARDS

    @pytest.mark.parametrize('model,item', [
        ('980737', 'MOTEUR ELEC.SEUL NA09 230V50HZ Electric Motor for NA09 230V 50HZ'),
        ('980739', 'MOTEUR ELEC.SEUL NA15 230V50HZ Electric Motor for NA15 230V 50HZ'),
        ('980800', 'CAPOT TIGE MS 300MM Length of stem cover is 300mm'),
    ])
    def test_표에_없는_품목은_빈칸_경고(self, model, item):
        """전기모터·스템커버는 고객 표에 없다 — 추측하지 않고 사람에게 넘긴다"""
        result = resolve_hs_code(os_name='Electric Spares', model_number=model, item_name=item)
        assert result.code == ''
        assert result.unmatched


class TestKeywordRules:
    """아직 출고 이력이 없는 표 항목들 (미래 대비)"""

    def test_히터는_대상_모델로_코드가_갈린다(self):
        assert resolve_hs_code(
            os_name='Electric Spares', item_name='HEATER for NA009',
        ).code == HS_HEATER_SMALL
        assert resolve_hs_code(
            os_name='Electric Spares', item_name='HEATER for NA250',
        ).code == HS_HEATER_LARGE

    def test_히터_모델_경계(self):
        """NA006-NA009 / NA015-NA250 — 표의 경계 그대로"""
        assert resolve_hs_code(item_name='HEATER NA006').code == HS_HEATER_SMALL
        assert resolve_hs_code(item_name='HEATER NA015').code == HS_HEATER_LARGE

    def test_모델을_못_읽는_히터는_경고로_넘긴다(self):
        """어느 쪽 코드인지 모르는데 하나를 고르면 안 된다"""
        result = resolve_hs_code(os_name='Electric Spares', item_name='HEATER')
        assert result.code == ''
        assert result.unmatched

    @pytest.mark.parametrize('item,expected', [
        ('Inner Frame assembly for SA005', '85030099'),
        ('GOVERNOR for SR30', '85352900'),
        ('Battery Pack for NA/SA RBP', '85352900'),
        ('HAND WHEEL', '84835080'),
        ('NA038 Handwheel', '84835080'),
    ])
    def test_표_항목_키워드(self, item, expected):
        assert resolve_hs_code(os_name='Electric Spares', item_name=item).code == expected

    def test_히터가_보드로_새지_않는다(self):
        """보드 규칙이 가장 넓어서 순서가 뒤집히면 히터가 보드 코드를 받는다"""
        assert resolve_hs_code(item_name='HEATER CIRCUIT for NA015').code == HS_HEATER_LARGE


class TestEnrichHsCodes:
    """DataFrame 보강 — 코드는 행에 붙어 다녀야 한다 (생성기가 정렬하므로)"""

    FRAME = pd.DataFrame([
        {'Item': 'NA15 150Nm 20s. 230V50HZ', 'Model number': '023130', 'OS name': 'Noah NA'},
        {'Item': 'CIRCUIT LED BOARD SA05/09 230VAC', 'Model number': '980748', 'OS name': 'Electric Spares'},
        {'Item': 'MOTEUR ELEC.SEUL NA15 230V50HZ', 'Model number': '980739', 'OS name': 'Electric Spares'},
        {'Item': 'FEDEX COST', 'Model number': None, 'OS name': 'Noah NA'},
    ])

    def test_코드_열이_행_순서대로_붙는다(self):
        enriched, _ = enrich_hs_codes(self.FRAME)
        assert list(enriched[HS_COLUMN]) == [HS_ELECTRIC_ACTUATOR, HS_BOARDS, '', '']

    def test_미매칭만_경고에_담긴다(self):
        """운임은 빈칸이지만 경고가 아니다 — 모터만 사람이 정할 줄"""
        _, unmatched = enrich_hs_codes(self.FRAME)
        assert len(unmatched) == 1
        assert 'MOTEUR' in unmatched[0]
        assert '980739' in unmatched[0]

    def test_원본을_건드리지_않는다(self):
        enrich_hs_codes(self.FRAME)
        assert HS_COLUMN not in self.FRAME.columns


class TestDefaultCode:
    """거래처 고유 코드를 받지 않은 문서 — 우리 수출신고 코드를 일괄로 (양식 통일)"""

    DEFAULT = '8481.90.0000'

    def test_물품_줄은_기본값을_받는다(self):
        result = default_hs_result('NA15 150Nm 20s. 230V50HZ', self.DEFAULT)
        assert result.code == self.DEFAULT
        assert not result.unmatched

    def test_운임_줄은_기본값도_넣지_않는다(self):
        """물품이 아니라 HS가 애초에 없다 — 운송비에 밸브 부품 세번이 찍히면 안 된다"""
        result = default_hs_result('FEDEX COST', self.DEFAULT)
        assert result.code == ''
        assert not result.unmatched

    def test_기본값_경로는_거래처_판정표를_타지_않는다(self):
        """보드·액추에이터 구분 없이 전부 기본값 — 그 구분은 고객 코드표가 있을 때만"""
        frame = pd.DataFrame([
            {'Item': 'NA15 150Nm 20s. 230V50HZ', 'Model number': '023130', 'OS name': 'Noah NA'},
            {'Item': 'CIRCUIT LED BOARD SA05/09 230VAC', 'Model number': '980748',
             'OS name': 'Electric Spares'},
            {'Item': 'FEDEX COST', 'Model number': None, 'OS name': 'Noah NA'},
        ])
        enriched, unmatched = enrich_hs_codes(frame, default_code=self.DEFAULT)
        assert list(enriched[HS_COLUMN]) == [self.DEFAULT, self.DEFAULT, '']
        assert unmatched == [], "기본값 경로에는 '확인 필요'가 생기지 않는다"

    def test_설정값이_템플릿에_있던_코드_그대로다(self):
        """헤더에서 라인으로 옮긴 것이지 값을 바꾼 게 아니다"""
        from po_generator.config import DEFAULT_HS_CODE

        assert DEFAULT_HS_CODE == '8481.90.0000'


class TestActuatorCorpus:
    """2026년 SECTORIEL이 실제로 받은 액추에이터 **34개 계열 전수**

    품목명 규칙이 `OS name` 폴백보다 **앞에** 있으므로, 규칙 하나를 넓히면 이 186줄
    위를 지나간다. 규칙을 손대다 액추에이터를 조용히 스페어로 재분류하는 사고를
    잡는 유일한 그물이다 — 예: 모터 규칙을 추가하면 `S.MOT. NA200PCU`(servomoteur =
    액추에이터 본체)가 끌려 들어간다.
    """

    FAMILIES = [
        ('Noah NA', 'NA06 60Nm 17s. 230V50HZ With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NA09 90Nm 17s. 380V50HZ With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NA15 150Nm 20s. 230V50HZ With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NA28 280Nm 24s. 230V50HZ With 2 extra switches ALS + Star Bush 22'),
        ('Noah NA', 'NA38 380Nm 24s. 230V50HZ With 2 extra switches ALS + Star Bush 27'),
        ('Noah NA', 'NA60 600Nm 29s. 230V50HZ With 2 extra switches ALS + Star Bush 27'),
        ('Noah NA', 'NA100 1000Nm 29s. 380V50HZ With 2 extra switches ALS + Star Bush 27'),
        ('Noah NA', 'NA200 2000Nm 87s. 380V50HZ With 2 extra switches ALS + Star Bush 36'),
        ('Noah NA', 'NA250 2500Nm 87s. 230V50HZ With 2 extra switches ALS + Star Bush 46'),
        ('Noah NA', 'NA06PCU 60Nm 24V60HZ With 2 extra switches ALS + Star Bush 17 PCU'),
        ('Noah NA', 'NA09PCU 90Nm 17s. 230V50HZ With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NA15PCU 150Nm 20s. 230V50HZ With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NA28PCU 280Nm 24s. 230V50HZ With 2 extra switches ALS + Star Bush 22'),
        ('Noah NA', 'NAX06 60Nm 17s.230V50HZ ATEX'),
        ('Noah NA', 'NAX09 90Nm 17s.230V50HZ ATEX With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NAX15 150Nm 20s.230V50HZ ATEX With 2 extra switches ALS + Star Bush 17'),
        ('Noah NA', 'NAX28 280Nm 24s.24V50HZ ATEX With 2 extra switches ALS + Star Bush 22'),
        ('Noah NA', 'NAX38 380Nm 24s.230V50HZ ATEX With 2 extra switches ALS + Star Bush 27'),
        ('Noah NA', 'NAX60 600Nm 29s.230V50HZ ATEX With 2 extra switches ALS + Star Bush 27'),
        ('Noah NA', 'S.MOT. NA200PCU 2000Nm 230VAC ALS Star Bush 36'),
        ('SA', 'SA03 30Nm 12s. 230V50HZ Colour YELLOW 1018'),
        ('SA', 'SA05 50Nm 17s. 230V50HZ Colour YELLOW 1018'),
        ('SA', 'SA05S 50Nm 100s. 24VCA/CC Colour YELLOW 1018'),
        ('SA', 'SA05PCU 50Nm 17s. 230V50HZ Colour YELLOW 1018'),
        ('SA', 'SA09 90Nm 22s. 24Vca/cc'),
        ('SA', 'SA09X 90Nm 230VCA/CC 32s ATEX PCU YELLOW 1018'),
        ('SR', 'SA05X 50Nm 17s. 230V50HZ ATEX Colour YELLOW 1018'),
        ('SR', 'SR03 230V50HZ RA.RESSORT PCU'),
        ('SR', 'SR05 50Nm 230V50HZ RA.RESSORT With 2 extra switches ALS'),
        ('SR', 'SR10 100Nm 230V50HZ RA.RESSORT With 2 extra switches ALS + Star Bush 17'),
        ('SR', 'SR10X 100Nm 230V50HZ RAZ ATEX With 2 extra switches ALS F07/F10 Star Bush 17 - Fail Close ATEX'),
        ('SR', 'SR30 230VAC ATEX F/O ALS 22STAR'),
        ('SR', 'SR30X 300Nm 230V50HZ RAZ ATEX With 2 extra switches ALS'),
        ('MS', 'MS 110Nm 230V 50/60Hz 1PH (220/230) F10 16.6rpm (50Hz) - 20.1 (60Hz) - 0.55kW'),
    ]

    @pytest.mark.parametrize('os_name,item', FAMILIES)
    def test_모든_액추에이터_계열이_본체_코드를_받는다(self, os_name, item):
        result = resolve_hs_code(os_name=os_name, item_name=item)
        assert result.code == HS_ELECTRIC_ACTUATOR, f"{item} → {result.rule}"
        assert not result.unmatched


class TestTableIntegrity:
    """비프로그래머가 표에 줄을 추가할 때의 안전망"""

    def test_모든_코드는_8자리_문자열(self):
        """int로 두면 0으로 시작하는 코드가 들어왔을 때 앞자리가 조용히 날아간다"""
        from po_generator import hs_code as mod

        codes = set(mod.SPARE_HS_BY_MODEL.values()) | {
            mod.HS_ELECTRIC_ACTUATOR, mod.HS_HEATER_SMALL, mod.HS_HEATER_LARGE,
            mod.HS_INNER_FRAME, mod.HS_BOARDS, mod.HS_GOVERNOR,
            mod.HS_BATTERY_PACK, mod.HS_HANDLE_WHEEL,
        }
        for code in codes:
            assert isinstance(code, str) and code.isdigit() and len(code) == 8, code

    def test_모델키는_정규화된_형태다(self):
        """표에 'CI 980748 ' 처럼 넣으면 조회 키와 안 맞아 조용히 미매칭된다"""
        from po_generator.hs_code import SPARE_HS_BY_MODEL, _norm_model

        for key in SPARE_HS_BY_MODEL:
            assert key == _norm_model(key), key

    def test_코드가_있으면_경고_대상이_아니다(self):
        """'코드도 주고 확인도 요청'하는 어정쩡한 결과가 나오면 안 된다"""
        from po_generator.hs_code import SPARE_HS_BY_MODEL

        for model in SPARE_HS_BY_MODEL:
            result = resolve_hs_code(model_number=model, item_name='x')
            assert result.code and not result.unmatched
