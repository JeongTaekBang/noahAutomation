"""
문서 레이아웃 계산·감시 테스트 (Excel 없이)
=============================================

해외 문서 5종(OC·FI·PI·CI·PL)의 레이아웃 규칙을 COM 없이 지킨다:

- `address_row_heights()` — 주소 블록 행 높이 (순수 함수)
- 레이아웃 상수 불변식 (병합 열·최소 행 높이·여유)
- 생성기들이 공용 헬퍼(`layout_item_rows`/`layout_address_rows`)를 쓰는지 소스 감시

(예전의 "한 페이지 빈 행 채우기" 테스트는 기능 제거와 함께 사라졌다 —
2026-08-05, 1아이템 문서가 빈 격자를 달고 나가 표를 아이템에서 끝내기로 결정.)
"""

import pytest
from pathlib import Path

from po_generator.excel_helpers import (
    HS_COLUMN_LETTER,
    ITEM_NAME_MERGED_COLS,
    ITEM_NAME_MERGED_COLS_HS,
    MIN_ITEM_ROW_HEIGHT,
    ROW_HEIGHT_PAD,
    _item_name_cols,
    address_row_heights,
)


class TestAddressRowHeights:
    """주소 블록 행 높이 계산 (순수 함수) — 기준값은 OC 템플릿 실측치

    주소 행 3개는 각 15.95pt. 왼쪽은 행별 한 줄 병합(A:E), 오른쪽은 3행을
    세로로 걸친 병합(G13:I15) 하나다.
    """

    CUR = [15.95, 15.95, 15.95]

    def test_짧은_주소는_현재_높이를_유지한다(self):
        """한 줄 측정치(~13pt)+여유가 15.95를 넘지 않으면 모양이 안 변한다"""
        assert address_row_heights(self.CUR, [13.0, 13.0, 13.0], 0.0, pad=2.0) == self.CUR

    def test_접힌_왼쪽_줄만_자란다(self):
        """SECTORIEL 실측: 74자 bill_to_2가 2줄(26pt)로 접히면 그 행만 커진다"""
        heights = address_row_heights(self.CUR, [13.0, 26.0, 13.0], 0.0, pad=2.0)
        assert heights == [15.95, 28.0, 15.95]

    def test_오른쪽_블록_부족분은_균등_분배된다(self):
        """납품 주소가 3행 합(47.85pt)보다 크면 세 행이 같이 자란다"""
        heights = address_row_heights(self.CUR, [13.0] * 3, 60.0, pad=2.0)
        assert sum(heights) == pytest.approx(62.0)
        assert heights[0] == pytest.approx(heights[1]) == pytest.approx(heights[2])

    def test_오른쪽이_행_합에_들어가면_분배하지_않는다(self):
        """2줄 납품 주소(~28pt)는 3행 합 47.85pt에 이미 들어간다 — SECTORIEL 실측"""
        assert address_row_heights(self.CUR, [13.0] * 3, 28.0, pad=2.0) == self.CUR

    def test_왼쪽_성장으로_이미_충분하면_더_키우지_않는다(self):
        heights = address_row_heights(self.CUR, [13.0, 40.0, 13.0], 50.0, pad=2.0)
        assert heights == [15.95, 42.0, 15.95]  # 합 73.9 >= 52

    def test_두_요구를_동시에_만족한다(self):
        """어느 쪽이 이기든 최종 합은 오른쪽 요구 이상, 각 행은 왼쪽 요구 이상"""
        for right in (0.0, 30.0, 55.0, 90.0):
            heights = address_row_heights(self.CUR, [13.0, 26.0, 13.0], right, pad=2.0)
            assert all(h >= n + 2.0 for h, n in zip(heights, [13.0, 26.0, 13.0]))
            assert all(h >= c for h, c in zip(heights, self.CUR))
            if right > 0:
                assert sum(heights) >= right + 2.0 - 1e-9

    def test_길이가_어긋나면_조용히_진행하지_않는다(self):
        with pytest.raises(ValueError):
            address_row_heights([15.95, 15.95], [13.0] * 3, 0.0)

    def test_빈_입력은_빈_결과(self):
        assert address_row_heights([], [], 60.0) == []


class TestGeneratorsLayoutAddresses:
    """주소를 쓰는 생성기(OC·FI)가 공용 주소 레이아웃을 부르는지 감시

    주소 칸은 병합 셀이라 wrap 없이 길면 병합 경계에서 잘린다(2026-08-05 실측
    74자/81자 클립). 값만 쓰고 layout_address_rows를 빼먹으면 재발한다.
    PI·CI·PL은 주소 셀을 채우지 않아 대상이 아니다.
    """

    @pytest.mark.parametrize('module_name', [
        'po_generator.oc_generator',
        'po_generator.fi_generator',
    ])
    def test_주소를_쓰면_layout_address_rows를_부른다(self, module_name):
        import importlib
        import inspect

        source = inspect.getsource(importlib.import_module(module_name))
        assert 'layout_address_rows(' in source, (
            f"{module_name}: 주소 셀을 채우면서 layout_address_rows를 부르지 않는다 — "
            f"긴 주소가 병합 경계에서 잘린다"
        )


class TestLayoutConstants:
    def test_최소_행_높이는_템플릿_기본값(self):
        assert MIN_ITEM_ROW_HEIGHT == 15.0

    def test_여유는_양수다(self):
        """병합 셀은 안쪽 여백이 미세하게 커서 딱 맞추면 마지막 줄이 잘린다"""
        assert ROW_HEIGHT_PAD > 0

    def test_병합_열은_문서_5종_공통이다(self):
        """CI와 PL은 늘 같이 첨부되므로 두 문서가 같은 값을 써야 한다"""
        assert ITEM_NAME_MERGED_COLS == 'ABCD'

    def test_HS판은_D만_떼어낸_A_C다(self):
        """라인별 HS 판의 품목명은 A:C — 떨어져 나온 D가 곧 HS 열이다"""
        assert ITEM_NAME_MERGED_COLS_HS == 'ABC'
        assert ITEM_NAME_MERGED_COLS.startswith(ITEM_NAME_MERGED_COLS_HS)
        assert ITEM_NAME_MERGED_COLS[len(ITEM_NAME_MERGED_COLS_HS):] == HS_COLUMN_LETTER

    def test_판은_둘뿐이고_헬퍼가_고른다(self):
        """호출부는 bool만 넘긴다 — 열 문자열을 직접 넘길 수 있으면 CI와 PL이 갈린다"""
        assert _item_name_cols(False) == ITEM_NAME_MERGED_COLS
        assert _item_name_cols(True) == ITEM_NAME_MERGED_COLS_HS

    def test_생성기들이_병합_열을_재정의하지_않는다(self):
        """한쪽만 바뀌면 같은 봉투 안의 두 장이 다른 줄간격으로 나간다"""
        import po_generator.ci_generator as ci
        import po_generator.fi_generator as fi
        import po_generator.oc_generator as oc
        import po_generator.pi_generator as pi
        import po_generator.pl_generator as pl

        for mod in (oc, fi, pi, ci, pl):
            assert not hasattr(mod, 'ITEM_NAME_MERGED_COLS'), (
                f"{mod.__name__}이 병합 열을 따로 정의하고 있다 — "
                f"excel_helpers.ITEM_NAME_MERGED_COLS 하나만 쓸 것"
            )


class TestGeneratorsUseSharedHelper:
    """다섯 생성기가 `rows.autofit()`으로 돌아가지 않았는지 감시

    병합 셀에는 그게 먹지 않는다 — 오히려 1줄로 줄여 긴 품목명을 잘라낸다.
    """

    @pytest.mark.parametrize('module_name', [
        'po_generator.oc_generator',
        'po_generator.fi_generator',
        'po_generator.pi_generator',
        'po_generator.ci_generator',
        'po_generator.pl_generator',
    ])
    def test_아이템_영역에_rows_autofit을_쓰지_않는다(self, module_name):
        import importlib
        import inspect

        source = inspect.getsource(importlib.import_module(module_name))
        offending = [
            line.strip() for line in source.splitlines()
            if 'rows.autofit()' in line and not line.strip().startswith('#')
        ]
        assert not offending, f"{module_name}: {offending}"

    @pytest.mark.parametrize('module_name', [
        'po_generator.oc_generator',
        'po_generator.fi_generator',
        'po_generator.pi_generator',
        'po_generator.ci_generator',
        'po_generator.pl_generator',
    ])
    def test_복합_헬퍼로_병합과_행높이를_처리한다(self, module_name):
        """layout_item_rows = 병합 보장 + 행 높이 교정 세트 — 한쪽만 부르는 실수 차단"""
        import importlib
        import inspect

        source = inspect.getsource(importlib.import_module(module_name))
        assert 'layout_item_rows(' in source, f"{module_name}: layout_item_rows 미사용"


class TestHsLayoutKeepsExistingColumns:
    """HS 열은 기존 열을 건드리지 않고 넣는다

    오른쪽 블록에서 폭을 걷어 보려다 두 번 사고가 났다 (2026-09-02·04 실측):
    헤더가 긴 열이 잘리고('Quantity'→'Quantit'), 값을 쓰기 전에 자동 맞춤을 돌려
    `PO No.` 열이 줄어 긴 고객 PO가 잘렸다. 지금은 초과분을 인쇄 배율로 흡수한다.
    """

    def test_기부_열_설정이_남아_있지_않다(self):
        """폭을 걷는 정책 자체를 없앴다 — 상수가 남아 있으면 되살아난다"""
        from po_generator import config

        for name in ('CI_HS_WIDTH_DONORS', 'PL_HS_WIDTH_DONORS'):
            assert not hasattr(config, name), f"{name}: 폭 걷기 정책은 제거됐다"

    def test_레이아웃_헬퍼가_기부_열을_받지_않는다(self):
        import inspect

        from po_generator.excel_helpers import apply_hs_layout

        params = inspect.signature(apply_hs_layout).parameters
        assert 'width_donors' not in params


class TestGeneratorsDoNotHardcodeColumns:
    """생성기가 병합 열 문자열을 직접 들고 있으면 CI와 PL이 갈린다"""

    @pytest.mark.parametrize('module_name', [
        'po_generator.ci_generator',
        'po_generator.pl_generator',
    ])
    def test_병합_열_리터럴이_없다(self, module_name):
        import importlib
        import inspect

        source = inspect.getsource(importlib.import_module(module_name))
        code_lines = [
            line for line in source.splitlines()
            if not line.strip().startswith('#')
        ]
        for literal in ("'ABCD'", '"ABCD"', "'ABC'", '"ABC"'):
            offending = [ln.strip() for ln in code_lines if literal in ln]
            assert not offending, f"{module_name}: {literal} 하드코딩 — {offending}"
