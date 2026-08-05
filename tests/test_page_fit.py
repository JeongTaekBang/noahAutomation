"""
인쇄 레이아웃 계산 테스트 (Excel 없이)
========================================

`fit_blank_rows()`는 COM과 분리된 순수 함수라 Excel 없이 검증할 수 있다.
기준값은 OC 템플릿 실측치다 (2026-08-03, xlwings PageSetup):

    인쇄 가능 높이 (A4 세로, 위·아래 여백 54pt)  733.89 pt
    헤더 rows 1-17 (인쇄제목으로 매 페이지 반복)   265.00 pt
    하단 rows 25-43 (Total + 은행정보 + 약관)      335.25 pt
    아이템 행 기본                                  15.00 pt

→ 아이템에 쓸 수 있는 높이 = 733.89 - 265 - 335.25 = 133.64pt = 15pt 행 8개
"""

import pytest

from po_generator.excel_helpers import (
    ITEM_NAME_MERGED_COLS,
    MIN_ITEM_ROW_HEIGHT,
    ROW_HEIGHT_PAD,
    address_row_heights,
    fit_blank_rows,
)

# OC 템플릿 실측치
PRINTABLE = 733.89
HEADER = 265.00
FOOTER = 335.25
ROW = 15.0

# 아이템 표에 쓸 수 있는 높이 (= 133.64pt) — fit_blank_rows의 available 인자
AVAILABLE = PRINTABLE - HEADER - FOOTER


def fit(item_heights, **kw):
    return fit_blank_rows(AVAILABLE, item_heights, blank_height=ROW, **kw)


class TestFitBlankRows:
    def test_한_페이지에_들어가는_아이템_행은_8개(self):
        """실측 기준: 133.64 / 15 = 8.9 -> 8행"""
        assert int(AVAILABLE // ROW) == 8

    def test_아이템_1개면_7행을_채운다(self):
        """가장 흔한 경우 — 표 한 줄만 남고 아래가 텅 비는 것을 막는다"""
        assert fit([ROW]) == 7

    def test_아이템_8개면_더_채우지_않는다(self):
        assert fit([ROW] * 8) == 0

    def test_아이템_7개면_1행_남는다(self):
        assert fit([ROW] * 7) == 1

    def test_채운_뒤_총_높이가_인쇄_가능_높이를_넘지_않는다(self):
        for n in range(1, 9):
            items = [ROW] * n
            blanks = fit(items)
            used = HEADER + FOOTER + sum(items) + blanks * ROW
            assert used <= PRINTABLE, f"{n}개 아이템에서 {used:.2f}pt로 넘침"

    def test_긴_품목명이_섞이면_채울_행이_준다(self):
        """3줄짜리(38pt) 품목명 하나가 15pt 행 약 1.5개를 먹는다"""
        assert fit([38.0]) < fit([ROW])

    def test_긴_품목명만_있으면_채우지_않는다(self):
        """38pt * 4 = 152pt > 133.64pt — 이미 한 페이지를 넘겼다"""
        assert fit([38.0] * 4) == 0

    def test_이미_넘친_문서는_채우지_않는다(self):
        """48아이템 OC 같은 다중 페이지 문서에 빈 격자를 더하면 안 된다"""
        assert fit([ROW] * 48) == 0

    def test_딱_맞으면_채우지_않는다(self):
        """남는 높이가 한 행에 못 미치면 0"""
        exact = [AVAILABLE - ROW + 0.1]
        assert fit(exact) == 0

    def test_아이템이_없어도_안전하다(self):
        assert fit([]) >= 0

    def test_행_높이가_0이면_무한루프_대신_0(self):
        assert fit_blank_rows(AVAILABLE, [], blank_height=0) == 0
        assert fit_blank_rows(AVAILABLE, [], blank_height=-5) == 0

    def test_상한을_넘지_않는다(self):
        """템플릿이 이상해도 무한정 채우지 않는다"""
        assert fit_blank_rows(100_000, [], blank_height=ROW, max_rows=10) == 10

    def test_반환값은_음수가_아니다(self):
        """쓸 수 있는 높이가 이미 음수여도 0 (넘친 문서)"""
        assert fit_blank_rows(-500.0, [15.0], blank_height=ROW) == 0


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
