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
