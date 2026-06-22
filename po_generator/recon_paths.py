"""Reconciliation 폴더 경로 헬퍼.

플랫(`RECON_DIR/{period}/`) / 연도 중첩(`RECON_DIR/{year}/{period}/`)
두 가지 레이아웃을 모두 지원한다.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Iterator

logger = logging.getLogger(__name__)

_PERIOD_RE = re.compile(r'^P\d+$', re.IGNORECASE)
_YEAR_RE = re.compile(r'^\d{4}$')


def resolve_period_dir(recon_dir: Path, period_code: str) -> Path | None:
    """주어진 period 디렉터리 경로를 반환.

    `RECON_DIR/{period}` 우선, 못 찾으면 연도 폴더 중 최신 연도부터 탐색.
    여러 연도에서 같은 period가 발견되면 최신 연도를 선택하고 경고를 남긴다.
    어디에도 없으면 None.
    """
    period = period_code.upper()
    direct = recon_dir / period
    # 플랫 폴더는 실제로 파일이 있을 때만 우선 — 빈 폴더가 연도 중첩 폴더를
    # 가리지 않도록 (빈 플랫 폴더면 아래 연도 스캔으로 폴백).
    direct_populated = direct.is_dir() and any(direct.glob('*.xlsx'))
    if not recon_dir.exists():
        return direct if direct_populated else None
    matches: list[Path] = []
    for sub in sorted(recon_dir.iterdir(), key=lambda p: p.name, reverse=True):
        if sub.is_dir() and _YEAR_RE.match(sub.name):
            cand = sub / period
            if cand.is_dir():
                matches.append(cand)
    if direct_populated:
        if matches:
            logger.warning(
                "[경고] %s/%s 플랫 폴더와 연도 폴더(%s) 양쪽에 존재 — 플랫 폴더 우선 사용",
                recon_dir.name, period, ", ".join(m.parent.name for m in matches),
            )
        return direct
    if not matches:
        # 연도 매치는 없지만 빈 플랫 폴더라도 존재하면 그 경로 반환 (기존 동작 유지).
        return direct if direct.is_dir() else None
    chosen = matches[0]
    if len(matches) > 1:
        others = ", ".join(m.parent.name for m in matches[1:])
        logger.warning(
            "[경고] %s/%s 가 여러 연도에 존재 — %s 사용 (다른 연도: %s)",
            recon_dir.name, period, chosen.parent.name, others,
        )
    return chosen


def iter_period_dirs(recon_dir: Path) -> Iterator[Path]:
    """플랫/연도 중첩 양쪽에서 period 디렉터리(P\\d+)를 모두 yield."""
    if not recon_dir.exists():
        return
    for sub in recon_dir.iterdir():
        if not sub.is_dir():
            continue
        if _PERIOD_RE.match(sub.name):
            yield sub
        elif _YEAR_RE.match(sub.name):
            for sub2 in sub.iterdir():
                if sub2.is_dir() and _PERIOD_RE.match(sub2.name):
                    yield sub2
