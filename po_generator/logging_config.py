"""
로깅 설정 모듈
==============

중앙화된 로깅 설정을 제공합니다.
모든 CLI에서 동일한 로깅 설정을 사용합니다.
"""

from __future__ import annotations

import logging
import sys


def setup_logging(verbose: bool = False) -> None:
    """로깅 설정

    print()는 사용자 출력(테이블, 진행상황)에 사용하고,
    logging은 오류 추적, 디버그, 검증 메시지에 사용합니다.

    Args:
        verbose: 상세 로깅 여부 (True면 DEBUG, False면 INFO)
    """
    level = logging.DEBUG if verbose else logging.INFO

    # 기존 핸들러 제거 (중복 방지)
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # 콘솔 핸들러
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)

    # 포맷 설정 (verbose 모드에서만 상세 정보)
    if verbose:
        console_format = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
    else:
        console_format = logging.Formatter('%(message)s')

    console_handler.setFormatter(console_format)

    # 루트 로거 설정
    root_logger.setLevel(level)
    root_logger.addHandler(console_handler)

    # po_generator 패키지 로거 설정
    pkg_logger = logging.getLogger('po_generator')
    pkg_logger.setLevel(level)

    if verbose:
        logging.debug("로깅 설정 완료 (verbose mode)")
