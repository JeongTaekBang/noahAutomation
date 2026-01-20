@echo off
REM ============================================================================
REM 사용자별 로컬 설정 파일 (예제)
REM ============================================================================
REM
REM 사용법:
REM   1. 이 파일을 local_config.bat 으로 복사
REM   2. 본인 환경에 맞게 PYTHON_PATH 수정
REM   3. local_config.bat 은 Git에 올라가지 않음
REM
REM ============================================================================

REM Python 실행 경로 설정
REM - miniconda: %LOCALAPPDATA%\miniconda3\envs\po-automate\python.exe
REM - anaconda:  C:\Users\사용자\anaconda3\envs\po-automate\python.exe

set PYTHON_PATH=%LOCALAPPDATA%\miniconda3\envs\po-automate\python.exe
