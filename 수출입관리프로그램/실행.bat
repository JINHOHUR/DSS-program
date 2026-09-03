@echo off
title DSS 수출입 관리
cd /d "%~dp0"

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "dss_trade_manager.py"
    goto :eof
)

where python >nul 2>nul
if not %errorlevel%==0 (
    echo.
    echo  [오류] Python 을 찾을 수 없습니다.
    echo  https://www.python.org/downloads/ 에서 설치한 뒤 다시 실행하세요.
    echo  설치 시 "Add Python to PATH" 를 반드시 체크하세요.
    echo.
    pause
    goto :eof
)

python "dss_trade_manager.py"
if errorlevel 1 (
    echo.
    echo  [오류] 프로그램 실행 중 문제가 발생했습니다. 위 메시지를 확인하세요.
    echo.
    pause
)
