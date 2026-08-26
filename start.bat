@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  견적서 프로그램을 시작합니다...
echo.
python backend\server.py
if errorlevel 1 (
  echo.
  echo  [오류] 실행에 실패했습니다. python 이 PATH 에 있는지 확인하세요.
  pause
)
