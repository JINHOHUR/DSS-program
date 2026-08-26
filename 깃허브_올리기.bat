@echo off
chcp 65001 >nul
cd /d "%~dp0"
title GitHub 업로드

echo.
echo  ============================================================
echo   견적서 프로그램을 GitHub 에 올립니다
echo   https://github.com/JINHOHUR/DSS-program
echo  ============================================================
echo.
echo  [1/2] 현재 상태
git log --oneline -3
echo.
echo  [2/2] 업로드 시작
echo.
echo  * 로그인 창이 뜨면 "Sign in with your browser" 를 고르세요.
echo  * 브라우저에서 GitHub 로그인만 하면 됩니다.
echo.

git push -u origin main

echo.
if errorlevel 1 (
  echo  ------------------------------------------------------------
  echo   실패했습니다. 위 오류 메시지를 확인해 주세요.
  echo  ------------------------------------------------------------
) else (
  echo  ------------------------------------------------------------
  echo   성공! 아래 주소에서 확인하세요.
  echo   https://github.com/JINHOHUR/DSS-program
  echo  ------------------------------------------------------------
)
echo.
pause
