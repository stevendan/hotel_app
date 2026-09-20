@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

rem ==========================================================
rem  Chay trang contract/index.html bang server local (Python)
rem  - Tu dong tim Python
rem  - Tu dong tim port trong (bat dau tu 8000)
rem  - Mo trinh duyet, giu cua so nay de server chay
rem  Dong cua so hoac nhan Ctrl+C de tat server
rem ==========================================================

cd /d "%~dp0"

rem --- Tim Python ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python --version >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo [LOI] Khong tim thay Python tren may.
    echo Cai dat tai https://www.python.org/downloads/ roi chay lai file nay.
    echo.
    pause
    exit /b 1
)

rem --- Tim port trong, bat dau tu 8000 ---
set "PORT="
for /l %%P in (8000,1,8020) do (
    if not defined PORT (
        netstat -ano -p tcp | findstr /r /c:":%%P .*LISTENING" >nul 2>&1
        if errorlevel 1 set "PORT=%%P"
    )
)
if not defined PORT (
    echo [LOI] Khong tim duoc port trong trong khoang 8000-8020.
    echo.
    pause
    exit /b 1
)

set "URL=http://localhost:%PORT%/index.html"

echo.
echo ====================================================
echo   Server:  %URL%
echo   Thu muc: %CD%
echo   Nhan Ctrl+C hoac dong cua so de tat server
echo ====================================================
echo.

rem --- Mo trinh duyet sau 1.5 giay de server kip khoi dong ---
start "" /b cmd /c "timeout /t 2 /nobreak >nul & start "" "%URL%""

rem --- Chay server (chay o foreground, giu cua so) ---
%PY% -m http.server %PORT% --bind 127.0.0.1

echo.
echo Server da dung.
pause
