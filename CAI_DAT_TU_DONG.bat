@echo off
title CAI DAT DASHBOARD TU DONG KHOI DONG
echo ============================================================
echo   DANG THIET LAP TU DONG CHAY DASHBOARD CUNG WINDOWS
echo ============================================================
echo.

set "TARGET_DIR=%~dp0"
if "%TARGET_DIR:~-1%"=="\" set "TARGET_DIR=%TARGET_DIR:~0,-1%"

set "STARTUP_VBS=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\start_dashboard_server.vbs"

echo Set WshShell = CreateObject("WScript.Shell") > "%STARTUP_VBS%"
echo WshShell.CurrentDirectory = "%TARGET_DIR%" >> "%STARTUP_VBS%"
echo WshShell.Run "pythonw server.py", 0, False >> "%STARTUP_VBS%"

echo [1/3] Da dang ky Server vao Windows Startup thanh cong!
echo.

echo [2/3] Dang khoi dong Python Server ngam...
wscript "%STARTUP_VBS%"

echo [3/3] Dang mo trang web http://localhost:8080 ...
ping -n 3 127.0.0.1 >nul
start "" "http://localhost:8080"

echo.
echo ============================================================
echo   CAI DAT HOAN TAT!
echo   Tu nay ve sau, moi khi bat may tinh len, ban chi can 
echo   mo trinh duyet go localhost:8080 la se tu dong chay.
echo ============================================================
echo.
pause
