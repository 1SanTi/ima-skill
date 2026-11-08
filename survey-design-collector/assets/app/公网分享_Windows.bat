@echo off
cd /d "%~dp0"
title 问卷系统 - 公网分享（请勿关闭本窗口）
set PY=python
where python >nul 2>nul || set PY=py

echo ============================================================
echo    问卷系统 - 公网分享（免下载、免注册）
echo ============================================================
echo  [1/2] 正在本机启动问卷系统 ...
start "问卷系统服务(答题期间勿关)" cmd /k "%PY% server.py"
timeout /t 3 >nul

echo  [2/2] 正在建立公网地址 ...
echo.
echo   >>> 稍等约 5-15 秒，下面会出现一行：
echo         Forwarding HTTP traffic from https://xxxx.serveousercontent.com
echo.
echo   >>> 把那段  https://....serveousercontent.com  复制出来，
echo       在浏览器里打开：  该地址后面加上 /admin
echo       例：https://xxxx.serveousercontent.com/admin
echo       这个页面上的二维码，就是公网地址，直接发到群里！
echo.
echo   注意：本窗口和系统窗口，在答题期间请勿关闭。
echo   （若下面长时间没有出现地址，说明当前网络不通 serveo，
echo     请改用 cpolar，见  部署到公网\内网穿透快速方案.md）
echo ============================================================
ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=NUL -o ServerAliveInterval=30 -R 80:localhost:8000 serveo.net

echo.
echo 公网分享已结束（隧道已断开）。
pause
