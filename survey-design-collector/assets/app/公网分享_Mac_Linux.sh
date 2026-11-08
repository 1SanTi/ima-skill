#!/bin/bash
cd "$(dirname "$0")"
echo "============================================================"
echo "  问卷系统 - 公网分享（Mac / Linux，免下载免注册）"
echo "============================================================"
python3 server.py > /tmp/survey_server.log 2>&1 &
echo "[1/2] 本机系统已启动 (http://127.0.0.1:8000)"
sleep 3
echo "[2/2] 正在建立公网地址 ..."
echo "  下面会出现一行：Forwarding HTTP traffic from https://xxxx.serveousercontent.com"
echo "  把该地址复制出来，浏览器打开  该地址/admin  即可看到二维码。"
echo "  （答题期间请勿关闭本窗口，Ctrl+C 结束）"
echo "------------------------------------------------------------"
ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=/dev/null -o ServerAliveInterval=30 -R 80:localhost:8000 serveo.net
