#!/bin/bash
cd "$(dirname "$0")"
echo "============================================================"
echo "  教师问卷在线答题与数据汇总系统"
echo "------------------------------------------------------------"
echo "  正在启动，请稍候……"
echo "  启动后会自动打开数据后台页面。"
echo "  停止服务请按 Ctrl+C。"
echo "============================================================"
(
  sleep 2
  if command -v open >/dev/null 2>&1; then open http://127.0.0.1:8000/admin
  elif command -v xdg-open >/dev/null 2>&1; then xdg-open http://127.0.0.1:8000/admin; fi
) &
python3 server.py
