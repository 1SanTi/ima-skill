#!/bin/bash
# ============================================================
#  问卷系统 · 一键部署到 Linux 云服务器（Ubuntu / Debian / CentOS）
#  用法：把整个「在线问卷系统」文件夹上传到服务器后，在该文件夹下执行
#        bash 部署到公网/install.sh
#  可选：自定义端口  PORT=8080 bash 部署到公网/install.sh
# ============================================================
set -e
APP_NAME="survey"
DEST="/opt/$APP_NAME"
PORT="${PORT:-80}"
SRC="$(cd "$(dirname "$0")/.." && pwd)"

echo "==> 1/5 检查 Python 环境"
if ! command -v python3 >/dev/null 2>&1; then
  if command -v apt-get >/dev/null 2>&1; then
    apt-get update -y && apt-get install -y python3 curl
  elif command -v yum >/dev/null 2>&1; then
    yum install -y python3 curl
  else
    echo "请先手动安装 python3 后重试"; exit 1
  fi
fi
python3 --version

echo "==> 2/5 复制程序到 $DEST"
mkdir -p "$DEST" "$DEST/data"
cp -f "$SRC/server.py" "$DEST/"
cp -f "$SRC/questions.json" "$DEST/"
rm -rf "$DEST/vendor"; cp -r "$SRC/vendor" "$DEST/vendor"

echo "==> 3/5 注册开机自启服务"
cat > /etc/systemd/system/$APP_NAME.service <<EOF
[Unit]
Description=Survey Server (问卷在线收集系统)
After=network.target

[Service]
WorkingDirectory=$DEST
Environment=PORT=$PORT
ExecStart=/usr/bin/env python3 $DEST/server.py
Restart=always
RestartSec=3
User=root

[Install]
WantedBy=multi-user.target
EOF

echo "==> 4/5 启动服务"
systemctl daemon-reload
systemctl enable $APP_NAME >/dev/null 2>&1
systemctl restart $APP_NAME

echo "==> 5/5 读取公网 IP"
IP=$(curl -s --max-time 6 ifconfig.me || curl -s --max-time 6 ip.sb || true)
[ -z "$IP" ] && IP="你的服务器公网IP"

echo ""
echo "============================================================"
echo "  部署完成！"
echo "  老师答题：  http://$IP/"
echo "  数据后台：  http://$IP/admin"
echo "------------------------------------------------------------"
echo "  若打不开：请到云服务器控制台【安全组/防火墙】放行 $PORT 端口"
echo "  查看状态：systemctl status $APP_NAME"
echo "  重启服务：systemctl restart $APP_NAME"
echo "  数据文件：$DEST/data/responses.db（备份即复制此文件）"
echo "============================================================"
