#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/opt/991x"
APP_NAME="991x"
PORT="${PORT:-9910}"
HOST="${HOST:-127.0.0.1}"
NGINX_SITE="/etc/nginx/conf.d/991x.conf"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ ! -d "$APP_DIR" ]]; then
  echo "未找到应用目录: $APP_DIR"
  echo "请先将 deploy/tencent-lighthouse-upload/991x 上传到 $APP_DIR"
  exit 1
fi

echo "[1/6] 安装系统依赖..."
sudo apt update
sudo apt install -y nginx curl ca-certificates

echo "[2/6] 安装 Node.js 20（如未安装）..."
if ! command -v node >/dev/null 2>&1; then
  curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
  sudo apt install -y nodejs
fi

echo "[3/6] 安装 PM2（如未安装）..."
if ! command -v pm2 >/dev/null 2>&1; then
  sudo npm install -g pm2
fi

echo "[4/6] 安装应用依赖..."
cd "$APP_DIR"
if [[ -f package-lock.json ]]; then
  npm ci --omit=dev
else
  npm install --omit=dev
fi

echo "[5/6] 启动应用..."
pm2 delete "$APP_NAME" >/dev/null 2>&1 || true
PORT="$PORT" HOST="$HOST" pm2 start server.js --name "$APP_NAME" --update-env
pm2 save
pm2 startup systemd -u "$USER" --hp "$HOME" >/dev/null || true

echo "[6/6] 配置 Nginx..."
if [[ -f "$SCRIPT_DIR/nginx-991x.conf" ]]; then
  sudo cp "$SCRIPT_DIR/nginx-991x.conf" "$NGINX_SITE"
  sudo sed -i "s/__APP_PORT__/$PORT/g" "$NGINX_SITE"
  sudo nginx -t
  sudo systemctl enable nginx
  sudo systemctl restart nginx
else
  echo "未找到 nginx-991x.conf，跳过 Nginx 配置。"
fi

echo "部署完成。"
echo "PM2 状态："
pm2 status
