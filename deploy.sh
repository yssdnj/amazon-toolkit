#!/bin/bash
# deploy.sh — 拉取最新代码并重启服务

set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP="app.py"
PORT=5002
BRANCH="dev"

echo "========================================"
echo "  Amazon Toolkit 部署脚本"
echo "  目录: $PROJECT_DIR"
echo "========================================"

# 1. 拉取最新代码
echo ""
echo "▶ 拉取代码 (origin/$BRANCH)..."
cd "$PROJECT_DIR"
git checkout -- deploy.sh
git pull origin "$BRANCH"

# 2. 停止旧进程
echo ""
echo "▶ 停止旧服务..."
pkill -f "python3 $APP" 2>/dev/null || true
OLD_PID=$(lsof -t -i:$PORT 2>/dev/null || true)
if [ -n "$OLD_PID" ]; then
    kill -9 $OLD_PID 2>/dev/null || true
    echo "  已停止端口 $PORT 上的进程 PID $OLD_PID"
else
    echo "  无运行中的旧进程"
fi
sleep 1

# 3. 启动新进程
echo ""
echo "▶ 启动服务..."
nohup python3 "$APP" > nohup.out 2>&1 &
sleep 2

# 4. 查看启动日志
tail -20 nohup.out

echo "========================================"
