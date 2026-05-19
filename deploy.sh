#!/bin/bash
# deploy.sh — 拉取最新代码并重启服务

set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$PROJECT_DIR/venv"
APP="app.py"
LOG="$PROJECT_DIR/app.log"
BRANCH="dev"

echo "========================================"
echo "  Amazon Toolkit 部署脚本"
echo "  目录: $PROJECT_DIR"
echo "========================================"

# 1. 拉取最新代码
echo ""
echo "▶ 拉取代码 (origin/$BRANCH)..."
cd "$PROJECT_DIR"
git pull origin "$BRANCH"

# 2. 激活虚拟环境
echo ""
echo "▶ 激活虚拟环境..."
source "$VENV/bin/activate"

# 3. 更新依赖
echo ""
echo "▶ 更新依赖..."
pip install -r requirements.txt -q

# 4. 停止旧进程
echo ""
echo "▶ 停止旧服务..."
pkill -f "python3 $APP" 2>/dev/null && echo "  旧进程已停止" || echo "  无运行中的旧进程"
sleep 1

# 5. 启动新进程
echo ""
echo "▶ 启动服务..."
nohup python3 "$APP" > "$LOG" 2>&1 &
sleep 2

# 6. 检查是否成功启动
if pgrep -f "python3 $APP" > /dev/null; then
    echo ""
    echo "✅ 服务启动成功！"
    echo "   PID: $(pgrep -f "python3 $APP")"
    echo "   日志: tail -f $LOG"
else
    echo ""
    echo "❌ 服务启动失败，查看日志："
    tail -20 "$LOG"
    exit 1
fi

echo "========================================"
