#!/bin/bash

# ==========================================
# Feishu DocGet Service Stopper
# ==========================================

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo "正在查找 feishu_docget 服务进程..."

# 查找运行 src/app.py 的 Python 进程（-i 忽略大小写，macOS 上 Python 路径可能含大写）
PIDS=$(pgrep -if "src/app.py" 2>/dev/null)

if [ -z "$PIDS" ]; then
    echo "未找到正在运行的 feishu_docget 服务。"
    exit 0
fi

echo "找到服务进程 (PID: $PIDS)，正在停止..."

for PID in $PIDS; do
    # 确认进程工作目录属于本项目
    PROC_CWD=$(lsof -p "$PID" -Fn 2>/dev/null | grep "^n/" | head -1 | sed 's/^n//')
    if [ -z "$PROC_CWD" ]; then
        PROC_CWD=$(readlink "/proc/$PID/cwd" 2>/dev/null)
    fi

    kill "$PID" 2>/dev/null
    if [ $? -eq 0 ]; then
        echo "已发送终止信号到进程 $PID"
    else
        echo "无法终止进程 $PID，尝试强制终止..."
        kill -9 "$PID" 2>/dev/null
    fi
done

# 等待进程退出
sleep 2

# 检查是否还有残留进程
REMAINING=$(pgrep -if "src/app.py" 2>/dev/null)
if [ -n "$REMAINING" ]; then
    echo "部分进程未响应，正在强制终止..."
    for PID in $REMAINING; do
        kill -9 "$PID" 2>/dev/null
    done
fi

echo "feishu_docget 服务已停止。"
