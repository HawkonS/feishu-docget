#!/bin/bash

# 获取脚本所在目录的上一级目录（项目根目录）
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR" || exit 1

echo "=========================================="
echo "   正在更新 feishu-docget..."
echo "=========================================="
echo "项目目录: $PROJECT_DIR"

# 1. 强制拉取远程代码
echo "[1/2] 正在拉取远程代码..."
git fetch --all
git reset --hard origin/main
if [ $? -ne 0 ]; then
    echo "❌ 代码更新失败，请检查网络或 Git 配置"
    exit 1
fi
echo "✅ 代码已更新到最新版本"

# 2. 尝试重启服务
echo "[2/2] 正在尝试重启服务..."

CONFIG_FILE="$PROJECT_DIR/feishu-docget.properties"
SUDO_PASS=""

if [ -f "$CONFIG_FILE" ]; then
    SUDO_PASS=$(grep "^system.sudo_password=" "$CONFIG_FILE" | cut -d'=' -f2 | tr -d '\r')
fi

if [ -z "$SUDO_PASS" ]; then
    echo "⚠️  注意：重启服务可能需要 sudo 权限"
    read -s -p "请输入当前用户的 sudo 密码 (留空则直接尝试): " USER_INPUT_PASS
    echo ""
    if [ -n "$USER_INPUT_PASS" ]; then
        SUDO_PASS="$USER_INPUT_PASS"
    fi
fi

SUDO_CMD="sudo"
if [ -n "$SUDO_PASS" ]; then
    SUDO_CMD="echo \"$SUDO_PASS\" | sudo -S"
fi

# 检查是否存在 systemd 服务
if systemctl status feishu-docget >/dev/null 2>&1; then
    echo "检测到 systemd 服务: feishu-docget"
    
    # 尝试重启
    echo "正在尝试重启..."
    eval "$SUDO_CMD systemctl restart feishu-docget"
    
    if [ $? -eq 0 ]; then
        echo "✅ 服务重启成功！"
        echo "------------------------------------------"
        systemctl status feishu-docget --no-pager
        echo "------------------------------------------"
    else
        echo "❌ 服务重启失败，请检查密码是否正确，或手动执行：sudo systemctl restart feishu-docget"
    fi
else
    echo "⚠️  未检测到 feishu-docget 系统服务，或者当前用户无权访问。"
    echo "✅ 代码已更新完成！"
    echo "👉 请手动重启您的程序 (例如: ./run.sh 或 kill 掉旧进程后重新启动)"
fi

echo "=========================================="
echo "   更新流程结束"
echo "=========================================="
