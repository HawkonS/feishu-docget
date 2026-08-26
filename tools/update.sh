#!/bin/bash

# 获取脚本所在目录的上一级目录（项目根目录）
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR" || exit 1

# 解析参数
SKIP_CONFIRM=false
for arg in "$@"; do
    case "$arg" in
        --yes|-y) SKIP_CONFIRM=true ;;
    esac
done

CONFIG_FILE="$PROJECT_DIR/feishu-docget.properties"

# 解析日志目录（从配置文件读取 log.dir，解析失败回退 logs）
LOG_DIR="logs"
if [ -f "$CONFIG_FILE" ]; then
    CFG_LOG_DIR=$(grep "^log.dir=" "$CONFIG_FILE" | cut -d'=' -f2 | tr -d '\r')
    if [ -n "$CFG_LOG_DIR" ]; then
        LOG_DIR="$CFG_LOG_DIR"
    fi
fi
case "$LOG_DIR" in
    /*) UPDATE_LOG="$LOG_DIR/update.log" ;;
    *)  UPDATE_LOG="$PROJECT_DIR/$LOG_DIR/update.log" ;;
esac

# 截断上次升级日志，仅保留最近一次升级的输出；路径无效时静默跳过
mkdir -p "$(dirname "$UPDATE_LOG")" 2>/dev/null
: > "$UPDATE_LOG" 2>/dev/null || true

echo "=========================================="
echo "   正在更新 feishu-docget..."
echo "=========================================="
echo "项目目录: $PROJECT_DIR"

# 1. 获取远程代码，并在失败时停止整个流程。
#    不能继续使用本地残留的 origin/main，否则网络失败时也会误应用旧版本并重启服务。
REMOTE_NAME="origin"
REMOTE_BRANCH="main"

CURRENT_COMMIT=$(git rev-parse --verify HEAD 2>/dev/null)
if [ -z "$CURRENT_COMMIT" ]; then
    echo "❌ 无法读取当前 Git 版本，已停止升级；服务未重启"
    exit 1
fi
CURRENT_VERSION=$(git show -s --format='%H %ad %s' --date=iso "$CURRENT_COMMIT" 2>/dev/null) || CURRENT_VERSION="$CURRENT_COMMIT"
echo "当前版本: $CURRENT_VERSION"

echo "[1/2] 正在拉取远程代码 ($REMOTE_NAME/$REMOTE_BRANCH)..."
if ! git fetch --prune "$REMOTE_NAME" "$REMOTE_BRANCH"; then
    echo "❌ 远程代码拉取失败，已停止升级；服务未重启"
    echo "当前版本保持不变: $CURRENT_VERSION"
    exit 1
fi

TARGET_COMMIT=$(git rev-parse --verify "refs/remotes/$REMOTE_NAME/$REMOTE_BRANCH^{commit}" 2>/dev/null)
if [ -z "$TARGET_COMMIT" ]; then
    echo "❌ 拉取成功但未找到远程目标版本 $REMOTE_NAME/$REMOTE_BRANCH，已停止升级；服务未重启"
    echo "当前版本保持不变: $CURRENT_VERSION"
    exit 1
fi
TARGET_VERSION=$(git show -s --format='%H %ad %s' --date=iso "$TARGET_COMMIT" 2>/dev/null) || TARGET_VERSION="$TARGET_COMMIT"
echo "目标版本: $TARGET_VERSION"

echo "=== 即将应用以下更新 ==="
git log "HEAD..$TARGET_COMMIT" --oneline 2>/dev/null || echo "(无法获取更新日志)"
echo ""
if [ "$SKIP_CONFIRM" != "true" ]; then
    echo "确认更新？(y/N)"
    read -r confirm
    if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
        echo "更新已取消"
        exit 0
    fi
fi

# 使用本次 fetch 确认过的提交哈希，避免 reset 到不明确或已变化的远程引用。
if ! git reset --hard "$TARGET_COMMIT"; then
    echo "❌ 代码更新失败，请检查网络或 Git 配置"
    echo "当前版本保持不变: $CURRENT_VERSION"
    exit 1
fi

APPLIED_COMMIT=$(git rev-parse --verify HEAD 2>/dev/null)
if [ "$APPLIED_COMMIT" != "$TARGET_COMMIT" ]; then
    echo "❌ 应用后的版本校验失败，已停止重启"
    echo "期望版本: $TARGET_VERSION"
    ACTUAL_VERSION=$(git show -s --format='%H %ad %s' --date=iso "$APPLIED_COMMIT" 2>/dev/null) || ACTUAL_VERSION="$APPLIED_COMMIT"
    echo "实际版本: $ACTUAL_VERSION"
    exit 1
fi
APPLIED_VERSION=$(git show -s --format='%H %ad %s' --date=iso "$APPLIED_COMMIT" 2>/dev/null) || APPLIED_VERSION="$APPLIED_COMMIT"
echo "✅ 代码已更新到版本: $APPLIED_VERSION"

# 2. 尝试重启服务
echo "[2/2] 正在尝试重启服务..."

SUDO_PASS=""

if [ -f "$CONFIG_FILE" ]; then
    SUDO_PASS=$(grep "^system.sudo_password=" "$CONFIG_FILE" | cut -d'=' -f2 | tr -d '\r')
fi

# --yes（非交互，如管理后台后台启动）时跳过 sudo 密码询问：无 systemctl 的环境不需要 sudo
if [ -z "$SUDO_PASS" ] && [ "$SKIP_CONFIRM" != "true" ]; then
    echo "⚠️  注意：重启服务可能需要 sudo 权限"
    read -s -p "请输入当前用户的 sudo 密码 (留空则直接尝试): " USER_INPUT_PASS
    echo ""
    if [ -n "$USER_INPUT_PASS" ]; then
        SUDO_PASS="$USER_INPUT_PASS"
    fi
fi

run_sudo() {
    if [ -n "$SUDO_PASS" ]; then
        echo "$SUDO_PASS" | sudo -S "$@"
    else
        sudo "$@"
    fi
}

# 检查 systemd 服务单元是否存在（用 list-unit-files 区分“unit 不存在”与“unit 存在但未运行”；无 systemctl 环境兼容）
UNIT_EXISTS=false
if command -v systemctl >/dev/null 2>&1; then
    if systemctl list-unit-files feishu-docget.service 2>/dev/null | grep -q "^feishu-docget\.service"; then
        UNIT_EXISTS=true
    fi
fi

if [ "$UNIT_EXISTS" = true ]; then
    echo "检测到 systemd 服务: feishu-docget"
    
    # 尝试重启
    echo "正在尝试重启..."
    run_sudo systemctl restart feishu-docget
    
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
    echo "尝试通过 stop.sh / run.sh 重启服务..."

    # 优先调用项目根目录的 stop.sh 停止旧进程
    if [ -f "$PROJECT_DIR/stop.sh" ]; then
        bash "$PROJECT_DIR/stop.sh"
    else
        echo "⚠️  未找到 stop.sh，跳过停止步骤"
    fi
    sleep 1

    # 复核旧进程已消失，避免新旧两个版本同时运行占用同一端口
    REMAINING=$(pgrep -if "src/app.py" 2>/dev/null)
    if [ -n "$REMAINING" ]; then
        echo "❌ 旧服务进程仍存活 (PID: $REMAINING)，无法安全启动新版本，请手动停止后重新执行升级"
        exit 1
    fi

    # 应用日志已由自身 logging 文件 Handler 完整记录，run.sh 的 stdout/stderr 重定向到 /dev/null，
    # 避免常驻服务输出导致 update.log 无界增长
    # stdin 重定向到 /dev/null，确保 run.sh 走非交互分支
    if [ -f "$PROJECT_DIR/run.sh" ]; then
        nohup bash "$PROJECT_DIR/run.sh" > /dev/null 2>&1 < /dev/null &
        echo "✅ 代码已更新完成，服务已在后台重启（升级日志见 $LOG_DIR/update.log）"
    else
        echo "❌ 未找到 run.sh，请手动重启您的程序"
    fi
fi

echo "=========================================="
echo "   更新流程结束"
echo "=========================================="
