#!/bin/bash

SERVICE_NAME="feishu-docget"
SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUN_SCRIPT="${PROJECT_ROOT}/run.sh"
CURRENT_USER="${SUDO_USER:-$USER}"

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

check_systemd() {
    if ! command -v systemctl &> /dev/null; then
        echo -e "${RED}错误: 未找到 systemctl 命令。${NC}"
        echo "此脚本专为使用 systemd 的 Linux 系统设计。"
        return 1
    fi
    return 0
}

is_installed() {
    if [ -f "$SERVICE_FILE" ]; then
        return 0
    else
        return 1
    fi
}

install_service() {
    echo -e "${BLUE}正在配置 systemd 服务...${NC}"
    
    if [ "$EUID" -ne 0 ]; then
        echo -e "${YELLOW}安装服务需要管理员权限。${NC}"
        if ! sudo -v; then
             echo -e "${RED}认证失败，操作已中止。${NC}"
             exit 1
        fi
        SUDO="sudo"
    else
        SUDO=""
    fi

    SERVICE_CONTENT="[Unit]
Description=Feishu DocGet Service
After=network.target

[Service]
Type=simple
User=$CURRENT_USER
WorkingDirectory=$PROJECT_ROOT
ExecStart=$RUN_SCRIPT
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target"

    echo "$SERVICE_CONTENT" | $SUDO tee "$SERVICE_FILE" > /dev/null
    
    if [ -f "$RUN_SCRIPT" ]; then
        $SUDO chmod +x "$RUN_SCRIPT"
        $SUDO sed -i 's/\r$//' "$RUN_SCRIPT" 2>/dev/null || true
    fi

    $SUDO systemctl daemon-reload
    $SUDO systemctl enable "$SERVICE_NAME"
    $SUDO systemctl start "$SERVICE_NAME"

    sleep 2
    if systemctl is-active --quiet "$SERVICE_NAME"; then
        echo -e "${GREEN}服务已成功安装并启动！${NC}"
    else
        echo -e "${RED}服务安装成功，但启动失败。${NC}"
        echo -e "${YELLOW}以下是最近的错误日志：${NC}"
        journalctl -u "$SERVICE_NAME" -n 20 --no-pager
    fi
}

get_status_text() {
    if systemctl is-active --quiet "$SERVICE_NAME"; then
        echo "running"
    else
        echo "stopped"
    fi
}

get_enable_status() {
    if systemctl is-enabled --quiet "$SERVICE_NAME"; then
        echo "enabled"
    else
        echo "disabled"
    fi
}

show_status() {
    echo -e "${BLUE}=== 服务状态 ===${NC}"
    systemctl status "$SERVICE_NAME" --no-pager
    echo -e "${BLUE}================${NC}"
}

check_config() {
    PROP_FILE="${PROJECT_ROOT}/feishu-docget.properties"
    
    TARGET_CONFIG="$PROP_FILE"

    if [ ! -f "$TARGET_CONFIG" ]; then
        echo -e "${YELLOW}配置文件未找到，正在初始化...${NC}"
        touch "$TARGET_CONFIG"
    fi

    APP_ID=$(grep "^feishu.app_id=" "$TARGET_CONFIG" | cut -d'=' -f2 | tr -d '\r')
    APP_SECRET=$(grep "^feishu.app_secret=" "$TARGET_CONFIG" | cut -d'=' -f2 | tr -d '\r')

    if [ -z "$APP_ID" ] || [ -z "$APP_SECRET" ]; then
        echo -e "${YELLOW}检测到缺少飞书配置，请根据提示输入。${NC}"
        echo -e "这些信息将被保存到 ${BLUE}$TARGET_CONFIG${NC} 中。"
        
        read -p "请输入飞书 App ID: " INPUT_APP_ID
        read -p "请输入飞书 App Secret: " INPUT_APP_SECRET
        
        
        if grep -q "^feishu.app_id=" "$TARGET_CONFIG"; then
            sed "s/^feishu.app_id=.*/feishu.app_id=${INPUT_APP_ID}/" "$TARGET_CONFIG" > "${TARGET_CONFIG}.tmp" && mv "${TARGET_CONFIG}.tmp" "$TARGET_CONFIG"
        else
            echo "feishu.app_id=${INPUT_APP_ID}" >> "$TARGET_CONFIG"
        fi

        if grep -q "^feishu.app_secret=" "$TARGET_CONFIG"; then
            sed "s/^feishu.app_secret=.*/feishu.app_secret=${INPUT_APP_SECRET}/" "$TARGET_CONFIG" > "${TARGET_CONFIG}.tmp" && mv "${TARGET_CONFIG}.tmp" "$TARGET_CONFIG"
        else
            echo "feishu.app_secret=${INPUT_APP_SECRET}" >> "$TARGET_CONFIG"
        fi
        
        echo -e "${GREEN}配置已更新。${NC}"
    fi
}


check_systemd || exit 1

check_config

if ! is_installed; then
    echo -e "${YELLOW}服务 '$SERVICE_NAME' 尚未在 systemd 中配置。${NC}"
    read -p "是否立即自动配置？(yes/no): " choice
    case "$choice" in 
        y|Y|yes|YES)
            install_service
            ;;
        *)
            echo "退出。"
            exit 0
            ;;
    esac
fi

while true; do
    CURRENT_STATUS=$(get_status_text)
    ENABLE_STATUS=$(get_enable_status)
    
    if [ "$CURRENT_STATUS" == "running" ]; then
        STATUS_DISPLAY="${GREEN}运行中${NC}"
    else
        STATUS_DISPLAY="${RED}已停止${NC}"
    fi

    if [ "$ENABLE_STATUS" == "enabled" ]; then
        ENABLE_DISPLAY="${GREEN}已开启${NC}"
    else
        ENABLE_DISPLAY="${RED}已关闭${NC}"
    fi
    
    echo
    echo -e "当前状态: $STATUS_DISPLAY  |  开机自启: $ENABLE_DISPLAY"
    echo "请选择操作 (输入数字):"
    
    PS3="> "
    options=("启动 (Start)" "停止 (Stop)" "重启 (Restart)" "开机自启 (Enable)" "取消自启 (Disable)" "状态 (Status)" "日志 (Logs)" "退出 (Exit)")
    select opt in "${options[@]}"
    do
        case $opt in
            "启动 (Start)")
                if [ "$CURRENT_STATUS" == "running" ]; then
                    echo -e "${YELLOW}服务已经在运行中。${NC}"
                else
                    if [ -f "$RUN_SCRIPT" ] && [ ! -x "$RUN_SCRIPT" ]; then
                        echo -e "${YELLOW}正在修复脚本权限...${NC}"
                        sudo chmod +x "$RUN_SCRIPT"
                        sudo sed -i 's/\r$//' "$RUN_SCRIPT" 2>/dev/null || true
                    fi

                    sudo systemctl start "$SERVICE_NAME"
                    echo -e "${BLUE}正在启动...${NC}"
                    sleep 2
                    if systemctl is-active --quiet "$SERVICE_NAME"; then
                        echo -e "${GREEN}服务已成功启动。${NC}"
                    else
                        echo -e "${RED}启动失败。${NC}"
                        echo -e "${YELLOW}查看最后 20 行日志：${NC}"
                        journalctl -u "$SERVICE_NAME" -n 20 --no-pager
                    fi
                fi
                break
                ;;
            "停止 (Stop)")
                if [ "$CURRENT_STATUS" == "stopped" ]; then
                    echo -e "${YELLOW}服务已经是停止状态。${NC}"
                else
                    sudo systemctl stop "$SERVICE_NAME"
                    echo -e "${RED}服务已停止。${NC}"
                fi
                break
                ;;
            "重启 (Restart)")
                sudo systemctl restart "$SERVICE_NAME"
                echo -e "${GREEN}服务已重启。${NC}"
                break
                ;;
            "开机自启 (Enable)")
                if [ "$ENABLE_STATUS" == "enabled" ]; then
                    echo -e "${YELLOW}开机自启已经是开启状态。${NC}"
                else
                    sudo systemctl enable "$SERVICE_NAME"
                    echo -e "${GREEN}开机自启已开启。${NC}"
                fi
                break
                ;;
            "取消自启 (Disable)")
                if [ "$ENABLE_STATUS" == "disabled" ]; then
                    echo -e "${YELLOW}开机自启已经是关闭状态。${NC}"
                else
                    sudo systemctl disable "$SERVICE_NAME"
                    echo -e "${RED}开机自启已关闭。${NC}"
                fi
                break
                ;;
            "状态 (Status)")
                show_status
                break
                ;;
            "日志 (Logs)")
                echo -e "${BLUE}正在显示日志 (按 Ctrl+C 退出)...${NC}"
                journalctl -u "$SERVICE_NAME" -f
                break
                ;;
            "退出 (Exit)")
                echo "再见。"
                exit 0
                ;;
            *) 
                echo -e "${YELLOW}无效选项 $REPLY${NC}"
                ;;
        esac
    done
done
