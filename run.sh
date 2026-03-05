#!/bin/bash

# ==========================================
# Feishu DocGet Service Runner
# ==========================================

# 1. 检查是否为交互式终端（用于首次设置）
if [ -t 0 ]; then
    INTERACTIVE=true
else
    INTERACTIVE=false
fi

check_command() {
    command -v "$1" >/dev/null 2>&1
}

install_pkg_linux() {
    PKG=$1
    if check_command apt-get; then
        sudo apt-get update && sudo apt-get install -y "$PKG"
    elif check_command yum; then
        sudo yum install -y "$PKG"
    elif check_command dnf; then
        sudo dnf install -y "$PKG"
    else
        echo "错误: 不支持的包管理器，请手动安装 $PKG。"
        exit 1
    fi
}

install_pkg_mac() {
    PKG=$1
    if check_command brew; then
        brew install "$PKG"
    else
        echo "错误: 未找到 Homebrew，请先安装 Homebrew。"
        exit 1
    fi
}

# 2. 检查并安装 Python3 和 pip3
if ! check_command python3; then
    echo "未找到 Python3。"
    if [ "$INTERACTIVE" = true ]; then
        read -p "是否安装 Python3? (y/n) " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            if [[ "$OSTYPE" == "darwin"* ]]; then
                install_pkg_mac python3
            else
                install_pkg_linux python3
                install_pkg_linux python3-pip
            fi
        else
            echo "已取消。"
            exit 1
        fi
    else
        echo "非交互模式: 缺少 Python3，退出。"
        exit 1
    fi
fi

# 确保 pip3 已安装
if ! check_command pip3; then
    echo "未找到 pip3。"
    if [ "$INTERACTIVE" = true ]; then
        read -p "是否安装 pip3? (y/n) " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            if [[ "$OSTYPE" == "darwin"* ]]; then
                 # Mac 上 python3 通常自带 pip3
                 echo "正在重新安装 python3 以修复 pip..."
                 install_pkg_mac python3
            else
                install_pkg_linux python3-pip
            fi
        else
            echo "已取消。"
            exit 1
        fi
    else
        echo "非交互模式: 缺少 pip3。尝试自动安装..."
        if [[ "$OSTYPE" == "linux-gnu"* ]]; then
             install_pkg_linux python3-pip
        else
             echo "无法自动安装 pip3，退出。"
             exit 1
        fi
    fi
fi

# 3. 安装 Python 依赖
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR" || exit

# 依赖包列表
REQUIRED_PACKAGES="Flask requests python-docx lxml Pillow"

echo "检查 Python 依赖: $REQUIRED_PACKAGES"

install_packages() {
    # 尝试使用 --break-system-packages 安装（适配新版 Python/OS）
    if pip3 install $REQUIRED_PACKAGES --break-system-packages; then
        return 0
    elif pip3 install $REQUIRED_PACKAGES; then
        return 0
    else
        echo "权限不足或安装失败。尝试使用 sudo..."
        if sudo pip3 install $REQUIRED_PACKAGES --break-system-packages; then
            return 0
        elif sudo pip3 install $REQUIRED_PACKAGES; then
            return 0
        else
            return 1
        fi
    fi
}

if [ "$INTERACTIVE" = true ]; then
    install_packages || {
        echo "安装 Python 依赖失败。请检查网络连接或权限。"
        exit 1
    }
else
    # 非交互模式下，尝试安装并忽略部分输出
    pip3 install $REQUIRED_PACKAGES --break-system-packages >/dev/null 2>&1 || pip3 install $REQUIRED_PACKAGES >/dev/null 2>&1
fi

# 4. 启动应用
echo "正在启动 feishu_docget 服务..."
export PYTHONPATH=$PYTHONPATH:.

# 检查并更新配置文件 (补充缺失的配置项)
python3 -c "from src.core.config_loader import ConfigLoader; ConfigLoader.load_config()"

exec python3 src/app.py