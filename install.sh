#!/bin/bash

# 预处理工具安装脚本
# 支持在Mac系统上一键安装

set -e  # 遇到错误时退出

echo "🚀 开始安装预处理工具..."

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误：未找到Python3，请先安装Python3"
    exit 1
fi

echo "✅ Python3已安装"

# 检查pip是否安装
if ! command -v pip3 &> /dev/null; then
    echo "❌ 错误：未找到pip3，请先安装pip3"
    exit 1
fi

echo "✅ pip3已安装"

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo "📁 安装目录：$SCRIPT_DIR"

# 安装Python依赖
echo "📦 安装Python依赖包..."

# 尝试多种安装方式
if pip3 install --user -r "$SCRIPT_DIR/requirements.txt" 2>/dev/null; then
    echo "✅ Python依赖安装成功（用户级安装）"
elif pip3 install --break-system-packages -r "$SCRIPT_DIR/requirements.txt" 2>/dev/null; then
    echo "✅ Python依赖安装成功（系统级安装）"
else
    echo "⚠️  标准安装失败，尝试创建虚拟环境..."
    
    # 创建虚拟环境
    python3 -m venv "$SCRIPT_DIR/venv"
    source "$SCRIPT_DIR/venv/bin/activate"
    
    if pip install -r "$SCRIPT_DIR/requirements.txt"; then
        echo "✅ Python依赖安装成功（虚拟环境）"
        
        # 修改启动脚本以使用虚拟环境
        cat > "$SCRIPT_DIR/yuchuli_launcher.sh" << EOF
#!/bin/bash
cd "$SCRIPT_DIR"
source "$SCRIPT_DIR/venv/bin/activate"
python3 "$SCRIPT_DIR/yuchuli.py" "\$@"
EOF
    else
        echo "❌ Python依赖安装失败"
        echo "请手动安装依赖：pip3 install --user pandas openpyxl colorama xlrd"
        exit 1
    fi
fi

# 创建可执行文件
INSTALL_DIR="/usr/local/bin"
COMMAND_NAME="yuchuli"
USER_BIN_DIR="$HOME/.local/bin"

echo "🔧 创建命令行工具..."

# 检查是否使用了虚拟环境
if [ ! -f "$SCRIPT_DIR/yuchuli_launcher.sh" ]; then
    # 创建启动脚本（非虚拟环境）
    cat > "$SCRIPT_DIR/yuchuli_launcher.sh" << EOF
#!/bin/bash
cd "$SCRIPT_DIR"
python3 "$SCRIPT_DIR/yuchuli.py" "\$@"
EOF
fi

# 设置执行权限
chmod +x "$SCRIPT_DIR/yuchuli_launcher.sh"
chmod +x "$SCRIPT_DIR/yuchuli.py"

# 尝试安装到系统目录
echo "🔐 尝试安装到系统目录..."
if sudo ln -sf "$SCRIPT_DIR/yuchuli_launcher.sh" "$INSTALL_DIR/$COMMAND_NAME" 2>/dev/null; then
    echo "✅ 命令行工具已安装到 $INSTALL_DIR/$COMMAND_NAME"
else
    echo "⚠️  无法安装到系统目录，将安装到用户目录..."
    
    # 创建用户本地bin目录
    mkdir -p "$USER_BIN_DIR"
    ln -sf "$SCRIPT_DIR/yuchuli_launcher.sh" "$USER_BIN_DIR/$COMMAND_NAME"
    
    echo "✅ 命令行工具已安装到 $USER_BIN_DIR/$COMMAND_NAME"
    
    # 检查并更新PATH环境变量
    if [[ ":$PATH:" != *":$USER_BIN_DIR:"* ]]; then
        echo "📝 正在更新PATH环境变量..."
        
        # 检查shell类型并更新相应的配置文件
        if [[ "$SHELL" == *"zsh"* ]]; then
            # 对于zsh，更新.zshrc
            if ! grep -q "export PATH=\"\$HOME/.local/bin:\$PATH\"" ~/.zshrc 2>/dev/null; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc
                echo "✅ 已更新 ~/.zshrc 文件"
            fi
        elif [[ "$SHELL" == *"bash"* ]]; then
            # 对于bash，更新.bashrc
            if ! grep -q "export PATH=\"\$HOME/.local/bin:\$PATH\"" ~/.bashrc 2>/dev/null; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
                echo "✅ 已更新 ~/.bashrc 文件"
            fi
        else
            # 默认更新.zshrc
            if ! grep -q "export PATH=\"\$HOME/.local/bin:\$PATH\"" ~/.zshrc 2>/dev/null; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc
                echo "✅ 已更新 ~/.zshrc 文件"
            fi
        fi
        
        echo "📝 注意：请重新打开终端或运行以下命令来应用PATH更改："
        echo "   source ~/.zshrc  # 如果使用zsh"
        echo "   source ~/.bashrc # 如果使用bash"
    fi
fi

echo ""
echo "🎉 安装完成！"
echo ""
echo "使用方法："
echo "  在终端中输入 '$COMMAND_NAME' 即可启动程序"
echo ""
echo "功能说明："
echo "  • 处理Excel表格数据"
echo "  • 提取商品编码和采购价"
echo "  • 自动计算成本价和基本售价"
echo "  • 生成标准格式的新Excel文件"
echo ""
echo "如需卸载，请运行："
echo "  sudo rm $INSTALL_DIR/$COMMAND_NAME  # 如果安装在系统目录"
echo "  或"
echo "  rm $USER_BIN_DIR/$COMMAND_NAME      # 如果安装在用户目录"
echo ""

# 验证安装
if command -v $COMMAND_NAME &> /dev/null; then
    echo "✅ 安装验证成功！现在可以使用 '$COMMAND_NAME' 命令"
else
    echo "⚠️  安装验证失败，请检查PATH环境变量"
    echo "   您可以手动运行以下命令来启动程序："
    echo "   $SCRIPT_DIR/yuchuli_launcher.sh"
fi