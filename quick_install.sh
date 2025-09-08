#!/bin/bash

# 预处理工具快速安装脚本
# 通过一行命令即可安装预处理工具

echo "🚀 开始安装预处理工具..."

# 检查系统是否有git
if ! command -v git &> /dev/null; then
    echo "❌ 错误：未找到git，请先安装git"
    exit 1
fi

# 创建临时目录
TEMP_DIR=$(mktemp -d)
cd "$TEMP_DIR"

# 克隆仓库
echo "📥 正在下载预处理工具..."
git clone https://github.com/KagoJoker999/yuchuli.git
cd yuchuli

# 运行安装脚本
echo "⚙️  正在安装依赖和配置..."
chmod +x install.sh
./install.sh

# 检查安装是否成功
if ! command -v yuchuli &> /dev/null; then
    echo "⚠️  检测到命令未找到，正在尝试修复PATH..."
    
    # 检查用户目录中的安装
    if [ -f "$HOME/.local/bin/yuchuli" ]; then
        echo "📝 检测到安装在用户目录，正在更新PATH..."
        
        # 检查shell类型
        if [[ "$SHELL" == *"zsh"* ]]; then
            # 对于zsh，更新.zshrc
            if ! grep -q 'export PATH="$HOME/.local/bin:$PATH"' ~/.zshrc; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc
                echo "✅ 已更新 ~/.zshrc 文件"
            fi
            source ~/.zshrc
        elif [[ "$SHELL" == *"bash"* ]]; then
            # 对于bash，更新.bashrc
            if ! grep -q 'export PATH="$HOME/.local/bin:$PATH"' ~/.bashrc; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
                echo "✅ 已更新 ~/.bashrc 文件"
            fi
            source ~/.bashrc
        fi
        
        echo "🔄 请重新打开终端或运行以下命令来应用更改："
        echo "   source ~/.zshrc  # 如果使用zsh"
        echo "   source ~/.bashrc # 如果使用bash"
        echo ""
        echo "然后就可以使用 'yuchuli' 命令了"
    fi
fi

# 清理临时文件
cd ~
rm -rf "$TEMP_DIR"

echo ""
echo "🎉 预处理工具安装完成！"
echo ""
echo "使用方法："
echo "  在终端中输入 'yuchuli' 即可启动程序"
echo ""
echo "如遇到命令未找到问题，请重新打开终端或手动执行："
echo "  source ~/.zshrc  # 如果使用zsh"
echo "  source ~/.bashrc # 如果使用bash"