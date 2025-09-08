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

# 清理临时文件
cd ~
rm -rf "$TEMP_DIR"

echo ""
echo "🎉 预处理工具安装完成！"
echo ""
echo "使用方法："
echo "  在终端中输入 'yuchuli' 即可启动程序"
echo ""
echo "如需查看帮助信息，请运行："
echo "  yuchuli --help"