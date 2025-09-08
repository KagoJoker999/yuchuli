# 预处理工具

一个用于处理Excel表格数据的命令行工具，可以提取商品编码和采购价信息，自动计算成本价和基本售价，并生成标准格式的新Excel文件。

## 功能特点

- 🖱️ 支持拖拽导入Excel文件
- 🎯 自动提取商品编码（M列）和采购价（S列）数据
- 🧮 自动计算成本价和基本售价
- 📊 生成标准格式的Excel文件
- 🎨 美观的交互式菜单界面
- ⌨️ 支持键盘导航（上下键选择，回车确认）
- 🚀 一键安装，安装后可直接使用 `yuchuli` 命令

## 安装方法

### 方法一：快速安装（推荐）

打开终端，运行以下命令即可一键安装：

```bash
curl -sSL https://raw.githubusercontent.com/KagoJoker999/yuchuli/master/quick_install.sh | bash
```

或者使用wget：

```bash
wget -qO- https://raw.githubusercontent.com/KagoJoker999/yuchuli/master/quick_install.sh | bash
```

### 方法二：手动安装

1. 克隆仓库：
   ```bash
   git clone https://github.com/KagoJoker999/yuchuli.git
   cd yuchuli
   ```

2. 运行安装脚本：
   ```bash
   chmod +x install.sh
   ./install.sh
   ```

### 方法三：手动安装依赖

如果自动安装失败，可以手动安装依赖：

```bash
pip3 install pandas openpyxl colorama xlrd
python3 yuchuli.py
```

## 使用说明

1. 启动程序：
   ```bash
   yuchuli
   ```

2. 使用键盘导航：
   - ↑↓ 键：选择菜单项
   - 回车键：确认选择
   - 0 键：退出程序（也可在菜单中选择"退出程序(按0键)"选项）

3. 选择"处理Excel文件"，然后拖拽Excel文件到终端窗口

4. 程序会自动处理并生成结果文件

## 数据处理规则

### 输入要求
- **商品编码**：从M列提取（或包含"商品编码"的列）
- **采购价**：从S列提取（或包含"采购价"的列）

### 计算规则
- **成本价** = 采购价 + 2
- **基本售价** = ((采购价 + 5) × 2)的整数部分 + 0.99
- **虚拟分类** = "可预售"（固定值）

### 输出格式
生成的Excel文件包含以下列：
- 商品编码
- 采购价
- 成本价
- 基本售价
- 虚拟分类

## 支持的文件格式

- **输入**：.xlsx、.xls格式的Excel文件
- **输出**：.xlsx格式的Excel文件

## 系统要求

- macOS 系统
- Python 3.6 或更高版本
- pip3 包管理器

## 常见问题解决

### 问题：zsh: command not found: yuchuli

这是最常见的问题，通常是由于PATH环境变量未正确配置导致的。请按以下步骤解决：

1. **重新打开终端**：
   关闭当前终端窗口，重新打开一个新的终端窗口，然后再次尝试运行[yuchuli](file:///Users/majingkai/Downloads/Demo/待上架预处理/yuchuli.py#L0-L285)命令。

2. **手动激活PATH**：
   如果重新打开终端无效，请运行以下命令：
   ```bash
   # 对于zsh用户（macOS默认）
   source ~/.zshrc
   
   # 对于bash用户
   source ~/.bashrc
   ```

3. **检查安装位置**：
   运行以下命令检查[yuchuli](file:///Users/majingkai/Downloads/Demo/待上架预处理/yuchuli.py#L0-L285)命令是否已正确安装：
   ```bash
   # 检查系统目录
   ls -la /usr/local/bin/yuchuli
   
   # 检查用户目录
   ls -la ~/.local/bin/yuchuli
   ```

4. **手动添加到PATH**：
   如果[yuchuli](file:///Users/majingkai/Downloads/Demo/待上架预处理/yuchuli.py#L0-L285)命令安装在`~/.local/bin`目录，请手动将该目录添加到PATH：
   ```bash
   # 临时添加（仅当前会话有效）
   export PATH="$HOME/.local/bin:$PATH"
   
   # 永久添加（需要根据您的shell选择）
   echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc  # 对于zsh
   echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc # 对于bash
   ```

5. **直接运行程序**：
   如果以上方法都无效，您可以直接运行程序：
   ```bash
   # 进入安装目录后运行
   ~/.local/bin/yuchuli
   # 或者
   /usr/local/bin/yuchuli
   ```

### 问题：权限不足

**解决方法**：使用sudo权限运行安装脚本，或安装到用户目录

### 问题：Python依赖安装失败

**解决方法**：手动安装依赖包：
```bash
pip3 install pandas openpyxl colorama xlrd
```

### 问题：无法读取Excel文件

**解决方法**：
1. 确保文件格式为.xlsx或.xls
2. 确保文件没有被其他程序占用
3. 检查文件路径是否正确

## 卸载方法

```bash
# 如果安装在系统目录
sudo rm /usr/local/bin/yuchuli

# 如果安装在用户目录
rm ~/.local/bin/yuchuli
```

## 技术支持

如遇到问题，请检查：
1. Python版本是否符合要求
2. 依赖包是否正确安装
3. Excel文件格式是否正确
4. 文件路径是否正确

## 更新日志

### v1.0.0
- 初始版本发布
- 支持Excel数据处理
- 实现交互式菜单
- 添加一键安装功能