#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
预处理工具 - Excel数据处理程序
支持拖拽导入Excel文件，提取商品编码和采购价数据，生成新的格式化表格
"""

import os
import sys
import pandas as pd
import math
from pathlib import Path
import getch
from colorama import init, Fore, Style, Back

# 初始化colorama
init()

class ExcelProcessor:
    def __init__(self):
        self.input_file = None
        self.output_file = None
        
    def display_banner(self):
        """显示程序横幅"""
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}             预处理工具 - Excel数据处理")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")
    
    def get_file_input(self):
        """获取用户输入的文件路径"""
        print(f"{Fore.YELLOW}请拖拽Excel文件到终端窗口，或输入文件路径：{Style.RESET_ALL}")
        file_path = input().strip()
        
        # 清理路径（移除可能的引号）
        file_path = file_path.strip('"\'')
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"{Fore.RED}错误：文件不存在！{Style.RESET_ALL}")
            return False
            
        # 检查文件扩展名
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            print(f"{Fore.RED}错误：请选择Excel文件（.xlsx或.xls格式）！{Style.RESET_ALL}")
            return False
            
        self.input_file = file_path
        return True
    
    def process_excel(self):
        """处理Excel文件"""
        try:
            print(f"\n{Fore.BLUE}正在读取Excel文件...{Style.RESET_ALL}")
            
            # 读取Excel文件
            df = pd.read_excel(self.input_file)
            
            # 打印列名以便调试
            print(f"{Fore.CYAN}表格列名：{list(df.columns)}{Style.RESET_ALL}")
            
            # 查找"商品编码"列（通常在M列，但我们通过列名查找更可靠）
            product_code_col = None
            purchase_price_col = None
            
            # 查找商品编码列
            for col in df.columns:
                if '商品编码' in str(col) or 'product_code' in str(col).lower():
                    product_code_col = col
                    break
            
            # 如果没找到，尝试M列（第13列，索引12）
            if product_code_col is None and len(df.columns) > 12:
                product_code_col = df.columns[12]  # M列
                print(f"{Fore.YELLOW}未找到'商品编码'列名，使用M列：{product_code_col}{Style.RESET_ALL}")
            
            # 查找采购价列
            for col in df.columns:
                if '采购价' in str(col) or 'purchase_price' in str(col).lower():
                    purchase_price_col = col
                    break
            
            # 如果没找到，尝试S列（第19列，索引18）
            if purchase_price_col is None and len(df.columns) > 18:
                purchase_price_col = df.columns[18]  # S列
                print(f"{Fore.YELLOW}未找到'采购价'列名，使用S列：{purchase_price_col}{Style.RESET_ALL}")
            
            if product_code_col is None or purchase_price_col is None:
                print(f"{Fore.RED}错误：无法找到商品编码或采购价列！{Style.RESET_ALL}")
                return False
            
            print(f"{Fore.GREEN}找到商品编码列：{product_code_col}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}找到采购价列：{purchase_price_col}{Style.RESET_ALL}")
            
            # 提取数据
            product_codes = df[product_code_col].dropna()
            purchase_prices = df[purchase_price_col].dropna()
            
            # 确保数据长度一致
            min_length = min(len(product_codes), len(purchase_prices))
            product_codes = product_codes.iloc[:min_length]
            purchase_prices = purchase_prices.iloc[:min_length]
            
            print(f"{Fore.BLUE}提取到 {min_length} 条有效数据{Style.RESET_ALL}")
            
            # 创建新的数据框
            new_data = []
            
            for i in range(min_length):
                product_code = product_codes.iloc[i]
                purchase_price = purchase_prices.iloc[i]
                
                # 跳过无效数据
                if pd.isna(product_code) or pd.isna(purchase_price):
                    continue
                
                try:
                    purchase_price = float(purchase_price)
                except (ValueError, TypeError):
                    continue
                
                # 计算成本价：采购价 + 2
                cost_price = purchase_price + 2
                
                # 计算基本售价：((采购价 + 5) * 2)的整数结果 + 0.99
                basic_selling_price = math.floor((purchase_price + 5) * 2) + 0.99
                
                # 虚拟分类固定为"可预售"
                virtual_category = "可预售"
                
                new_data.append({
                    '商品编码': product_code,
                    '采购价': purchase_price,
                    '成本价': cost_price,
                    '基本售价': basic_selling_price,
                    '虚拟分类': virtual_category
                })
            
            # 创建新的DataFrame
            new_df = pd.DataFrame(new_data)
            
            # 生成输出文件名
            if self.input_file:
                input_path = Path(self.input_file)
                output_filename = f"{input_path.stem}_预处理结果.xlsx"
                self.output_file = input_path.parent / output_filename
            else:
                return False
            
            # 保存新的Excel文件
            print(f"\n{Fore.BLUE}正在保存处理结果...{Style.RESET_ALL}")
            new_df.to_excel(self.output_file, index=False)
            
            print(f"\n{Fore.GREEN}✅ 处理完成！{Style.RESET_ALL}")
            print(f"{Fore.GREEN}📁 输出文件：{self.output_file}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}📊 处理数据条数：{len(new_data)}{Style.RESET_ALL}")
            
            return True
            
        except Exception as e:
            print(f"{Fore.RED}处理文件时发生错误：{str(e)}{Style.RESET_ALL}")
            return False

class InteractiveMenu:
    def __init__(self):
        self.options = [
            "处理Excel文件",
            "查看帮助",
            "退出程序(按0键)"
        ]
        self.selected = 0
        self.processor = ExcelProcessor()
    
    def display_menu(self):
        """显示菜单"""
        os.system('clear' if os.name == 'posix' else 'cls')
        self.processor.display_banner()
        
        print(f"{Fore.YELLOW}使用 ↑↓ 键选择，回车确认，0 键退出{Style.RESET_ALL}\n")
        
        for i, option in enumerate(self.options):
            if i == self.selected:
                print(f"{Back.BLUE}{Fore.WHITE}  → {option}  {Style.RESET_ALL}")
            else:
                print(f"    {option}")
        print()
    
    def handle_input(self):
        """处理用户输入"""
        try:
            key = getch.getch()
            
            if key == '\x1b':  # ESC序列开始
                key += getch.getch()
                if key == '\x1b[':
                    key += getch.getch()
                    if key == '\x1b[A':  # 上箭头
                        self.selected = (self.selected - 1) % len(self.options)
                    elif key == '\x1b[B':  # 下箭头
                        self.selected = (self.selected + 1) % len(self.options)
            elif key == '\r' or key == '\n':  # 回车
                return self.execute_option()
            elif key == '0':  # 0键退出
                return False
            elif key == 'q' or key == 'Q':  # q键退出
                return False
            
        except KeyboardInterrupt:
            return False
        
        return True
    
    def execute_option(self):
        """执行选中的选项"""
        if self.selected == 0:  # 处理Excel文件
            os.system('clear' if os.name == 'posix' else 'cls')
            self.processor.display_banner()
            
            if self.processor.get_file_input():
                if self.processor.process_excel():
                    print(f"\n{Fore.GREEN}按任意键返回主菜单...{Style.RESET_ALL}")
                    getch.getch()
                else:
                    print(f"\n{Fore.RED}按任意键返回主菜单...{Style.RESET_ALL}")
                    getch.getch()
            else:
                print(f"\n{Fore.RED}按任意键返回主菜单...{Style.RESET_ALL}")
                getch.getch()
                
        elif self.selected == 1:  # 查看帮助
            self.show_help()
            
        elif self.selected == 2:  # 退出程序
            return False
            
        return True
    
    def show_help(self):
        """显示帮助信息"""
        os.system('clear' if os.name == 'posix' else 'cls')
        self.processor.display_banner()
        
        print(f"{Fore.CYAN}📖 使用帮助{Style.RESET_ALL}\n")
        print(f"{Fore.YELLOW}功能说明：{Style.RESET_ALL}")
        print("• 本工具用于处理Excel表格数据，提取商品编码和采购价信息")
        print("• 自动计算成本价和基本售价")
        print("• 生成标准格式的新Excel文件\n")
        
        print(f"{Fore.YELLOW}使用步骤：{Style.RESET_ALL}")
        print("1. 选择「处理Excel文件」")
        print("2. 拖拽Excel文件到终端窗口")
        print("3. 程序自动处理并生成结果文件\n")
        
        print(f"{Fore.YELLOW}计算规则：{Style.RESET_ALL}")
        print("• 成本价 = 采购价 + 2")
        print("• 基本售价 = ((采购价 + 5) × 2)的整数部分 + 0.99")
        print("• 虚拟分类 = 可预售\n")
        
        print(f"{Fore.YELLOW}支持格式：{Style.RESET_ALL}")
        print("• 输入：.xlsx、.xls格式的Excel文件")
        print("• 输出：.xlsx格式的Excel文件\n")
        
        print(f"{Fore.GREEN}按任意键返回主菜单...{Style.RESET_ALL}")
        getch.getch()
    
    def run(self):
        """运行主循环"""
        while True:
            self.display_menu()
            if not self.handle_input():
                break
        
        print(f"\n{Fore.CYAN}感谢使用预处理工具！{Style.RESET_ALL}")

def main():
    """主函数"""
    try:
        menu = InteractiveMenu()
        menu.run()
    except KeyboardInterrupt:
        print(f"\n\n{Fore.CYAN}程序已退出{Style.RESET_ALL}")
    except Exception as e:
        print(f"\n{Fore.RED}程序运行出错：{str(e)}{Style.RESET_ALL}")

if __name__ == "__main__":
    main()