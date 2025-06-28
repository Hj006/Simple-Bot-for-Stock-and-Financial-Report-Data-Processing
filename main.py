
import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import re


def load_and_process_data():
    from processor.process_excel import process_excel_file

    print("\n📁 请选择包含财务数据文件的文件夹...")

    try:
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="选择数据集文件夹")
        if not folder_path:
            print("未选择任何文件夹，操作已取消。")
            return

        print(f"你选择的文件夹是：{folder_path}")

        def contains_non_ascii(s):
            return not all(ord(c) < 128 for c in s)

        if contains_non_ascii(folder_path):
            print("选中的路径包含中文或特殊字符，建议使用英文路径。")
            return

        # 找到所有 .xlsx 文件（暂不支持 .csv）
        files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xlsx')]

        if not files:
            print("该文件夹中未找到 .xlsx 文件。")
            return

        print(f"\n找到以下文件：")
        for f in files:
            print("   -", f)

        # 逐个处理
        for f in files:
            full_path = os.path.join(folder_path, f)
            output_path = os.path.join(folder_path, f.replace('.xlsx', '_v1.xlsx'))
            print(f"\n正在处理: {f}")
            try:
                process_excel_file(full_path, output_path)
            except Exception as e:
                print(f"处理失败: {f}，原因：{e}")

        print("\n所有处理完成。")

    except Exception as e:
        print("出现错误：", str(e))



def filter_by_indicators():
    pass
def query_data():
    pass
def define_custom_indicator():
    pass

def menu():
    while True:
        print("\n Simple Financial Data Bot - 主菜单")
        print("1. 读取原始数据并计算保存指标")
        print("2. 设置筛选条件，输出满足条件的公司及财报时间")
        print("3. 查询某个公司或指标的值")
        print("4. 自定义新指标")
        print("5. 退出程序")
        
        choice = input("请输入选项 (1-5): ").strip()

        if choice == "1":
            load_and_process_data()
        elif choice == "2":
            filter_by_indicators()
        elif choice == "3":
            query_data()
        elif choice == "4":
            define_custom_indicator()
        elif choice == "5":
            print("👋 程序已退出。")
            break
        else:
            print("无效输入，请输入 1-5 之间的数字。")

if __name__ == "__main__":
    menu()
